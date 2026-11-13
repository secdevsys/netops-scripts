#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests de hpe5940_parser.py

Lancer (Windows) :  py -m unittest test_hpe5940_parser -v
Lancer (Linux)   :  python3 -m unittest test_hpe5940_parser -v

Les tests utilisent sample_hpe5940.cfg s'il est présent dans le même dossier,
et génèrent leurs propres configs de test dans un dossier temporaire.
"""

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import hpe5940_parser as P  # noqa: E402

if getattr(P, "__version__", "0") < "1.9.0":
    sys.exit(f"ERREUR : {P.__file__} est une ancienne version du script "
             f"({getattr(P, '__version__', 'sans numéro')}). Remplacez-le par la version 1.9.0 ou plus.")

SAMPLE = HERE / "sample_hpe5940.cfg"


def rows(xlsx, sheet):
    """Retourne les lignes d'un onglet sous forme de dicts {en-tête: valeur}."""
    ws = load_workbook(xlsx)[sheet]
    data = list(ws.iter_rows(values_only=True))
    head = data[0]
    return [dict(zip(head, r)) for r in data[1:]]


def by_name(lst, name, key="Interface"):
    return next(r for r in lst if r[key] == name)


def all_cells(xlsx):
    wb = load_workbook(xlsx)
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                yield ws.title, c


class TmpDirTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="hpe5940_test_"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write_cfg(self, name, text, encoding="utf-8"):
        p = self.tmp / name
        p.write_bytes(text.encode(encoding) if isinstance(text, str) else text)
        return p

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(HERE / "hpe5940_parser.py"), *map(str, args)],
                              capture_output=True, text=True, encoding="utf-8", errors="replace")


# --------------------------------------------------------------------------
# Fonctions utilitaires
# --------------------------------------------------------------------------
class TestHelpers(unittest.TestCase):
    def test_vlan_tokens(self):
        self.assertEqual(P.parse_vlan_tokens("10 20 30 to 33".split()), {10, 20, 30, 31, 32, 33})
        self.assertEqual(len(P.parse_vlan_tokens(["all"])), 4094)

    def test_fmt_vlans(self):
        self.assertEqual(P.fmt_vlans({1, 2, 3, 10, 12, 13}), "1-3,10,12-13")
        self.assertEqual(P.fmt_vlans(set()), "")
        self.assertEqual(P.fmt_vlans(set(range(1, 4095))), "all (1-4094)")

    def test_ip_info(self):
        self.assertEqual(P.ip_info("10.20.0.1", "24"),
                         ("10.20.0.1", "255.255.255.0", "/24", "10.20.0.0/24"))
        self.assertEqual(P.ip_info("10.1.1.1", "255.255.255.252")[3], "10.1.1.0/30")

    def test_wildcard(self):
        self.assertEqual(P.wildcard_to_prefix("172.16.0.0", "0.0.0.7"), "172.16.0.0/29")
        self.assertEqual(P.wildcard_to_prefix("10.255.0.1", "0.0.0.0"), "10.255.0.1/32")

    def test_mask_secrets(self):
        for line, secret in [("peer 1.1.1.1 password cipher $c$3$abc==", "$c$3$abc=="),
                             ("ospf authentication-mode simple plain MyS3cret", "MyS3cret"),
                             ("authentication-mode md5 1 plain OspfPwd", "OspfPwd"),
                             ("password hash $h$6$xyz", "$h$6$xyz")]:
            self.assertNotIn(secret, P.mask(line), line)
            self.assertIn("******", P.mask(line))

    def test_static_route(self):
        r = P.parse_static("ip route-static 10.201.0.0 255.255.0.0 Vlan-interface999 172.16.0.2 tag 100", "X")
        self.assertEqual((r["dest"], r["out_if"], r["nh"], r["tag"]),
                         ("10.201.0.0/16", "Vlan-interface999", "172.16.0.2", "100"))
        r = P.parse_static("ip route-static vpn-instance A 0.0.0.0 0 192.168.255.2 description Def route", "X")
        self.assertEqual((r["vrf"], r["dest"], r["nh"], r["desc"]), ("A", "0.0.0.0/0", "192.168.255.2", "Def route"))


# --------------------------------------------------------------------------
# Régression sur la config d'exemple
# --------------------------------------------------------------------------
@unittest.skipUnless(SAMPLE.exists(), "sample_hpe5940.cfg absent")
class TestSample(TmpDirTest):
    @classmethod
    def setUpClass(cls):
        cls.dir = Path(tempfile.mkdtemp(prefix="hpe5940_sample_"))
        res = P.parse_config(SAMPLE)
        cls.xlsx = cls.dir / "r.xlsx"
        P.write_report([res], cls.xlsx)
        cls.l2 = rows(cls.xlsx, "Interfaces L2")
        cls.l3 = rows(cls.xlsx, "Interfaces L3")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.dir, ignore_errors=True)

    def test_sheets(self):
        self.assertEqual(load_workbook(self.xlsx).sheetnames,
                         ["Synthèse", "Interfaces L2", "Interfaces L3", "Routage dynamique",
                          "Voisins BGP", "Routes statiques", "VLANs", "Plan d'adressage", "Réadressage"])

    def test_trunk(self):
        r = by_name(self.l2, "Bridge-Aggregation1")
        self.assertEqual((r["Mode"], r["VLANs autorisées (trunk)"], r["Nb VLANs trunk"]),
                         ("trunk", "10,20,30-35,100", 9))           # undo permit vlan 1
        self.assertEqual(r["Membres (agrégat)"], "Ten-GigabitEthernet1/0/1, Ten-GigabitEthernet1/0/2")

    def test_member_inherits(self):
        r = by_name(self.l2, "Ten-GigabitEthernet1/0/2")
        self.assertEqual(r["Membre de l'agrégat"], "Bridge-Aggregation1")
        self.assertEqual(r["VLANs autorisées (trunk)"], "10,20,30-35,100")
        self.assertTrue(r["Config héritée"].startswith("oui"))

    def test_access_hybrid_shutdown(self):
        self.assertEqual(by_name(self.l2, "Ten-GigabitEthernet1/0/4")["VLAN access"], "10")
        h = by_name(self.l2, "Ten-GigabitEthernet1/0/5")
        self.assertEqual((h["Mode"], h["Hybrid tagged"], h["Hybrid untagged"], h["PVID"]),
                         ("hybrid", "20", "10", "10"))
        self.assertEqual(by_name(self.l2, "Ten-GigabitEthernet1/0/6")["Statut admin"], "shutdown")

    def test_l3(self):
        v10 = by_name(self.l3, "Vlan-interface10")
        self.assertEqual((v10["Adresse IP"], v10["Masque"], v10["IP secondaires"]),
                         ("10.10.10.1", "255.255.255.0", "10.10.11.1/24"))
        self.assertEqual(v10["Serveurs DHCP relay"], "10.100.0.10, 10.100.0.11")
        self.assertIn("10.10.10.254", v10["VRRP"])
        self.assertEqual(by_name(self.l3, "Vlan-interface20")["IP MTU"], "1400")
        self.assertEqual(by_name(self.l3, "Route-Aggregation10")["VRF"], "CLIENT-A")
        self.assertEqual(by_name(self.l3, "HundredGigE1/0/53.100")["dot1q"], "100")
        names = {r["Interface"] for r in self.l3} | {r["Interface"] for r in self.l2}
        self.assertNotIn("NULL0", names)
        self.assertIn("Ten-GigabitEthernet1/0/47", {r["Interface"] for r in self.l3})  # link-mode route

    def test_routing(self):
        rt = rows(self.xlsx, "Routage dynamique")
        nets = {r["Valeur"] for r in rt if r["Catégorie"] == "Network" and r["Protocole"] == "OSPF"}
        self.assertTrue({"10.255.0.1/32", "172.16.0.0/29", "10.10.10.0/24", "192.168.255.0/30"} <= nets)
        self.assertTrue(any(r["Catégorie"] == "Type de zone" and r["Valeur"] == "stub" for r in rt))
        bgp_rows = [r for r in rt if r["Protocole"] == "BGP"]
        self.assertTrue(all(r["Router-ID"] == "10.255.0.1" for r in bgp_rows))

    def test_bgp(self):
        b = rows(self.xlsx, "Voisins BGP")
        ibgp = by_name(b, "10.255.0.2", "Voisin / Groupe")
        self.assertEqual((ibgp["Session"], ibgp["AS distant"], ibgp["Connect-interface"]),
                         ("iBGP", "65001", "LoopBack0"))           # hérité du peer-group
        isp = by_name(b, "203.0.113.1", "Voisin / Groupe")
        self.assertEqual((isp["Session"], isp["Auth MD5"], isp["BFD"], isp["Route-policy import"]),
                         ("eBGP", "oui", "oui", "RP-ISP-IN"))
        self.assertEqual(by_name(b, "192.168.255.2", "Voisin / Groupe")["VRF"], "CLIENT-A")

    def test_statics(self):
        st = rows(self.xlsx, "Routes statiques")
        self.assertEqual(len(st), 6)
        self.assertEqual({r["VRF"] for r in st}, {"public", "CLIENT-A", "mgmt"})

    def test_vlans_usage_no_members(self):
        v10 = by_name(rows(self.xlsx, "VLANs"), 10, "VLAN")
        self.assertNotIn("Ten-GigabitEthernet1/0/1", v10["Ports trunk/hybrid"])  # représenté par BAGG1
        self.assertIn("Bridge-Aggregation1", v10["Ports trunk/hybrid"])

    def test_no_secret_in_report(self):
        for _, c in all_cells(self.xlsx):
            self.assertNotIn("$c$3$", str(c.value or ""))


# --------------------------------------------------------------------------
# Cas limites
# --------------------------------------------------------------------------
EDGE = """<EDGE>display current-configuration
#
 sysname EDGE-SW
#
interface Ten-GigabitEthernet1/0/1
 port link-mode bridge
 description === UPLINK CORE ===
 port access vlan 10
#
interface Ten-GigabitEthernet1/0/2
 port link-mode bridge
 description bell\x07char\x00
#
interface Vlan-interface50
 ip address dhcp-alloc client-identifier mac Vlan-interface50
 ospf authentication-mode simple plain MyS3cret
#
ospf 1
 area 0.0.0.0
  authentication-mode md5 1 plain OspfPwd
#
bgp 65001
 router-id 9.9.9.9
 ip vpn-instance A
  peer 1.1.1.1 as-number 100
 #
 address-family ipv4 unicast
  peer 2.2.2.2 enable
 peer 2.2.2.2 as-number 200
#
return
"""


class TestEdgeCases(TmpDirTest):
    def report(self, text, name="edge.cfg", encoding="utf-8"):
        cfg = self.write_cfg(name, text, encoding)
        res = P.parse_config(cfg)
        out = self.tmp / "out.xlsx"
        P.write_report([res], out)
        return out

    def test_control_chars_do_not_crash(self):
        out = self.report(EDGE)
        d = by_name(rows(out, "Interfaces L2"), "Ten-GigabitEthernet1/0/2")["Description"]
        self.assertEqual(d, "bellchar")

    def test_equal_sign_is_text_not_formula(self):
        out = self.report(EDGE)
        for _, c in all_cells(out):
            self.assertNotEqual(c.data_type, "f", f"formule inattendue : {c.value}")
        self.assertEqual(by_name(rows(out, "Interfaces L2"), "Ten-GigabitEthernet1/0/1")["Description"],
                         "=== UPLINK CORE ===")

    def test_dhcp_alloc(self):
        r = by_name(rows(self.report(EDGE), "Interfaces L3"), "Vlan-interface50")
        self.assertEqual(r["Adresse IP"], "dhcp-alloc client-identifier")
        self.assertIn(r["Masque"], (None, ""))

    def test_plain_passwords_masked(self):
        out = self.report(EDGE)
        for _, c in all_cells(out):
            for secret in ("MyS3cret", "OspfPwd"):
                self.assertNotIn(secret, str(c.value or ""))

    def test_bgp_public_af_after_vrf(self):
        b = rows(self.report(EDGE), "Voisins BGP")
        p2 = by_name(b, "2.2.2.2", "Voisin / Groupe")
        self.assertEqual((p2["VRF"], p2["AS distant"], p2["Session"]), ("public", "200", "eBGP"))
        self.assertEqual(by_name(b, "1.1.1.1", "Voisin / Groupe")["VRF"], "A")

    def test_encodings(self):
        text = EDGE.replace("=== UPLINK CORE ===", "Imprimante étage")
        for enc, name in (("utf-8-sig", "bom.cfg"), ("cp1252", "cp.cfg"), ("utf-16", "u16.cfg")):
            out = self.report(text, name, enc)
            d = by_name(rows(out, "Interfaces L2"), "Ten-GigabitEthernet1/0/1")["Description"]
            self.assertEqual(d, "Imprimante étage", enc)

    def test_more_prompt_and_crlf(self):
        text = EDGE.replace("\n", "\r\n").replace(" port access vlan 10",
                                                   "  ---- More ----\x1b[42D \x1b[42D port access vlan 10")
        r = by_name(rows(self.report(text), "Interfaces L2"), "Ten-GigabitEthernet1/0/1")
        self.assertEqual(r["VLAN access"], "10")

    def test_real_putty_more_erase_sequence(self):
        # Comware efface le prompt avec ESC[16D + 16 espaces + ESC[16D : l'indentation
        # résiduelle ne doit pas rattacher l'interface au bloc précédent
        erase = "  ---- More ----\x1b[16D                \x1b[16D"
        text = EDGE.replace("interface Vlan-interface50", erase + "interface Vlan-interface50")
        text = text.replace(" port access vlan 10", erase + " port access vlan 10")
        out = self.report(text)
        self.assertEqual(by_name(rows(out, "Interfaces L3"), "Vlan-interface50")["Adresse IP"],
                         "dhcp-alloc client-identifier")
        self.assertEqual(by_name(rows(out, "Interfaces L2"), "Ten-GigabitEthernet1/0/1")["VLAN access"], "10")

    def test_empty_file(self):
        out = self.report("")
        self.assertTrue(out.exists())
        self.assertEqual(rows(out, "Interfaces L2"), [])


# --------------------------------------------------------------------------
# Ligne de commande
# --------------------------------------------------------------------------
class TestCLI(TmpDirTest):
    def test_default_output_next_to_config(self):
        sub = self.tmp / "site A"                  # dossier avec espace, lancé depuis ailleurs
        sub.mkdir()
        cfg = sub / "a.txt"
        cfg.write_text(EDGE, encoding="utf-8")
        r = self.run_cli(cfg)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(sorted(p.name for p in sub.glob("*.xlsx")), ["EDGE-SW.xlsx"], r.stdout)

    def test_folder_glob_merge_and_duplicate_sysname(self):
        self.write_cfg("a.cfg", EDGE)
        self.write_cfg("b.txt", EDGE)                     # même sysname
        self.write_cfg("ignore.xlsx", b"PK")              # extension non retenue
        out = self.tmp / "rapports"
        r = self.run_cli(self.tmp, str(self.tmp / "*.cfg"), "-d", out, "--merge", "tous.xlsx")
        self.assertEqual(r.returncode, 0, r.stderr)
        produced = sorted(p.name for p in out.glob("*.xlsx"))
        self.assertEqual(produced, ["EDGE-SW.xlsx", "EDGE-SW_b.xlsx", "tous.xlsx"],
                         r.stdout + r.stderr)
        self.assertIn("2 rapport(s)", r.stdout)            # a.cfg pas compté deux fois

    def test_missing_file(self):
        r = self.run_cli(self.tmp / "nope.cfg")
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("introuvable", r.stderr)

    def test_debug(self):
        cfg = self.write_cfg("a.cfg", EDGE)
        r = self.run_cli(cfg, "--debug")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("[DEBUG] Fichier", r.stdout)
        self.assertIn("existe : True", r.stdout)


# --------------------------------------------------------------------------
# Parcours récursif (arbre de nœuds)
# --------------------------------------------------------------------------
def cfg_for(sysname):
    return EDGE.replace("EDGE-SW", sysname)


class TestRecursive(TmpDirTest):
    def node(self, rel, sysname, filename=None):
        d = self.tmp / "root" / rel
        d.mkdir(parents=True, exist_ok=True)
        (d / (filename or f"{sysname.lower()}.cfg")).write_text(cfg_for(sysname), encoding="utf-8")
        return d

    def build_tree(self):
        #  CORE
        #  ├─ DIST-A ─┬─ ACC-01   (dossier "ACC 01" avec espace)
        #  │          └─ ACC-03
        #  └─ DIST-B ── "Batiment C" (sans config) ── ACC-02
        self.node("", "CORE")
        self.node("DIST-A", "DIST-A", "dist-a.txt")
        self.node("DIST-B", "DIST-B")
        self.node("DIST-A/ACC 01", "ACC-01")
        self.node("DIST-A/ACC-03", "ACC-03")
        self.node("DIST-B/Batiment C/ACC-02", "ACC-02", "acc02.txt")
        (self.tmp / "root" / "notes.txt").write_text("pas une config", encoding="utf-8")
        (self.tmp / "root" / ".git").mkdir()
        (self.tmp / "root" / ".git" / "x.cfg").write_text(cfg_for("HIDDEN"), encoding="utf-8")
        return self.tmp / "root"

    def reports(self):
        return sorted(str(p.relative_to(self.tmp / "root")).replace("\\", "/")
                      for p in (self.tmp / "root").rglob("*.xlsx"))

    def syn(self, xlsx, dev):
        return by_name(rows(xlsx, "Synthèse"), dev, "Équipement")

    def test_without_r_only_top_level(self):
        root = self.build_tree()
        r = self.run_cli(root)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(self.reports(), ["CORE.xlsx"])
        self.assertIn("[IGNORÉ]", r.stdout)                 # notes.txt

    def test_each_report_in_its_folder(self):
        root = self.build_tree()
        r = self.run_cli(root, "-r", "--sans-consolide")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(self.reports(), ["CORE.xlsx", "DIST-A/ACC 01/ACC-01.xlsx",
                                          "DIST-A/ACC-03/ACC-03.xlsx", "DIST-A/DIST-A.xlsx",
                                          "DIST-B/Batiment C/ACC-02/ACC-02.xlsx", "DIST-B/DIST-B.xlsx"])
        self.assertIn("1 fichier(s) ignoré(s)", r.stdout)   # notes.txt ; .git ignoré

    def test_parent_children_level(self):
        root = self.build_tree()
        self.run_cli(root, "-r")
        core = self.syn(root / "CORE.xlsx", "CORE")
        self.assertEqual((core["Nœud parent"], core["Nœud(s) enfant(s)"], core["Niveau"]),
                         (None, "DIST-A, DIST-B", 0))
        a = self.syn(root / "DIST-A" / "DIST-A.xlsx", "DIST-A")
        self.assertEqual((a["Nœud parent"], a["Nœud(s) enfant(s)"], a["Niveau"]),
                         ("CORE", "ACC-01, ACC-03", 1))
        # dossier intermédiaire sans config : ACC-02 rattaché à DIST-B, niveau 2
        acc2 = self.syn(root / "DIST-B" / "Batiment C" / "ACC-02" / "ACC-02.xlsx", "ACC-02")
        self.assertEqual((acc2["Nœud parent"], acc2["Niveau"]), ("DIST-B", 2))

    def test_merge_tree_sheet_in_tree_order(self):
        root = self.build_tree()
        r = self.run_cli(root, "-r", "--merge", "parc.xlsx")
        self.assertEqual(r.returncode, 0, r.stderr)
        tree = [(t["Niveau"], t["Nœud"], t["Nœud parent"]) for t in rows(root / "parc.xlsx", "Arborescence")]
        self.assertEqual(tree, [(0, "CORE", None), (1, "DIST-A", "CORE"), (2, "ACC-01", "DIST-A"),
                                (2, "ACC-03", "DIST-A"), (1, "DIST-B", "CORE"), (2, "ACC-02", "DIST-B")])

    def test_hierarchy_columns_in_detail_sheets(self):
        root = self.build_tree()
        r = self.run_cli(root, "-r", "--merge", "parc.xlsx")
        self.assertEqual(r.returncode, 0, r.stderr)
        merged = root / "parc.xlsx"
        for sheet in ("Interfaces L2", "Interfaces L3", "Routage dynamique", "Voisins BGP",
                      "Routes statiques", "VLANs"):
            head = next(load_workbook(merged)[sheet].iter_rows(max_row=1, values_only=True))
            self.assertEqual(head[:4], ("Niveau 0", "Niveau 1", "Niveau 2", "Équipement"), sheet)
        l2 = rows(merged, "Interfaces L2")
        acc2 = next(x for x in l2 if x["Équipement"] == "ACC-02")
        self.assertEqual((acc2["Niveau 0"], acc2["Niveau 1"], acc2["Niveau 2"]), ("CORE", "DIST-B", "ACC-02"))
        dista = next(x for x in l2 if x["Équipement"] == "DIST-A")
        self.assertEqual((dista["Niveau 0"], dista["Niveau 1"], dista["Niveau 2"]), ("CORE", "DIST-A", None))
        # lignes dans l'ordre de l'arbre : chaque équipement apparaît en un seul bloc, parent d'abord
        order = []
        for x in l2:
            if not order or order[-1] != x["Équipement"]:
                order.append(x["Équipement"])
        self.assertEqual(order, ["CORE", "DIST-A", "ACC-01", "ACC-03", "DIST-B", "ACC-02"])
        # rapport individuel : chemin de l'équipement uniquement
        ind = rows(root / "DIST-A" / "ACC 01" / "ACC-01.xlsx", "Interfaces L2")
        self.assertEqual((ind[0]["Niveau 0"], ind[0]["Niveau 1"], ind[0]["Niveau 2"]), ("CORE", "DIST-A", "ACC-01"))

    def test_no_hierarchy_without_r(self):
        cfg = self.write_cfg("a.cfg", EDGE)
        self.run_cli(cfg)
        head = next(load_workbook(self.tmp / "EDGE-SW.xlsx")["Interfaces L2"].iter_rows(max_row=1, values_only=True))
        self.assertEqual(head[0], "Équipement")

    def test_consolidated_in_given_folder_even_if_root_has_no_config(self):
        # racine sans config : Reseau/ -> CORE/ -> DIST/ ...
        top = self.tmp / "Reseau"
        (top / "CORE").mkdir(parents=True)
        (top / "CORE" / "core.cfg").write_text(cfg_for("CORE"), encoding="utf-8")
        (top / "CORE" / "DIST").mkdir()
        (top / "CORE" / "DIST" / "d.txt").write_text(cfg_for("DIST"), encoding="utf-8")
        r = self.run_cli(top, "-r", "--merge", "parc.xlsx")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertTrue((top / "parc.xlsx").exists(), r.stdout)
        self.assertIn(f"Fichier consolidé : {top.resolve() / 'parc.xlsx'}", r.stdout)

    def test_r_alone_produces_hierarchical_file(self):
        root = self.build_tree()
        r = self.run_cli(root, "-r")
        self.assertEqual(r.returncode, 0, r.stderr)
        f = root / "parc_hierarchique.xlsx"
        self.assertTrue(f.exists(), r.stdout)
        self.assertEqual(rows(f, "Interfaces L2")[0]["Niveau 0"], "CORE")
        r = self.run_cli(root, "-r", "--sans-consolide")
        self.assertNotIn("consolidé", r.stdout)

    def test_links_not_followed(self):
        root = self.build_tree()
        try:
            os.symlink(root / "DIST-A" / "ACC 01", root / "DIST-B" / "lien", target_is_directory=True)
            os.symlink(root, root / "DIST-A" / "ACC-03" / "boucle", target_is_directory=True)
        except (OSError, NotImplementedError) as e:
            self.skipTest(f"liens symboliques non autorisés ici : {e}")
        r = self.run_cli(root, "-r")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(r.stderr.count("non suivi"), 2)
        self.assertIn("6 rapport(s)", r.stdout)
        self.assertNotIn("DIST-B/lien/ACC-01.xlsx", self.reports())


# --------------------------------------------------------------------------
# Intégration avec generer_arborescence.py (si présent)
# --------------------------------------------------------------------------
GEN = HERE / "generer_arborescence.py"


@unittest.skipUnless(GEN.exists(), "generer_arborescence.py absent")
class TestGenerator(TmpDirTest):
    def run_gen(self, *args):
        return subprocess.run([sys.executable, str(GEN), *map(str, args)],
                              capture_output=True, text=True, encoding="utf-8", errors="replace")

    def test_generated_tree_matches_parser(self):
        import generer_arborescence as G
        out = self.tmp / "arbre"
        g = self.run_gen(out, "-n", "4", "-f", "2,3,2", "--pieges")
        self.assertEqual(g.returncode, 0, g.stderr)
        root = G.build_tree(4, [2, 3, 2])                   # même arbre que le script
        expected = {n.name: (n.parent.name if n.parent else None) for n in root.walk()}
        r = self.run_cli(out, "-r", "-d", self.tmp / "rapports", "--merge", "parc.xlsx")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn(f"{len(expected)} rapport(s)", r.stdout)
        self.assertIn("1 fichier(s) ignoré(s)", r.stdout)                  # LISEZMOI.txt
        tree = rows(self.tmp / "rapports" / "parc.xlsx", "Arborescence")
        self.assertEqual({t["Nœud"]: t["Nœud parent"] for t in tree}, expected)
        # chaque lien montant décrit dans une config pointe vers le bon parent
        l2 = rows(self.tmp / "rapports" / "parc.xlsx", "Interfaces L2")
        l3 = rows(self.tmp / "rapports" / "parc.xlsx", "Interfaces L3")
        for row in l2 + l3:
            desc = row["Description"] or ""
            if desc.startswith("UPLINK vers "):
                self.assertEqual(desc.split()[2], expected[row["Équipement"]], desc)
        # pièges : pas d'IP en double, pas de secret, pas de formule
        ips = [x["Adresse IP"] for x in l3 if x["Adresse IP"]]
        self.assertEqual(len(ips), len(set(ips)))
        for _, c in all_cells(self.tmp / "rapports" / "parc.xlsx"):
            self.assertNotEqual(c.data_type, "f")
            self.assertNotIn("$c$3$", str(c.value or ""))

    def test_refuses_non_empty_folder(self):
        out = self.tmp / "arbre"
        out.mkdir()
        (out / "important.txt").write_text("x", encoding="utf-8")
        g = self.run_gen(out)
        self.assertNotEqual(g.returncode, 0)
        self.assertTrue((out / "important.txt").exists())


# --------------------------------------------------------------------------
# Classification des subnets (Plan d'adressage / Réadressage)
# --------------------------------------------------------------------------
def mini_cfg(sysname, body):
    return f"#\n sysname {sysname}\n#\n{body}\n#\nreturn\n"


ACCESS_PORTS = """interface GigabitEthernet1/0/1
 port access vlan {v}
#
interface GigabitEthernet1/0/2
 port access vlan {v}
#
interface Ten-GigabitEthernet1/0/49
 port link-type trunk
 port trunk permit vlan all"""

TRUNK_ONLY = """interface Ten-GigabitEthernet1/0/1
 port link-type trunk
 port trunk permit vlan all
#
interface Ten-GigabitEthernet1/0/49
 port link-type trunk
 port trunk permit vlan all"""


class TestSubnetClassification(TmpDirTest):
    def build(self):
        nodes = {
            # CORE : LAN serveurs (VLAN 60) dont les ports d'accès sont 2 niveaux plus bas,
            #        transit /29 partagé avec les 2 DIST, SVI /28 sans aucun indice
            "": ("CORE", """vlan 60
 name SRV
#
interface Vlan-interface60
 ip address 10.60.0.1 255.255.255.0
#
interface Vlan-interface900
 ip address 172.31.0.1 255.255.255.248
 ospf 1 area 0.0.0.0
#
interface Vlan-interface77
 ip address 10.77.0.1 255.255.255.240
#
interface Ten-GigabitEthernet1/0/10
 port link-mode route
 ip binding vpn-instance A
 ip address 10.1.1.1 255.255.255.252
#
interface Ten-GigabitEthernet1/0/11
 port link-mode route
 ip binding vpn-instance B
 ip address 10.1.1.1 255.255.255.0
 dhcp select relay
 dhcp relay server-address 10.1.1.200
""" + TRUNK_ONLY),
            # DIST-A / DIST-B : paire VRRP sur le même LAN (VLAN 50)
            "DIST-A": ("DIST-A", """interface Vlan-interface50
 ip address 10.50.0.2 255.255.255.0
 vrrp vrid 50 virtual-ip 10.50.0.1
#
interface Vlan-interface900
 ip address 172.31.0.2 255.255.255.248
 ospf 1 area 0.0.0.0
""" + TRUNK_ONLY),
            "DIST-B": ("DIST-B", """interface Vlan-interface50
 ip address 10.50.0.3 255.255.255.0
 vrrp vrid 50 virtual-ip 10.50.0.1
#
interface Vlan-interface900
 ip address 172.31.0.3 255.255.255.248
 ospf 1 area 0.0.0.0
""" + TRUNK_ONLY),
            # ACC-A1 : switch d'accès L3 (SVI + DHCP relay) ; ports du VLAN 60 du CORE
            "DIST-A/ACC-A1": ("ACC-A1", """interface Vlan-interface70
 ip address 10.70.0.1 255.255.255.0
 dhcp select relay
 dhcp relay server-address 10.60.0.10
#
""" + ACCESS_PORTS.format(v=60)),
            "DIST-B/ACC-B1": ("ACC-B1", ACCESS_PORTS.format(v=50)),
        }
        for rel, (name, body) in nodes.items():
            d = self.tmp / "root" / rel
            d.mkdir(parents=True, exist_ok=True)
            (d / f"{name}.cfg").write_text(mini_cfg(name, body), encoding="utf-8")
        r = self.run_cli(self.tmp / "root", "-r")
        self.assertEqual(r.returncode, 0, r.stderr)
        plan = rows(self.tmp / "root" / "parc_hierarchique.xlsx", "Plan d'adressage")
        return {(p["VRF"], p["Subnet"]): p for p in plan}

    def test_categories(self):
        plan = self.build()
        expect = {
            ("public", "10.60.0.0/24"): "LAN",        # SVI CORE, ports d'accès sur ACC-A1 (niveau 2)
            ("public", "10.50.0.0/24"): "LAN",        # paire VRRP : 2 équipements mais LAN
            ("public", "10.70.0.0/24"): "LAN",        # LAN porté par un switch d'accès
            ("public", "172.31.0.0/29"): "INTERCO",   # transit partagé par 3 équipements, OSPF actif
            ("public", "10.77.0.0/28"): "À VÉRIFIER", # aucun indice
            ("A", "10.1.1.0/30"): "INTERCO",          # même IP 10.1.1.1 dans 2 VRF différentes
            ("B", "10.1.1.0/24"): "LAN",
        }
        got = {k: plan[k]["Catégorie"] for k in expect}
        self.assertEqual(got, expect, {k: plan[k]["Raisons"] for k in expect if got[k] != expect[k]})
        self.assertIn("ACC-A1", plan[("public", "10.60.0.0/24")]["Raisons"])
        self.assertIn("qualifier manuellement", plan[("public", "10.77.0.0/28")]["Raisons"])
        self.assertIn("10.50.0.1", plan[("public", "10.50.0.0/24")]["VIP VRRP"])
        self.assertIn("serveur DHCP relay 10.60.0.10", plan[("public", "10.60.0.0/24")]["Références à mettre à jour"])

    def test_readdressing_sheet(self):
        self.build()
        rr = rows(self.tmp / "root" / "parc_hierarchique.xlsx", "Réadressage")
        vrrp = [r for r in rr if r["Subnet"] == "10.50.0.0/24"]
        self.assertEqual(sorted((r["Équipement"], r["Type d'adresse"], r["IP actuelle"]) for r in vrrp),
                         [("DIST-A", "VIP VRRP 50", "10.50.0.1"), ("DIST-A", "primaire", "10.50.0.2"),
                          ("DIST-B", "VIP VRRP 50", "10.50.0.1"), ("DIST-B", "primaire", "10.50.0.3")])
        self.assertTrue(all(r["Nouvelle IP"] is None and r["Catégorie"] for r in rr))
        # colonnes de hiérarchie présentes
        self.assertEqual((vrrp[0]["Niveau 0"], vrrp[0]["Niveau 1"]), ("CORE", "DIST-A"))

    def test_l3_sheet_has_category(self):
        self.build()
        l3 = rows(self.tmp / "root" / "parc_hierarchique.xlsx", "Interfaces L3")
        v900 = [x["Catégorie subnet"] for x in l3 if x["Interface"] == "Vlan-interface900"]
        self.assertEqual(v900, ["INTERCO"] * 3)

    def test_individual_report_uses_whole_park(self):
        self.build()
        # dans le rapport individuel de DIST-A, le transit /29 reste INTERCO grâce aux autres équipements
        plan = rows(self.tmp / "root" / "DIST-A" / "DIST-A.xlsx", "Plan d'adressage")
        p = next(x for x in plan if x["Subnet"] == "172.31.0.0/29")
        self.assertEqual((p["Catégorie"], p["Nb équipements"]), ("INTERCO", 3))


# --------------------------------------------------------------------------
# ArubaOS-CX (lecteur_aoscx.py) et parc mixte Comware + AOS-CX
# --------------------------------------------------------------------------
DIST_CX = """DIST-CX-01# show running-config
Current configuration:
!
!Version ArubaOS-CX PL.10.13.1190
!export-password: default
hostname DIST-CX-01
user admin group administrators password ciphertext AQBapSECRETSECRET
vrf PROD
    rd 65001:100
ntp server 10.100.0.5
!
vlan 1
vlan 10
    name USERS
vlan 20
    name VOIP
vlan 30-31
vlan 999
    name MGMT
spanning-tree
interface mgmt
    no shutdown
    ip static 192.168.10.11/24
    default-gateway 192.168.10.1
interface loopback 0
    ip address 10.255.1.11/32
    ip ospf 1 area 0.0.0.0
interface lag 1 multi-chassis
    description DOWNLINK vers ACC-CX-01
    no shutdown
    no routing
    vlan trunk native 1
    vlan trunk allowed 10,20,999
    lacp mode active
interface lag 256
    description VSX ISL
    no shutdown
    no routing
    vlan trunk native 1 tag
    vlan trunk allowed all
    lacp mode active
interface 1/1/1
    no shutdown
    mtu 9198
    lag 1
interface 1/1/47
    no shutdown
    lag 256
interface 1/1/48
    description UPLINK vers CORE-SW1 Te1/0/10
    no shutdown
    routing
    ip mtu 9000
    ip address 10.254.10.1/31
    ip ospf 1 area 0.0.0.0
    ip ospf network point-to-point
    no ip ospf passive
    ip ospf authentication message-digest
    ip ospf message-digest-key 1 md5 ciphertext AQBapOSPFSECRET
interface 1/1/50
    description Lien PROD
    no shutdown
    routing
    vrf attach PROD
    ip address 172.30.0.1/30
interface vlan 10
    description Passerelle USERS
    ip mtu 9198
    ip address 10.10.10.2/24
    active-gateway ip mac 02:00:0a:01:0a:01
    active-gateway ip 10.10.10.1
    ip helper-address 10.100.0.10
    ip helper-address 10.100.0.11
    ip ospf 1 area 0.0.0.0
interface vlan 20
    ip address 10.20.0.2/24
    vrrp 20 address-family ipv4
        address 10.20.0.1 primary
        priority 120
        no shutdown
    ip ospf 1 area 0.0.0.0
interface vlan 30
    vrf attach PROD
    ip address 10.30.0.1/24 
    ip address 10.31.0.1/24 secondary
vsx
    system-mac 02:00:00:00:10:01
    inter-switch-link lag 256
    role primary
    keepalive peer 192.168.10.12 source 192.168.10.11 vrf mgmt
router ospf 1
    router-id 10.255.1.11
    passive-interface default
    area 0.0.0.0
router bgp 65001
    bgp router-id 10.255.1.11
    neighbor 10.255.0.1 remote-as 65001
    neighbor 10.255.0.1 description CORE-SW1
    neighbor 10.255.0.1 update-source loopback 0
    neighbor 172.30.0.2 remote-as 64600
    neighbor 172.30.0.2 password ciphertext AQBapBGPSECRET
    address-family ipv4 unicast
        neighbor 10.255.0.1 activate
        network 10.10.10.0/24
    exit-address-family
    vrf PROD
        neighbor 172.30.0.2 remote-as 64600
        neighbor 172.30.0.2 description FW-PROD
        address-family ipv4 unicast
            neighbor 172.30.0.2 activate
            neighbor 172.30.0.2 route-map RM-FW-IN in
        exit-address-family
ip route 0.0.0.0/0 10.254.10.0
ip route 10.200.0.0/16 blackhole distance 250
ip route 0.0.0.0/0 172.30.0.2 vrf PROD
https-server vrf mgmt
"""

ACC_CX = """!Version ArubaOS-CX FL.10.13.1190
hostname ACC-CX-01
vlan 1,10,20,999
interface mgmt
    no shutdown
    ip dhcp
interface lag 1
    description UPLINK vers DIST-CX-01
    no shutdown
    no routing
    vlan trunk native 1
    vlan trunk allowed 10,20,999
    lacp mode active
interface 1/1/1
    description PC bureau
    no shutdown
    vlan access 10
    spanning-tree port-type admin-edge
interface 1/1/2
    description Telephone
    no shutdown
    vlan trunk native 10
    vlan trunk allowed 10,20
-- MORE --, next page: Space, next line: Enter, quit: Control-C\x1b[2K\rinterface 1/1/3
    shutdown
interface 1/1/49
    no shutdown
    lag 1
interface 1/1/50
    no shutdown
    lag 1
interface vlan 999
    description Management
    ip address 172.21.0.21/24
ip route 0.0.0.0/0 172.21.0.1
"""


@unittest.skipUnless((HERE / "lecteur_aoscx.py").exists(), "lecteur_aoscx.py absent")
class TestAosCx(TmpDirTest):
    def report(self, text, name):
        cfg = self.write_cfg(name, text)
        res = P.parse_config(cfg)
        out = self.tmp / (name + ".xlsx")
        P.write_report([res], out)
        return res, out

    def test_detection_and_identity(self):
        res, out = self.report(DIST_CX, "dist.txt")
        syn = rows(out, "Synthèse")[0]
        self.assertEqual((syn["Équipement"], syn["Version"], syn["Router-ID"], syn["AS BGP"]),
                         ("DIST-CX-01", "ArubaOS-CX PL.10.13.1190", "10.255.1.11", "65001"))
        self.assertEqual(syn["VRF (RD)"], "PROD (65001:100)")
        # une config Comware reste lue par le lecteur Comware
        self.assertNotIn("cw_name", P.parse_config(SAMPLE)["interfaces"]["Vlan-interface10"])

    def test_l2_ports_lag_and_inheritance(self):
        _, out = self.report(ACC_CX, "acc.cfg")
        l2 = {r["Interface"]: r for r in rows(out, "Interfaces L2")}
        self.assertEqual((l2["1/1/1"]["Mode"], l2["1/1/1"]["VLAN access"], l2["1/1/1"]["STP edge"]),
                         ("access", "10", "oui"))
        self.assertEqual((l2["1/1/2"]["Mode"], l2["1/1/2"]["PVID"], l2["1/1/2"]["VLANs autorisées (trunk)"]),
                         ("trunk", "10", "10,20"))
        self.assertEqual(l2["1/1/3"]["Statut admin"], "shutdown")       # ligne après « -- MORE -- »
        self.assertEqual((l2["lag 1"]["Type"], l2["lag 1"]["Membres (agrégat)"]), ("Agrégat L2", "1/1/49, 1/1/50"))
        self.assertEqual((l2["1/1/49"]["Membre de l'agrégat"], l2["1/1/49"]["VLANs autorisées (trunk)"]),
                         ("lag 1", "10,20,999"))
        self.assertTrue(l2["1/1/49"]["Config héritée"].startswith("oui"))

    def test_l3_interfaces(self):
        _, out = self.report(DIST_CX, "dist.txt")
        l3 = {r["Interface"]: r for r in rows(out, "Interfaces L3")}
        v10 = l3["vlan 10"]
        self.assertEqual((v10["Type"], v10["Adresse IP"], v10["Masque"]), ("SVI", "10.10.10.2", "255.255.255.0"))
        self.assertEqual(v10["Serveurs DHCP relay"], "10.100.0.10, 10.100.0.11")
        self.assertIn("10.10.10.1", v10["VRRP"])                          # active-gateway
        self.assertIn("10.20.0.1", l3["vlan 20"]["VRRP"])                 # vrrp N address-family
        self.assertEqual((l3["vlan 30"]["VRF"], l3["vlan 30"]["IP secondaires"]), ("PROD", "10.31.0.1/24"))
        self.assertEqual((l3["1/1/48"]["OSPF network-type"], l3["1/1/48"]["OSPF auth"]), ("p2p", "oui"))
        self.assertEqual((l3["loopback 0"]["Type"], l3["mgmt"]["Type"]), ("Loopback", "Management"))
        self.assertEqual(l3["1/1/50"]["VRF"], "PROD")

    def test_routing_bgp_statics(self):
        _, out = self.report(DIST_CX, "dist.txt")
        bgp = {(r["VRF"], r["Voisin / Groupe"]): r for r in rows(out, "Voisins BGP")}
        core = bgp[("public", "10.255.0.1")]
        self.assertEqual((core["Session"], core["Connect-interface"], core["Familles activées"]),
                         ("iBGP", "loopback 0", "ipv4 unicast"))
        fw = bgp[("PROD", "172.30.0.2")]
        self.assertEqual((fw["Session"], fw["Description"], fw["Route-policy import"]), ("eBGP", "FW-PROD", "RM-FW-IN"))
        st = {(r["VRF"], r["Destination"]): r for r in rows(out, "Routes statiques")}
        self.assertEqual(st[("public", "0.0.0.0/0")]["Next-hop"], "10.254.10.0")
        self.assertEqual((st[("public", "10.200.0.0/16")]["Interface sortie"],
                          st[("public", "10.200.0.0/16")]["Préférence"]), ("blackhole", "250"))
        self.assertEqual(st[("PROD", "0.0.0.0/0")]["Next-hop"], "172.30.0.2")
        rt = rows(out, "Routage dynamique")
        silent = {r["Valeur"] for r in rt if r["Catégorie"] == "Silent-interface"}
        self.assertEqual(silent, {"silent-interface all", "undo silent-interface 1/1/48"})

    def test_classification(self):
        _, out = self.report(DIST_CX, "dist.txt")
        plan = {(p["VRF"], p["Subnet"]): p["Catégorie"] for p in rows(out, "Plan d'adressage")}
        self.assertEqual(plan[("public", "10.254.10.0/31")], "INTERCO")
        self.assertEqual(plan[("PROD", "172.30.0.0/30")], "INTERCO")
        self.assertEqual(plan[("public", "10.10.10.0/24")], "LAN")
        self.assertEqual(plan[("public", "10.20.0.0/24")], "LAN")
        self.assertEqual(plan[("public", "192.168.10.0/24")], "MGMT")
        self.assertEqual(plan[("public", "10.255.1.11/32")], "LOOPBACK")
        self.assertEqual(plan[("PROD", "10.30.0.0/24")], "À VÉRIFIER")    # aucun indice
        rr = rows(out, "Réadressage")
        self.assertIn("VIP active-gateway", {r["Type d'adresse"] for r in rr})

    def test_no_secret(self):
        _, out = self.report(DIST_CX, "dist.txt")
        for _, c in all_cells(out):
            for secret in ("SECRETSECRET", "OSPFSECRET", "BGPSECRET"):
                self.assertNotIn(secret, str(c.value or ""))

    def test_mixed_tree(self):
        # CORE Comware -> DIST AOS-CX -> ACC AOS-CX
        root = self.tmp / "root"
        (root / "DIST-CX-01" / "ACC-CX-01").mkdir(parents=True)
        shutil.copy(SAMPLE, root / "core.cfg")
        (root / "DIST-CX-01" / "dist.txt").write_text(DIST_CX, encoding="utf-8")
        (root / "DIST-CX-01" / "ACC-CX-01" / "acc.cfg").write_text(ACC_CX, encoding="utf-8")
        r = self.run_cli(root, "-r")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("3 rapport(s)", r.stdout)
        f = root / "parc_hierarchique.xlsx"
        tree = {t["Nœud"]: t["Nœud parent"] for t in rows(f, "Arborescence")}
        self.assertEqual(tree, {"CORE-SW1": None, "DIST-CX-01": "CORE-SW1", "ACC-CX-01": "DIST-CX-01"})
        acc = next(x for x in rows(f, "Interfaces L2") if x["Équipement"] == "ACC-CX-01")
        self.assertEqual((acc["Niveau 0"], acc["Niveau 1"], acc["Niveau 2"]), ("CORE-SW1", "DIST-CX-01", "ACC-CX-01"))
        # 10.10.10.0/24 : SVI sur le CORE Comware ET sur le DIST AOS-CX, ports d'accès sur l'ACC AOS-CX
        p = next(x for x in rows(f, "Plan d'adressage") if x["Subnet"] == "10.10.10.0/24")
        self.assertEqual((p["Catégorie"], p["Nb équipements"]), ("LAN", 2))
        self.assertIn("ACC-CX-01", p["Raisons"])

    def test_missing_reader_is_an_error_not_a_wrong_report(self):
        cfg = self.write_cfg("dist.txt", DIST_CX)
        saved, P.AOSCX = P.AOSCX, None
        try:
            with self.assertRaises(RuntimeError):
                P.parse_config(cfg)
        finally:
            P.AOSCX = saved


if __name__ == "__main__":
    unittest.main(verbosity=2)
