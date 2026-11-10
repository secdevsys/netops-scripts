#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generer_arborescence.py - Génère un arbre de dossiers de test pour hpe5940_parser.py

Chaque dossier = un nœud (un switch HPE 5940 / Comware 7) contenant sa configuration
(<SYSNAME>.cfg ou .txt) ; ses sous-dossiers = ses nœuds enfants.

    arbre_test/
    └── CORE-01/                CORE-01.cfg
        ├── DIST-01/            DIST-01.txt
        │   ├── ACC-01/         ACC-01.cfg
        │   └── ACC-02/         ACC-02.txt
        └── DIST-02/ ...

Les configurations sont cohérentes entre elles :
  - CORE : loopback, OSPF area 0, BGP (eBGP vers un opérateur), liens L3 /30 vers chaque DIST ;
  - DIST : lien L3 /30 vers le CORE, SVI utilisateurs (DHCP relay, VRRP), trunks LACP vers les ACC ;
  - ACC  : trunk LACP vers son parent, ports access / hybrid (téléphone), SVI de management ;
  - les descriptions nomment l'équipement et le port d'en face (UPLINK / DOWNLINK).

Usage (Windows : "py" ou "python") :
  py generer_arborescence.py                                   -> .\\arbre_test (3 niveaux, 2 enfants)
  py generer_arborescence.py C:\\Temp\\arbre -n 4 -f 3          -> 4 niveaux, 3 enfants par nœud
  py generer_arborescence.py C:\\Temp\\arbre -f 3,4             -> 3 DIST sous le CORE, 4 ACC par DIST
  py generer_arborescence.py C:\\Temp\\arbre --pieges           -> + cas pièges pour tester le parser
  py generer_arborescence.py C:\\Temp\\arbre --force            -> écrase un dossier existant

Puis :  py hpe5940_parser.py C:\\Temp\\arbre -r --merge parc.xlsx
Aucune dépendance externe.
"""

__version__ = "1.0.0"

import argparse
import ipaddress
import random
import shutil
import sys
from pathlib import Path

ROLES = ["CORE", "DIST", "ACC"]          # au-delà du 3e niveau : ACC en cascade
UPLINKS = ("HundredGigE1/0/53", "HundredGigE1/0/54")


# --------------------------------------------------------------------------
# Plan d'adressage
# --------------------------------------------------------------------------
class Plan:
    """Distribue les adresses sans doublon dans tout l'arbre."""

    def __init__(self):
        self.p2p = ipaddress.ip_network("10.254.0.0/16").subnets(new_prefix=30)
        self.loop = iter(ipaddress.ip_network("10.255.0.0/16").hosts())
        self.mgmt = iter(ipaddress.ip_network("192.168.0.0/16").hosts())
        self.inband = iter(ipaddress.ip_network("172.20.0.0/16").hosts())   # SVI mgmt des ACC
        self.next(self.mgmt)                      # 192.168.0.1 = passerelle mgmt
        self.user_net = ipaddress.ip_network("10.0.0.0/8").subnets(new_prefix=24)
        self.next(self.user_net)                  # 10.0.0.0/24 réservé
        self.vlan = 100

    @staticmethod
    def next(it):
        return next(it)

    def new_p2p(self):
        a, b = list(self.next(self.p2p).hosts())
        return str(a), str(b)

    def new_vlans(self, count=3):
        vl = []
        for name in ("USERS", "VOIP", "IMPRIMANTES")[:count]:
            self.vlan += 10
            vl.append((self.vlan, name, self.next(self.user_net)))
        return vl


# --------------------------------------------------------------------------
# Modèle de l'arbre
# --------------------------------------------------------------------------
class Node:
    counters = {}

    def __init__(self, level, parent=None):
        self.level = level
        self.role = ROLES[min(level, len(ROLES) - 1)]
        Node.counters[self.role] = Node.counters.get(self.role, 0) + 1
        self.name = f"{self.role}-{Node.counters[self.role]:02d}"
        self.parent = parent
        self.children = []
        self.uplink = None          # dict décrivant le lien vers le parent
        self.downlinks = []         # liens vers les enfants
        self.vlans = []

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()


def build_tree(levels, fanout):
    Node.counters = {}
    root = Node(0)
    todo = [root]
    while todo:
        n = todo.pop(0)
        if n.level + 1 >= levels:
            continue
        for _ in range(fanout[min(n.level, len(fanout) - 1)]):
            c = Node(n.level + 1, n)
            n.children.append(c)
            todo.append(c)
    return root


def plan_links(root, plan):
    """Choisit les ports et adresses de chaque lien parent <-> enfant."""
    for n in root.walk():
        n.loopback = str(plan.next(plan.loop))
        n.mgmt_ip = str(plan.next(plan.mgmt))
        n.inband_ip = str(plan.next(plan.inband)) if n.role == "ACC" else None
        if n.role == "DIST":
            n.vlans = plan.new_vlans()
        port = 1
        for i, c in enumerate(n.children, 1):
            if c.role == "DIST":                    # CORE -> DIST : L3 point à point
                ip_parent, ip_child = plan.new_p2p()
                link = {"type": "l3", "parent_port": f"Ten-GigabitEthernet1/0/{port}",
                        "child_port": UPLINKS[0], "parent_ip": ip_parent, "child_ip": ip_child}
                port += 1
            else:                                   # DIST/ACC -> ACC : trunk LACP 2 liens
                link = {"type": "l2", "agg": i,
                        "parent_ports": [f"Ten-GigabitEthernet1/0/{port}",
                                         f"Ten-GigabitEthernet1/0/{port + 1}"],
                        "child_ports": list(UPLINKS)}
                port += 2
            link["parent"], link["child"] = n, c
            n.downlinks.append(link)
            c.uplink = link
    # VLANs : chaque ACC hérite des VLANs de son DIST (premier ancêtre DIST)
    for n in root.walk():
        if n.role == "ACC":
            p = n.parent
            while p is not None and p.role != "DIST":
                p = p.parent
            n.vlans = p.vlans if p else plan.new_vlans()


# --------------------------------------------------------------------------
# Génération des configurations Comware 7
# --------------------------------------------------------------------------
def vlan_ids(vlans):
    return " ".join(str(v[0]) for v in vlans)


def gen_config(n, rnd):
    L = []
    a = L.append
    a(f"<{n.name}>display current-configuration")
    a("#")
    a(" version 7.1.070, Release 6710P01")
    a("#")
    a(f" sysname {n.name}")
    a("#")
    a(" irf mac-address persistent timer")
    a("#")
    a(f" router id {n.loopback}")
    a("#")
    a(" lldp global enable")
    a("#")
    if n.role == "CORE":
        a("ip vpn-instance mgmt")
        a(" route-distinguisher 65000:999")
        a("#")

    # VLANs
    a("vlan 1")
    a("#")
    a("vlan 999")
    a(" name MGMT")
    a("#")
    for vid, name, _ in n.vlans:
        a(f"vlan {vid}")
        a(f" name {name}")
        a("#")

    # Agrégats L2 (montant puis descendants)
    allowed = f"{vlan_ids(n.vlans)} 999".strip()
    if n.uplink and n.uplink["type"] == "l2":
        a("interface Bridge-Aggregation1")
        a(f" description UPLINK vers {n.parent.name} BAGG{n.uplink['agg'] + 10}")
        a(" port link-type trunk")
        a(" undo port trunk permit vlan 1")
        a(f" port trunk permit vlan {allowed}")
        a(" link-aggregation mode dynamic")
        a("#")
    for d in n.downlinks:
        if d["type"] == "l2":
            a(f"interface Bridge-Aggregation{d['agg'] + 10}")
            a(f" description DOWNLINK vers {d['child'].name} BAGG1")
            a(" port link-type trunk")
            a(" undo port trunk permit vlan 1")
            a(f" port trunk permit vlan {allowed}")
            a(" link-aggregation mode dynamic")
            a("#")

    a("interface NULL0")
    a("#")
    a("interface LoopBack0")
    a(" description Router-ID")
    a(f" ip address {n.loopback} 255.255.255.255")
    if n.role in ("CORE", "DIST"):
        a(" ospf 1 area 0.0.0.0")
    a("#")

    # SVIs
    if n.role == "DIST":
        for vid, name, net in n.vlans:
            hosts = list(net.hosts())
            real_ip = hosts[1] if n.name.endswith(("1", "3", "5", "7", "9")) else hosts[2]
            a(f"interface Vlan-interface{vid}")
            a(f" description Passerelle {name}")
            a(f" ip address {real_ip} {net.netmask}")
            if name != "IMPRIMANTES":
                a(" dhcp select relay")
                a(" dhcp relay server-address 10.100.0.10")
                a(" dhcp relay server-address 10.100.0.11")
            if name == "VOIP":
                a(" ip mtu 1400")
            a(f" vrrp vrid {vid % 255} virtual-ip {hosts[0]}")
            a(f" vrrp vrid {vid % 255} priority {rnd.choice((110, 120))}")
            a(" ospf 1 area 0.0.0.0")
            a("#")
    if n.role == "ACC":
        a("interface Vlan-interface999")
        a(" description Management")
        a(f" ip address {n.inband_ip} 255.255.0.0")
        a("#")
    a("interface M-GigabitEthernet0/0/0")
    if n.role == "CORE":
        a(" ip binding vpn-instance mgmt")
    a(f" ip address {n.mgmt_ip} 255.255.0.0")
    a("#")

    # Ports physiques descendants
    for d in n.downlinks:
        c = d["child"]
        if d["type"] == "l3":
            a(f"interface {d['parent_port']}")
            a(" port link-mode route")
            a(f" description DOWNLINK vers {c.name} {d['child_port']}")
            a(f" ip address {d['parent_ip']} 255.255.255.252")
            a(" ospf 1 area 0.0.0.0")
            a(" ospf network-type p2p")
            a("#")
        else:
            for pp, cp in zip(d["parent_ports"], d["child_ports"]):
                a(f"interface {pp}")
                a(" port link-mode bridge")
                a(f" description DOWNLINK vers {c.name} {cp}")
                a(" port link-type trunk")
                a(" undo port trunk permit vlan 1")
                a(f" port trunk permit vlan {allowed}")
                a(f" port link-aggregation group {d['agg'] + 10}")
                a("#")

    # Ports d'accès (ACC)
    if n.role == "ACC" and n.vlans:
        users = n.vlans[0][0]
        voip = n.vlans[1][0] if len(n.vlans) > 1 else users
        prn = n.vlans[2][0] if len(n.vlans) > 2 else users
        first = 2 * len(n.downlinks) + 1
        for i in range(first, first + 8):
            a(f"interface Ten-GigabitEthernet1/0/{i}")
            a(" port link-mode bridge")
            kind = rnd.choice(("pc", "pc", "tel", "prn", "off"))
            if kind == "pc":
                a(f" description PC bureau {rnd.randint(100, 399)}")
                a(f" port access vlan {users}")
                a(" stp edged-port")
            elif kind == "tel":
                a(" description Telephone + PC")
                a(" port link-type hybrid")
                a(" undo port hybrid vlan 1")
                a(f" port hybrid vlan {voip} tagged")
                a(f" port hybrid vlan {users} untagged")
                a(f" port hybrid pvid vlan {users}")
                a(" stp edged-port")
            elif kind == "prn":
                a(f" description Imprimante etage {rnd.randint(0, 5)}")
                a(f" port access vlan {prn}")
                a(" stp edged-port")
            else:
                a(" description Libre")
                a(" shutdown")
            a("#")

    # Ports montants
    if n.uplink:
        p = n.parent
        if n.uplink["type"] == "l3":
            a(f"interface {n.uplink['child_port']}")
            a(" port link-mode route")
            a(f" description UPLINK vers {p.name} {n.uplink['parent_port']}")
            a(f" ip address {n.uplink['child_ip']} 255.255.255.252")
            a(" ospf 1 area 0.0.0.0")
            a(" ospf network-type p2p")
            a("#")
        else:
            for cp, pp in zip(n.uplink["child_ports"], n.uplink["parent_ports"]):
                a(f"interface {cp}")
                a(" port link-mode bridge")
                a(f" description UPLINK vers {p.name} {pp}")
                a(" port link-type trunk")
                a(" undo port trunk permit vlan 1")
                a(f" port trunk permit vlan {allowed}")
                a(" port link-aggregation group 1")
                a("#")

    # Routage
    if n.role in ("CORE", "DIST"):
        a(f"ospf 1 router-id {n.loopback}")
        if n.role == "CORE":
            a(" default-route-advertise")
        a(" silent-interface all")
        for d in n.downlinks:
            if d["type"] == "l3":
                a(f" undo silent-interface {d['parent_port']}")
        if n.uplink and n.uplink["type"] == "l3":
            a(f" undo silent-interface {n.uplink['child_port']}")
        a(" area 0.0.0.0")
        a(f"  network {n.loopback} 0.0.0.0")
        a("#")
    if n.role == "CORE":
        a("bgp 65000")
        a(f" router-id {n.loopback}")
        a(" peer 203.0.113.1 as-number 64500")
        a(" peer 203.0.113.1 description OPERATEUR-INTERNET")
        a(" peer 203.0.113.1 password cipher $c$3$GENERATED==")
        a(" #")
        a(" address-family ipv4 unicast")
        a("  network 10.0.0.0 255.0.0.0")
        a("  peer 203.0.113.1 enable")
        a("  peer 203.0.113.1 route-policy RP-IN import")
        a("  peer 203.0.113.1 route-policy RP-OUT export")
        a("#")
        a(" ip route-static 0.0.0.0 0 203.0.113.1 preference 70 description Defaut operateur")
        a(" ip route-static vpn-instance mgmt 0.0.0.0 0 192.168.0.1")
    else:
        a(" ip route-static 0.0.0.0 0 192.168.0.1 description Defaut management")
    a("#")
    a("return")
    return "\n".join(L) + "\n"


# --------------------------------------------------------------------------
# Écriture sur disque
# --------------------------------------------------------------------------
def write_tree(root, out, rnd):
    paths = {}
    for n in root.walk():
        d = (paths[n.parent.name] / n.name) if n.parent else (out / n.name)
        paths[n.name] = d
        d.mkdir(parents=True, exist_ok=True)
        ext = ".cfg" if rnd.random() < 0.5 else ".txt"
        (d / f"{n.name}{ext}").write_text(gen_config(n, rnd), encoding="utf-8", newline="\r\n")
    return paths


def add_traps(root, paths, rnd):
    """Cas réels qui doivent être gérés par le parser."""
    traps = []
    # 1. fichier qui n'est pas une config
    (paths[root.name] / "LISEZMOI.txt").write_text("Sauvegardes du 10/11\n", encoding="utf-8")
    traps.append("LISEZMOI.txt dans la racine (doit être ignoré)")
    # 2. dossier caché avec une config
    hidden = paths[root.name] / ".git"
    hidden.mkdir(exist_ok=True)
    (hidden / "x.cfg").write_text("#\n sysname CACHE\n#\n", encoding="utf-8")
    traps.append(".git/x.cfg (dossier caché, doit être ignoré)")
    nodes = [n for n in root.walk() if n.parent]
    if not nodes:
        return traps
    # 3. dossier intermédiaire sans config : on déplace un nœud dans un sous-dossier "Batiment X"
    n = rnd.choice([x for x in nodes if not x.children] or nodes)
    old = paths[n.name]
    new = old.parent / "Batiment A" / old.name
    new.parent.mkdir(exist_ok=True)
    shutil.move(str(old), str(new))
    for x in n.walk():
        paths[x.name] = Path(str(paths[x.name]).replace(str(old), str(new), 1))
    traps.append(f"{n.name} rangé sous un dossier 'Batiment A' sans config (parent attendu : {n.parent.name})")
    # 4. capture PuTTY : prompts "---- More ----", encodage Windows-1252, accents
    n, n2 = rnd.sample(nodes, 2) if len(nodes) >= 2 else (nodes[0], None)
    f = next(paths[n.name].glob(f"{n.name}.*"))
    text = f.read_text(encoding="utf-8").replace("\r\n", "\n")
    text = text.replace("interface LoopBack0", "  ---- More ----\x1b[16D                \x1b[16D"
                        "interface LoopBack0", 1)
    text = text.replace("description Router-ID", "description Router-ID équipement", 1)
    f.write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    traps.append(f"{f.name} : capture PuTTY (---- More ----, cp1252, accents)")
    # 5. description qui commence par '=' et caractère de contrôle
    if n2 is None:
        return traps
    n = n2
    f = next(paths[n.name].glob(f"{n.name}.*"))
    text = f.read_text(encoding="utf-8")
    text = text.replace(" description Router-ID", " description === LOOPBACK \x07===", 1)
    f.write_text(text, encoding="utf-8")
    traps.append(f"{f.name} : description '=== LOOPBACK ===' + caractère de contrôle")
    return traps


def print_tree(root, paths, out):
    def rec(n, prefix="", last=True):
        f = next(paths[n.name].glob(f"{n.name}.*")).name
        rel = paths[n.name].relative_to(out)
        branch = "" if not n.parent else ("└── " if last else "├── ")
        extra = f"   ({rel.parent})" if "Batiment" in str(rel) else ""
        print(f"{prefix}{branch}{n.name}/  {f}{extra}")
        kid_prefix = prefix + ("" if not n.parent else ("    " if last else "│   "))
        for i, c in enumerate(n.children):
            rec(c, kid_prefix, i == len(n.children) - 1)
    rec(root)


def parse_fanout(text):
    try:
        vals = [int(x) for x in str(text).split(",")]
    except ValueError:
        raise argparse.ArgumentTypeError("format attendu : 2 ou 1,3,4")
    if not vals or any(v < 1 or v > 40 for v in vals):
        raise argparse.ArgumentTypeError("chaque valeur doit être entre 1 et 40")
    return vals


def main():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except AttributeError:
            pass
    ap = argparse.ArgumentParser(description="Génère un arbre de configs HPE 5940 de test")
    ap.add_argument("dossier", nargs="?", default="arbre_test", help="dossier à créer (défaut : arbre_test)")
    ap.add_argument("-n", "--niveaux", type=int, default=3,
                    help="profondeur de l'arbre, racine comprise (défaut : 3 = CORE/DIST/ACC)")
    ap.add_argument("-f", "--enfants", type=parse_fanout, default=[2],
                    help="enfants par nœud : 2, ou une valeur par niveau ex. 3,4 (défaut : 2)")
    ap.add_argument("--graine", type=int, default=5940, help="graine aléatoire (même graine = même arbre)")
    ap.add_argument("--pieges", action="store_true",
                    help="ajoute des cas pièges (non-config, dossier caché, PuTTY, cp1252...)")
    ap.add_argument("--force", action="store_true", help="supprime le dossier s'il existe déjà")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args()

    if not 1 <= args.niveaux <= 8:
        ap.error("--niveaux doit être entre 1 et 8")
    out = Path(args.dossier).resolve()
    if out.exists() and any(out.iterdir()):
        if not args.force:
            sys.exit(f"[ERREUR] {out} existe et n'est pas vide. Ajoutez --force pour le remplacer.")
        shutil.rmtree(out)

    total, width = 1, 1
    for lvl in range(args.niveaux - 1):
        width *= args.enfants[min(lvl, len(args.enfants) - 1)]
        total += width
    if total > 2000:
        sys.exit(f"[ERREUR] {total} nœuds demandés : réduisez --niveaux ou --enfants (max 2000).")

    rnd = random.Random(args.graine)
    root = build_tree(args.niveaux, args.enfants)
    plan_links(root, Plan())
    out.mkdir(parents=True, exist_ok=True)
    paths = write_tree(root, out, rnd)
    traps = add_traps(root, paths, rnd) if args.pieges else []

    print(f"Arbre généré dans {out}\n")
    print_tree(root, paths, out)
    print(f"\n{total} nœud(s) sur {args.niveaux} niveau(x).")
    if traps:
        print("Cas pièges ajoutés :")
        for t in traps:
            print(f"  - {t}")
    print(f"\nPour tester le parser :\n  py hpe5940_parser.py \"{out}\" -r --merge parc.xlsx")


if __name__ == "__main__":
    main()
