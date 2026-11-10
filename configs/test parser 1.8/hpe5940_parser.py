#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hpe5940_parser.py - Parse une configuration HPE FlexFabric 5940 (Comware 7)
et produit un rapport Excel (.xlsx).

Onglets générés :
  - Synthèse          : équipement, version, router-id, VRF, compteurs
  - Interfaces L2     : mode access/trunk/hybrid, VLANs, PVID, agrégat, MTU, description...
  - Interfaces L3     : IP/masque, secondaires, IPv6, VRF, DHCP relay, IP MTU, VRRP, OSPF...
  - Routage dynamique : OSPF (process, areas, networks, interfaces), BGP, RIP, IS-IS
  - Voisins BGP       : un voisin / peer-group par ligne avec ses options
  - Routes statiques  : IPv4 / IPv6, VRF, next-hop, préférence, tag, description
  - VLANs             : ID, nom, description
  - Plan d'adressage  : un subnet IPv4 par ligne, classé INTERCO / LAN / MGMT / LOOPBACK /
                        À VÉRIFIER, avec la confiance, les raisons et les références à modifier
  - Réadressage       : une adresse par ligne (primaire, secondaire, VIP VRRP) avec les colonnes
                        vides "Nouvelle IP" / "Nouveau masque" à remplir par l'intégrateur

Usage (Windows : "py" ou "python") :
  py hpe5940_parser.py C:\\Backups\\SW1.cfg                -> C:\\Backups\\<sysname>.xlsx (dossier du .cfg/.txt)
  py hpe5940_parser.py configs\\*.cfg -d rapports        -> un rapport par équipement
  py hpe5940_parser.py C:\\Backups\\5940 -d rapports      -> tous les .cfg/.txt/.log d'un dossier
  py hpe5940_parser.py configs -d rapports --merge tous.xlsx -> + un rapport consolidé (optionnel)
  py hpe5940_parser.py C:\\Backups -r                     -> arbre des nœuds : un rapport par switch
                                                             + C:\\Backups\\parc_hierarchique.xlsx
  py hpe5940_parser.py C:\\Backups -r --merge parc.xlsx   -> même chose, consolidé nommé parc.xlsx

Arbre des nœuds (-r) : chaque dossier = un nœud (sa config .cfg/.txt), ses sous-dossiers = ses
enfants ; un nœud a un seul parent (le dossier qui le contient). Les liens symboliques et
jonctions ne sont pas suivis.
  py hpe5940_parser.py SW1.cfg --debug                    -> détail de la lecture (dépannage)
Les jokers (*.cfg) sont développés par le script : fonctionne sous cmd et PowerShell.

Entrée : sortie "display current-configuration" ou fichier startup.cfg.
Dépendance : openpyxl  (pip install openpyxl)
"""

__version__ = "1.8.0"

import argparse
import glob
import ipaddress
import os
import re
import subprocess
import sys
from collections import OrderedDict
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    sys.exit("openpyxl est requis : pip install openpyxl")


# --------------------------------------------------------------------------
# Utilitaires
# --------------------------------------------------------------------------
ALL_VLANS = frozenset(range(1, 4095))
IP_RE = re.compile(r"^\d{1,3}(\.\d{1,3}){3}$")

L2_ALWAYS = ("Bridge-Aggregation", "S-Channel")
L3_ALWAYS = ("Vlan-interface", "LoopBack", "Route-Aggregation", "Tunnel",
             "M-GigabitEthernet", "Vsi-interface", "Virtual-Template")
IGNORED_IFS = ("NULL", "Register-Tunnel", "InLoopBack")


SECRET_RE = re.compile(r"\b(cipher|simple|plain|hash)\s+(?:(?:cipher|plain|simple)\s+)?\S+",
                       re.IGNORECASE)


def mask(line):
    """Masque les secrets : cipher/simple/plain/hash <valeur> -> <mot-clé> ******"""
    return SECRET_RE.sub(lambda m: f"{m.group(1)} ******", line)


def natural_key(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(text))]


def parse_vlan_tokens(tokens):
    """['10','20','to','30'] -> {10,20..30} ; ['all'] -> 1..4094"""
    vlans = set()
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if tok == "all":
            return set(ALL_VLANS)
        if not tok.isdigit():
            i += 1
            continue
        if i + 2 < len(tokens) and tokens[i + 1] == "to" and tokens[i + 2].isdigit():
            vlans.update(range(int(tok), int(tokens[i + 2]) + 1))
            i += 3
        else:
            vlans.add(int(tok))
            i += 1
    return vlans


def fmt_vlans(vlans):
    """{1,2,3,10} -> '1-3,10'"""
    if not vlans:
        return ""
    if set(vlans) == ALL_VLANS:
        return "all (1-4094)"
    s = sorted(vlans)
    out, start, prev = [], s[0], s[0]
    for v in s[1:]:
        if v == prev + 1:
            prev = v
            continue
        out.append(str(start) if start == prev else f"{start}-{prev}")
        start = prev = v
    out.append(str(start) if start == prev else f"{start}-{prev}")
    return ",".join(out)


def ip_info(ip, netmask):
    """Retourne (ip, masque pointé, /prefix, réseau) ; netmask peut être '24' ou '255.255.255.0'."""
    try:
        net = ipaddress.IPv4Network(f"{ip}/{netmask}", strict=False)
        return ip, str(net.netmask), f"/{net.prefixlen}", f"{net.network_address}/{net.prefixlen}"
    except ValueError:
        return ip, netmask, "", ""


def wildcard_to_prefix(addr, wildcard):
    try:
        wc = int(ipaddress.IPv4Address(wildcard))
        netmask = str(ipaddress.IPv4Address(0xFFFFFFFF ^ wc))
        net = ipaddress.IPv4Network(f"{addr}/{netmask}", strict=False)
        return f"{net.network_address}/{net.prefixlen}"
    except ValueError:
        return f"{addr} {wildcard}"


# --------------------------------------------------------------------------
# Lecture / découpage en blocs
# --------------------------------------------------------------------------
# Prompt "---- More ----" suivi de l'effacement par le terminal (ESC[16D, espaces, ESC[16D ou
# retours arrière) : on supprime jusqu'à la DERNIÈRE séquence d'effacement, ce qui suit est
# la vraie ligne avec sa propre indentation. Sans séquence d'effacement, on supprime aussi
# les espaces qui suivent le prompt.
MORE_RE = re.compile(r"[ \t]*-+[ \t]*More[ \t]*-+"
                     r"(?:(?:[ \t]*(?:\x1b\[[0-9;]*[A-Za-z]|\x08))+|[ \t]*)")


def clean_lines(text):
    """Nettoie une capture CLI (prompts, ---- More ----, caractères de contrôle)."""
    lines = []
    for raw in text.splitlines():
        # "---- More ----" puis effacement par le terminal : ESC[16D + espaces + ESC[16D
        # (ou retours arrière). Tout ce bloc est supprimé, l'indentation réelle est conservée.
        line = MORE_RE.sub("", raw)
        line = re.sub(r"\x1b\[[0-9;]*[A-Za-z]|[\x08\r]", "", line)
        line = line.replace("\t", " ").rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if re.match(r"^[<\[][^>\]]+[>\]]", stripped):   # prompt <SW> ou [SW]
            continue
        if stripped == "return":
            break
        lines.append(line)
    return lines


def split_blocks(lines):
    """
    Découpe la configuration Comware en blocs.
    Un bloc = une ligne d'en-tête + ses lignes plus indentées.
    Corps : liste de (indentation_relative, texte).
    Les commandes globales indentées d'un espace (sysname, ip route-static...)
    deviennent des blocs à part entière.
    """
    blocks, cur, head_indent = [], None, 0
    for line in lines:
        indent = len(line) - len(line.lstrip(" "))
        text = line.strip()
        if text == "#":
            if indent == 0:
                cur = None
            continue
        if cur is None or indent <= head_indent:
            cur = {"head": text, "body": []}
            head_indent = indent
            blocks.append(cur)
        else:
            cur["body"].append((indent - head_indent, text))
    return blocks


# --------------------------------------------------------------------------
# Interfaces
# --------------------------------------------------------------------------
def new_interface(name):
    return {
        "name": name, "description": "", "shutdown": False, "link_mode": None,
        "link_type": None, "access_vlan": None, "pvid": None,
        "trunk_vlans": {1}, "hyb_tagged": set(), "hyb_untagged": {1},
        "agg_group": None, "lacp_mode": "", "mtu": "", "jumbo": "",
        "speed": "", "duplex": "", "stp_edge": False,
        "vrf": "", "ipv4": [], "ipv6": [], "dhcp_mode": "", "dhcp_servers": [],
        "ip_mtu": "", "vrrp": OrderedDict(), "ospf_proc": "", "ospf_area": "",
        "ospf_cost": "", "ospf_nettype": "", "ospf_auth": False,
        "dot1q": "", "others": [],
    }


def parse_interface(name, body):
    it = new_interface(name)
    for _, line in body:
        tok = line.split()
        if line.startswith("description "):
            it["description"] = line[len("description "):]
        elif line == "shutdown":
            it["shutdown"] = True
        elif line.startswith("port link-mode "):
            it["link_mode"] = tok[2]
        elif line.startswith("port link-type "):
            it["link_type"] = tok[2]
        elif line.startswith("port access vlan "):
            it["access_vlan"] = tok[3]
            it["link_type"] = it["link_type"] or "access"
        elif line.startswith(("port trunk pvid vlan ", "port hybrid pvid vlan ")):
            it["pvid"] = tok[4]
        elif line.startswith("port trunk permit vlan "):
            it["trunk_vlans"] |= parse_vlan_tokens(tok[4:])
        elif line.startswith("undo port trunk permit vlan "):
            it["trunk_vlans"] -= parse_vlan_tokens(tok[5:])
        elif line.startswith("port hybrid vlan "):
            vl = parse_vlan_tokens(tok[3:-1])
            if tok[-1] == "tagged":
                it["hyb_tagged"] |= vl
                it["hyb_untagged"] -= vl
            else:
                it["hyb_untagged"] |= vl
                it["hyb_tagged"] -= vl
        elif line.startswith("undo port hybrid vlan "):
            vl = parse_vlan_tokens(tok[4:])
            it["hyb_tagged"] -= vl
            it["hyb_untagged"] -= vl
        elif line.startswith("port link-aggregation group "):
            it["agg_group"] = tok[3]
        elif line.startswith("link-aggregation mode "):
            it["lacp_mode"] = tok[2]
        elif tok[0] == "mtu" and len(tok) > 1:
            it["mtu"] = tok[1]
        elif line.startswith("jumboframe enable"):
            it["jumbo"] = tok[2] if len(tok) > 2 else "oui"
        elif tok[0] == "speed" and len(tok) > 1:
            it["speed"] = tok[1]
        elif tok[0] == "duplex" and len(tok) > 1:
            it["duplex"] = tok[1]
        elif line.startswith("stp edged-port"):
            it["stp_edge"] = True
        elif line.startswith("ip binding vpn-instance "):
            it["vrf"] = tok[3]
        elif line.startswith(("ip address dhcp-alloc", "ip address unnumbered")):
            it["ipv4"].append({"ip": " ".join(tok[2:4]), "mask": "", "prefix": "", "network": "",
                               "sub": False})
        elif line.startswith("ip address ") and len(tok) >= 4:
            ip, netmask, pfx, net = ip_info(tok[2], tok[3])
            it["ipv4"].append({"ip": ip, "mask": netmask, "prefix": pfx, "network": net,
                               "sub": "sub" in tok[4:]})
        elif line.startswith("ipv6 address ") and len(tok) >= 3:
            addr = tok[2] if "/" in tok[2] or len(tok) < 4 or not tok[3].isdigit() else f"{tok[2]}/{tok[3]}"
            it["ipv6"].append(addr)
        elif line.startswith("dhcp select "):
            it["dhcp_mode"] = tok[2]
        elif line.startswith("dhcp relay server-address "):
            it["dhcp_servers"].append(tok[3])
        elif line.startswith("ip mtu "):
            it["ip_mtu"] = tok[2]
        elif line.startswith("vrrp vrid ") and len(tok) >= 5:
            v = it["vrrp"].setdefault(tok[2], {"vip": [], "prio": ""})
            if tok[3] == "virtual-ip":
                v["vip"].append(tok[4])
            elif tok[3] == "priority":
                v["prio"] = tok[4]
            else:
                it["others"].append(mask(line))
        elif re.match(r"^ospf \d+ area ", line):
            it["ospf_proc"], it["ospf_area"] = tok[1], tok[3]
        elif line.startswith("ospf cost "):
            it["ospf_cost"] = tok[2]
        elif line.startswith("ospf network-type "):
            it["ospf_nettype"] = tok[2]
        elif line.startswith("ospf authentication-mode"):
            it["ospf_auth"] = True
        elif line.startswith("vlan-type dot1q vid "):
            it["dot1q"] = " ".join(tok[3:])
        else:
            it["others"].append(mask(line))
    if it["link_type"] is None and it["access_vlan"] is None:
        it["_explicit_l2"] = False
        it["link_type"] = "access"
    else:
        it["_explicit_l2"] = True
    return it


def classify(it):
    """Retourne 'L2', 'L3' ou None (interface ignorée)."""
    n = it["name"]
    if n.startswith(IGNORED_IFS):
        return None
    if n.startswith(L2_ALWAYS):
        return "L2"
    if n.startswith(L3_ALWAYS) or "." in n or it["ipv4"] or it["ipv6"]:
        return "L3"
    return "L3" if it["link_mode"] == "route" else "L2"


# --------------------------------------------------------------------------
# Routage
# --------------------------------------------------------------------------
def parse_ospf(head, body, device):
    tok = head.split()
    proc = tok[1] if len(tok) > 1 else "1"
    rid = tok[tok.index("router-id") + 1] if "router-id" in tok else ""
    vrf = tok[tok.index("vpn-instance") + 1] if "vpn-instance" in tok else ""
    rows = []

    def add(area, cat, val):
        rows.append([device, "OSPF", proc, vrf or "public", rid, area, cat, val])

    area = ""
    for indent, line in body:
        if indent == 1 and line.startswith("area "):
            area = line.split()[1]
            add(area, "Zone", "déclarée")
            continue
        if indent == 1:
            area = ""
        t = line.split()
        if line.startswith("network ") and len(t) >= 3:
            add(area, "Network", wildcard_to_prefix(t[1], t[2]))
        elif line in ("stub", "nssa") or line.startswith(("stub ", "nssa ")):
            add(area, "Type de zone", line)
        elif line.startswith("import-route"):
            add(area, "Redistribution", line[len("import-route "):])
        elif "silent-interface" in line:
            add(area, "Silent-interface", line)
        elif line.startswith("authentication-mode"):
            add(area, "Authentification", mask(line))
        else:
            add(area, "Paramètre", mask(line))
    return proc, vrf, rid, rows


def parse_bgp(head, body, device):
    tok = head.split()
    local_as = tok[1] if len(tok) > 1 else ""
    rows, peers = [], OrderedDict()
    vrf, af, af_indent = "", "", 0
    rid = next((l.split()[1] for i, l in body if i == 1 and l.startswith("router-id ")), "")

    def peer(v, pid):
        key = (v, pid)
        if key not in peers:
            peers[key] = {"vrf": v or "public", "id": pid, "kind": "peer", "group": "",
                          "as": "", "group_type": "", "desc": "", "conn_if": "",
                          "auth": False, "bfd": False, "multihop": "", "afs": [],
                          "rp_in": [], "rp_out": [], "others": []}
        return peers[key]

    def add(cat, val):
        rows.append([device, "BGP", local_as, vrf or "public", rid, af, cat, val])

    for indent, line in body:
        t = line.split()
        if indent == 1:                       # retour au niveau "bgp" : on quitte la VRF
            vrf = ""
        if af and indent <= af_indent:        # on quitte l'address-family
            af = ""
        if indent == 1 and line.startswith("ip vpn-instance ") and len(t) >= 3:
            vrf, af = t[2], ""
            add("VRF", f"vpn-instance {vrf}")
            continue
        if line.startswith("address-family "):
            af, af_indent = line[len("address-family "):], indent
            add("Famille d'adresses", af)
            continue

        if line.startswith("router-id ") and len(t) >= 2:
            add("Router-ID", t[1])
        elif t[0] == "group" and len(t) >= 2:
            p = peer(vrf, t[1])
            p["kind"] = "group"
            p["group_type"] = t[2] if len(t) > 2 else "internal"
        elif t[0] == "peer" and len(t) >= 3:
            p = peer(vrf, t[1])
            sub, rest = t[2], t[3:]
            clean = mask(" ".join(t[2:]))
            if sub == "as-number" and rest:
                p["as"] = rest[0]
            elif sub == "group" and rest:
                p["group"] = rest[0]
            elif sub == "description":
                p["desc"] = " ".join(rest)
            elif sub == "connect-interface" and rest:
                p["conn_if"] = rest[0]
            elif sub == "password":
                p["auth"] = True
            elif sub == "bfd":
                p["bfd"] = True
            elif sub == "ebgp-max-hop":
                p["multihop"] = rest[0] if rest else "255"
            elif sub == "enable":
                p["afs"].append(af or "ipv4 unicast")
            elif sub == "route-policy" and len(rest) >= 2:
                (p["rp_in"] if rest[1] == "import" else p["rp_out"]).append(rest[0])
            else:
                p["others"].append(f"[{af}] {clean}" if af else clean)
        elif t[0] == "network":
            pfx = ip_info(t[1], t[2])[3] if len(t) >= 3 and IP_RE.match(t[1]) else " ".join(t[1:])
            add("Network annoncé", pfx or " ".join(t[1:]))
        elif t[0] == "import-route":
            add("Redistribution", " ".join(t[1:]))
        else:
            add("Paramètre", mask(line))

    # héritage depuis les peer-groups
    for p in peers.values():
        if p["kind"] == "group" and p["group_type"] == "internal" and not p["as"]:
            p["as"] = local_as
    for (v, _), p in peers.items():
        if p["kind"] == "peer" and p["group"]:
            g = peers.get((v, p["group"]))
            if g:
                p["as"] = p["as"] or g["as"]
                p["conn_if"] = p["conn_if"] or g["conn_if"]
                p["auth"] = p["auth"] or g["auth"]
                p["bfd"] = p["bfd"] or g["bfd"]
                p["afs"] = p["afs"] or [f"{a} (groupe)" for a in g["afs"]]
                p["rp_in"] = p["rp_in"] or g["rp_in"]
                p["rp_out"] = p["rp_out"] or g["rp_out"]
                if not p["as"] and g["group_type"] == "internal":
                    p["as"] = local_as
    return local_as, rid, rows, peers


def parse_generic_routing(proto, head, body, device):
    tok = head.split()
    proc = tok[1] if len(tok) > 1 else ""
    vrf = tok[tok.index("vpn-instance") + 1] if "vpn-instance" in tok else ""
    return [[device, proto, proc, vrf or "public", "", "", "Paramètre", mask(line)] for _, line in body] or \
           [[device, proto, proc, vrf or "public", "", "", "Process", head]]


def parse_static(line, device):
    t = line.split()
    v6 = t[0] == "ipv6"
    t = t[2:]
    r = {"device": device, "af": "IPv6" if v6 else "IPv4", "vrf": "public", "dest": "",
         "nh_vrf": "", "out_if": "", "nh": "", "pref": "", "tag": "", "desc": "", "others": []}
    if t and t[0] == "vpn-instance":
        r["vrf"], t = t[1], t[2:]
    if len(t) < 2:
        return None
    dest, plen, t = t[0], t[1], t[2:]
    r["dest"] = f"{dest}/{plen}" if v6 else ip_info(dest, plen)[3] or f"{dest} {plen}"
    i = 0
    while i < len(t):
        k = t[i]
        if k == "description":
            r["desc"] = " ".join(t[i + 1:])
            break
        if k in ("preference", "tag") and i + 1 < len(t):
            r["pref" if k == "preference" else "tag"] = t[i + 1]
            i += 2
        elif k == "vpn-instance" and i + 1 < len(t):
            r["nh_vrf"] = t[i + 1]
            i += 2
        elif IP_RE.match(k) or (v6 and ":" in k):
            r["nh"] = k
            i += 1
        elif re.match(r"^[A-Za-z][A-Za-z-]*\d", k):
            r["out_if"] = k
            i += 1
        else:
            r["others"].append(k)
            i += 1
    return r


# --------------------------------------------------------------------------
# Analyse d'un fichier
# --------------------------------------------------------------------------
def dbg(debug, msg):
    if debug:
        print(f"  [DEBUG] {msg}")


def parse_config(path, debug=False):
    p = Path(path)
    raw = p.read_bytes()
    dbg(debug, f"Fichier      : {p.resolve()}")
    dbg(debug, f"Taille       : {len(raw)} octets")
    try:
        text = raw.decode("utf-8-sig")          # UTF-8 avec ou sans BOM (Notepad)
        enc = "utf-8 (BOM)" if raw.startswith(b"\xef\xbb\xbf") else "utf-8"
    except UnicodeDecodeError:
        text = raw.decode("cp1252", errors="replace")   # capture PuTTY / SecureCRT Windows
        enc = "cp1252"
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16")
        enc = "utf-16"
    dbg(debug, f"Encodage     : {enc}")
    raw_lines = text.splitlines()
    dbg(debug, f"Lignes brutes: {len(raw_lines)}")
    if debug:
        print("  [DEBUG] 10 premières lignes :")
        for i, l in enumerate(raw_lines[:10], 1):
            print(f"  [DEBUG]   {i:>3} | {l[:100]!r}")
    lines = clean_lines(text)
    dbg(debug, f"Lignes utiles après nettoyage : {len(lines)}")
    if len(raw_lines) > 20 and len(lines) < len(raw_lines) / 4:
        dbg(debug, "ATTENTION : la lecture s'est arrêtée tôt (ligne 'return' ou prompt <..> en début ?)")
    blocks = split_blocks(lines)
    if debug:
        from collections import Counter
        kinds = Counter(b["head"].split()[0] for b in blocks if b["head"].split())
        dbg(debug, f"Blocs        : {len(blocks)}  -> " +
            ", ".join(f"{k}={v}" for k, v in kinds.most_common(8)))
        if not kinds.get("interface"):
            dbg(debug, "ATTENTION : aucune ligne 'interface ...' trouvée - est-ce bien une config Comware ?")
    res = {"device": Path(path).stem, "version": "", "router_id": "", "vrfs": OrderedDict(),
           "interfaces": OrderedDict(), "vlans": OrderedDict(), "routing": [],
           "bgp_peers": [], "bgp_as": "", "statics": [],
           "source": str(p.resolve()), "sysname_found": False, "parent": "", "children": [],
           "level": ""}

    for b in blocks:
        if b["head"].startswith("sysname "):
            res["device"] = b["head"][len("sysname "):].strip()
            res["sysname_found"] = True
    dev = res["device"]

    for b in blocks:
        head, body = b["head"], b["body"]
        tok = head.split()
        key = tok[0]
        if key == "version":
            res["version"] = head[len("version "):]
        elif head.startswith("router id "):
            res["router_id"] = tok[2]
        elif head.startswith("ip vpn-instance ") and len(tok) >= 3:
            rd = next((l.split()[1] for _, l in body if l.startswith("route-distinguisher ")), "")
            res["vrfs"][tok[2]] = rd
        elif key == "vlan" and len(tok) >= 2:
            ids = parse_vlan_tokens(tok[1:])
            name = next((l[5:] for _, l in body if l.startswith("name ")), "")
            desc = next((l[12:] for _, l in body if l.startswith("description ")), "")
            for v in sorted(ids):
                res["vlans"][v] = {"name": name, "desc": desc}
        elif key == "interface" and len(tok) >= 2:
            name = " ".join(tok[1:])
            res["interfaces"][name] = parse_interface(name, body)
        elif key == "ospf":
            _, vrf, _, rows = parse_ospf(head, body, dev)
            res["routing"].extend(rows)
            if vrf:
                res["vrfs"].setdefault(vrf, "")
        elif key == "bgp":
            las, _, rows, peers = parse_bgp(head, body, dev)
            res["bgp_as"] = las
            res["routing"].extend(rows)
            for p in peers.values():
                p["local_as"] = las
                res["bgp_peers"].append(p)
        elif key in ("rip", "isis", "ospfv3", "ripng"):
            res["routing"].extend(parse_generic_routing(key.upper(), head, body, dev))
        elif head.startswith(("ip route-static ", "ipv6 route-static ")):
            r = parse_static(head, dev)
            if r:
                res["statics"].append(r)

    dbg(debug, f"sysname      : {res['device']}  (version {res['version'] or '?'})")
    dbg(debug, f"Trouvé       : {len(res['interfaces'])} interfaces, {len(res['vlans'])} VLANs, "
               f"{len(res['bgp_peers'])} voisins/groupes BGP, {len(res['statics'])} routes statiques")

    # Agrégats : membres + héritage de la config L2 du Bridge-Aggregation
    ifs = res["interfaces"]
    for it in ifs.values():
        it["members"] = []
    for it in ifs.values():
        g = it["agg_group"]
        if not g:
            continue
        cands = [f"Route-Aggregation{g}", f"Bridge-Aggregation{g}"]
        if it["link_mode"] != "route":
            cands.reverse()
        agg = next((c for c in cands if c in ifs), cands[0])
        it["agg_name"] = agg
        if agg in ifs:
            ifs[agg]["members"].append(it["name"])
            parent = ifs[agg]
            if agg.startswith("Bridge") and not it["_explicit_l2"]:
                for k in ("link_type", "access_vlan", "pvid", "trunk_vlans", "hyb_tagged", "hyb_untagged"):
                    it[k] = parent[k]
                it["inherited"] = True

    # Interfaces OSPF -> onglet routage
    for it in ifs.values():
        if it["ospf_proc"]:
            det = [f"{it['name']}"]
            if it["ospf_cost"]:
                det.append(f"cost {it['ospf_cost']}")
            if it["ospf_nettype"]:
                det.append(f"network-type {it['ospf_nettype']}")
            if it["ospf_auth"]:
                det.append("auth")
            res["routing"].append([dev, "OSPF", it["ospf_proc"], it["vrf"] or "public", "",
                                   it["ospf_area"], "Interface", " | ".join(det)])
    return res


# --------------------------------------------------------------------------
# Construction des lignes Excel
# --------------------------------------------------------------------------
L2_COLS = ["Équipement", "Interface", "Type", "Description", "Statut admin", "Mode",
           "VLAN access", "PVID", "VLANs autorisées (trunk)", "Nb VLANs trunk",
           "Hybrid tagged", "Hybrid untagged", "Membre de l'agrégat", "Mode LACP",
           "Membres (agrégat)", "MTU / Jumbo", "Vitesse", "Duplex", "STP edge",
           "Config héritée", "Autres commandes"]

L3_COLS = ["Équipement", "Interface", "Type", "Description", "Statut admin", "VRF",
           "Adresse IP", "Masque", "Préfixe", "Réseau", "Catégorie subnet", "IP secondaires", "IPv6",
           "DHCP mode", "Serveurs DHCP relay", "IP MTU", "MTU", "VRRP",
           "OSPF process", "OSPF area", "OSPF cost", "OSPF network-type", "OSPF auth",
           "dot1q", "Membre de l'agrégat", "Membres (agrégat)", "Autres commandes"]

RT_COLS = ["Équipement", "Protocole", "Process / AS", "VRF", "Router-ID",
           "Zone / Famille", "Catégorie", "Valeur"]

BGP_COLS = ["Équipement", "AS local", "VRF", "Voisin / Groupe", "Type", "Peer-group",
            "AS distant", "Session", "Description", "Connect-interface", "Auth MD5",
            "BFD", "eBGP multihop", "Familles activées", "Route-policy import",
            "Route-policy export", "Autres options"]

ST_COLS = ["Équipement", "Famille", "VRF", "Destination", "Interface sortie",
           "Next-hop", "VRF next-hop", "Préférence", "Tag", "Description", "Autres"]

VLAN_COLS = ["Équipement", "VLAN", "Nom", "Description", "Interface L3 (SVI)",
             "Ports access", "Ports trunk/hybrid"]


def if_type(name):
    if "." in name:
        return "Sous-interface"
    for p, label in (("Bridge-Aggregation", "Agrégat L2"), ("Route-Aggregation", "Agrégat L3"),
                     ("Vlan-interface", "SVI"), ("LoopBack", "Loopback"),
                     ("M-GigabitEthernet", "Management"), ("Tunnel", "Tunnel")):
        if name.startswith(p):
            return label
    return "Physique"


def yn(b):
    return "oui" if b else ""


def build_rows(res):
    dev = res["device"]
    l2, l3 = [], []
    for it in sorted(res["interfaces"].values(), key=lambda x: natural_key(x["name"])):
        cls = classify(it)
        status = "shutdown" if it["shutdown"] else "up"
        members = ", ".join(sorted(it["members"], key=natural_key))
        mtu = " / ".join(x for x in (f"mtu {it['mtu']}" if it["mtu"] else "",
                                     f"jumbo {it['jumbo']}" if it["jumbo"] else "") if x)
        if cls == "L2":
            mode = it["link_type"]
            trunk = it["trunk_vlans"] if mode == "trunk" else set()
            l2.append([
                dev, it["name"], if_type(it["name"]), it["description"], status, mode,
                (it["access_vlan"] or "1") if mode == "access" else "",
                (it["pvid"] or "1") if mode in ("trunk", "hybrid") else "",
                fmt_vlans(trunk), len(trunk) if trunk else "",
                fmt_vlans(it["hyb_tagged"]) if mode == "hybrid" else "",
                fmt_vlans(it["hyb_untagged"]) if mode == "hybrid" else "",
                it.get("agg_name", ""), it["lacp_mode"], members, mtu,
                it["speed"], it["duplex"], yn(it["stp_edge"]),
                f"oui ({it['agg_name']})" if it.get("inherited") else "",
                "\n".join(it["others"]),
            ])
        elif cls == "L3":
            prim = next((a for a in it["ipv4"] if not a["sub"]), None) or {}
            secs = ", ".join(f"{a['ip']}{a['prefix']}" for a in it["ipv4"] if a["sub"])
            vrrp = "; ".join(f"vrid {k}: {','.join(v['vip'])}" + (f" prio {v['prio']}" if v["prio"] else "")
                             for k, v in it["vrrp"].items())
            l3.append([
                dev, it["name"], if_type(it["name"]), it["description"], status,
                it["vrf"] or "public", prim.get("ip", ""), prim.get("mask", ""),
                prim.get("prefix", ""), prim.get("network", ""),
                res.get("_cat", {}).get((it["vrf"] or "public", prim.get("network", "")), ("",))[0],
                secs,
                ", ".join(it["ipv6"]), it["dhcp_mode"], ", ".join(it["dhcp_servers"]),
                it["ip_mtu"], it["mtu"], vrrp, it["ospf_proc"], it["ospf_area"],
                it["ospf_cost"], it["ospf_nettype"], yn(it["ospf_auth"]), it["dot1q"],
                it.get("agg_name", ""), members, "\n".join(it["others"]),
            ])

    bgp = []
    for p in res["bgp_peers"]:
        if p["kind"] == "group":
            session = "iBGP" if p["group_type"] == "internal" else "eBGP"
        elif p["as"]:
            session = "iBGP" if p["as"] == p["local_as"] else "eBGP"
        else:
            session = ""
        bgp.append([dev, p["local_as"], p["vrf"], p["id"], p["kind"], p["group"], p["as"],
                    session, p["desc"], p["conn_if"], yn(p["auth"]), yn(p["bfd"]),
                    p["multihop"], ", ".join(p["afs"]), ", ".join(p["rp_in"]),
                    ", ".join(p["rp_out"]), "\n".join(p["others"])])

    st = [[r["device"], r["af"], r["vrf"], r["dest"], r["out_if"], r["nh"], r["nh_vrf"],
           r["pref"], r["tag"], r["desc"], " ".join(r["others"])] for r in res["statics"]]

    # VLANs avec usage
    acc, trk = {}, {}
    for it in res["interfaces"].values():
        if classify(it) != "L2" or it.get("agg_name", "").startswith("Bridge-Aggregation"):
            continue
        if it["link_type"] == "access":
            acc.setdefault(int(it["access_vlan"] or 1), []).append(it["name"])
        elif it["link_type"] == "trunk":
            for v in it["trunk_vlans"]:
                trk.setdefault(v, []).append(it["name"])
        elif it["link_type"] == "hybrid":
            for v in it["hyb_tagged"] | it["hyb_untagged"]:
                trk.setdefault(v, []).append(it["name"])
    vl = []
    for v, d in res["vlans"].items():
        svi = f"Vlan-interface{v}"
        vl.append([dev, v, d["name"], d["desc"], svi if svi in res["interfaces"] else "",
                   ", ".join(sorted(acc.get(v, []), key=natural_key)),
                   ", ".join(sorted(trk.get(v, []), key=natural_key))])

    rt = sorted(res["routing"], key=lambda r: (r[1], natural_key(r[2]), r[3]))
    return l2, l3, rt, bgp, st, vl


# --------------------------------------------------------------------------
# Écriture Excel
# --------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="01A982")   # vert HPE
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SHUT_FILL = PatternFill("solid", fgColor="F2F2F2")


def xl(value):
    """Rend une valeur sûre pour Excel : supprime les caractères de contrôle interdits."""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def append_row(ws, values):
    """ws.append() sans risque : pas de caractère interdit, et un texte commençant par
    '=' (ex. description '=== UPLINK ===') reste du texte au lieu de devenir une formule."""
    ws.append([xl(v) for v in values])
    for c in ws[ws.max_row]:
        if isinstance(c.value, str) and c.value.startswith("="):
            c.data_type = "s"


LEVEL_FILL = PatternFill("solid", fgColor="425563")   # gris-bleu : colonnes de hiérarchie
LEVEL_CELL_FILL = PatternFill("solid", fgColor="EEF2F5")


def write_sheet(wb, title, cols, rows, status_col=None, levels=0, freeze_col=3, cat_col=None):
    """levels = nombre de colonnes « Niveau n » placées en tête (0 = pas de hiérarchie)."""
    ws = wb.create_sheet(title)
    cols = [f"Niveau {i}" for i in range(levels)] + list(cols)
    if status_col is not None:
        status_col += levels
    ws.append(cols)
    for i, c in enumerate(ws[1]):
        c.fill = LEVEL_FILL if i < levels else HEADER_FILL
        c.font, c.border = HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        append_row(ws, r)
    widths = [len(str(c)) for c in cols]
    for row in ws.iter_rows(min_row=2):
        shut = status_col is not None and row[status_col].value == "shutdown"
        for i, c in enumerate(row):
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if shut:
                c.fill = SHUT_FILL
                c.font = Font(color="808080")
            elif i < levels:
                c.fill = LEVEL_CELL_FILL
                c.font = Font(bold=True)
            elif cat_col is not None and cols[i] == cat_col and c.value in CAT_FILLS:
                c.fill = PatternFill("solid", fgColor=CAT_FILLS[c.value])
                c.font = Font(bold=True)
            longest = max((len(x) for x in str(c.value or "").split("\n")), default=0)
            widths[i] = max(widths[i], longest)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 8), 60)
    ws.freeze_panes = f"{get_column_letter(levels + freeze_col)}2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(rows) + 1}"
    ws.row_dimensions[1].height = 32
    return ws


TREE_COLS = ["Niveau", "Nœud", "Nœud parent", "Nœud(s) enfant(s)", "Arbre", "Dossier"]


def hierarchy(res, depth):
    """Colonnes Niveau 0..depth-1 : chemin racine -> équipement, complété par des vides."""
    path = res.get("path") or []
    return path + [""] * (depth - len(path))


def write_report(results, out, tree=None):
    # ordre de l'arbre (parent avant ses enfants) quand l'arborescence est connue
    results = sorted(results, key=lambda r: r.get("order", 10 ** 9))
    depth = max((len(r.get("path") or []) for r in results), default=0)
    # classification des subnets : celle calculée sur tout le parc si disponible
    plan = results[0].get("_plan") if results and all("_plan" in r for r in results) else None
    if plan is None:
        plan = classify_subnets(results)
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    ws.append(["Équipement", "Version", "Router-ID", "AS BGP", "VRF (RD)", "Nb VLANs",
               "Interfaces L2", "Interfaces L3", "Interfaces shutdown", "Process OSPF",
               "Voisins BGP", "Routes statiques", "Nœud parent", "Nœud(s) enfant(s)",
               "Niveau", "Fichier source"])
    all_rows = [[] for _ in range(6)]
    for res in results:
        rows = build_rows(res)
        prefix = hierarchy(res, depth)
        for acc, r in zip(all_rows, rows):
            acc.extend(prefix + list(x) for x in r)
        l2, l3, rt, bgp, st, vl = rows
        ospf_procs = sorted({r[2] for r in rt if r[1] == "OSPF"}, key=natural_key)
        append_row(ws, [res["device"], res["version"], res["router_id"], res["bgp_as"],
                   ", ".join(f"{k} ({v})" if v else k for k, v in res["vrfs"].items()),
                   len(res["vlans"]), len(l2), len(l3),
                   sum(1 for r in l2 if r[4] == "shutdown") + sum(1 for r in l3 if r[4] == "shutdown"),
                   ", ".join(ospf_procs), sum(1 for p in bgp if p[4] == "peer"), len(st),
                   res.get("parent", ""), ", ".join(res.get("children", [])),
                   res.get("level", ""), res.get("source", "")])
    for c in ws[1]:
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for i in range(1, 17):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["E"].width = 34
    ws.column_dimensions["M"].width = 30
    ws.column_dimensions["N"].width = 30
    ws.column_dimensions["O"].width = 8
    ws.column_dimensions["P"].width = 60
    ws.freeze_panes = "B2"

    l2, l3, rt, bgp, st, vl = all_rows
    write_sheet(wb, "Interfaces L2", L2_COLS, l2, status_col=4, levels=depth)
    write_sheet(wb, "Interfaces L3", L3_COLS, l3, status_col=4, levels=depth, cat_col="Catégorie subnet")
    write_sheet(wb, "Routage dynamique", RT_COLS, rt, levels=depth)
    write_sheet(wb, "Voisins BGP", BGP_COLS, bgp, levels=depth)
    write_sheet(wb, "Routes statiques", ST_COLS, st, levels=depth)
    write_sheet(wb, "VLANs", VLAN_COLS, vl, levels=depth)
    prow, rrow = plan_rows(plan, results, depth)
    write_sheet(wb, "Plan d'adressage", PLAN_COLS, prow, levels=depth, freeze_col=4, cat_col="Catégorie")
    write_sheet(wb, "Réadressage", READDR_COLS, rrow, levels=depth, cat_col="Catégorie")
    if tree:
        write_sheet(wb, "Arborescence", TREE_COLS, tree)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"L2": len(l2), "L3": len(l3), "Routage": len(rt), "BGP": len(bgp),
            "Statiques": len(st), "VLANs": len(vl), "Subnets": len(prow)}



# --------------------------------------------------------------------------
# Classification des subnets : INTERCO / LAN / MGMT / LOOPBACK / À VÉRIFIER
# --------------------------------------------------------------------------
# Principe : chaque subnet IPv4 reçoit des points « interco » et des points « LAN » à partir
# d'indices trouvés sur TOUS les équipements analysés (un subnet LAN peut être porté par une
# SVI du CORE, d'un DIST ou d'un ACC ; ses ports d'accès peuvent être plus bas dans l'arbre).
# La catégorie l'emporte avec au moins 3 points et plus que l'autre ; sinon « À VÉRIFIER ».
# Chaque point est justifié dans la colonne « Raisons ».

CAT_INTERCO, CAT_LAN, CAT_MGMT, CAT_LOOP, CAT_CHECK = "INTERCO", "LAN", "MGMT", "LOOPBACK", "À VÉRIFIER"
CAT_ORDER = {CAT_INTERCO: 0, CAT_LAN: 1, CAT_MGMT: 2, CAT_LOOP: 3, CAT_CHECK: 4}
CAT_FILLS = {CAT_INTERCO: "FCE4D6", CAT_LAN: "E2EFDA", CAT_MGMT: "DDEBF7",
             CAT_LOOP: "EDEDED", CAT_CHECK: "FFC7CE"}

INTERCO_WORDS = re.compile(r"interco|transit|\bwan\b|p2p|point.?(a|à).?point|backbone|uplink|"
                           r"downlink|\blien\b|liaison|peering|\bbgp\b|core.?link", re.I)
LAN_WORDS = re.compile(r"user|utilisat|voip|voix|t[ée]l[ée]phon|serv|imprim|print|wifi|wlan|"
                       r"\blan\b|poste|bureau|data|camera|video|badge|invit|guest|dmz", re.I)
MGMT_WORDS = re.compile(r"mgmt|manage|admin|\boob\b|supervision|\bnms\b|\bipmi\b|\bilo\b", re.I)


def _net(cidr):
    try:
        return ipaddress.IPv4Network(cidr, strict=False)
    except ValueError:
        return None


def _addr(ip):
    try:
        return ipaddress.IPv4Address(ip)
    except ValueError:
        return None


def _ospf_silent_rules(res):
    """{process: (silent_all, {interfaces silencieuses}, {interfaces non silencieuses})}"""
    rules = {}
    for r in res["routing"]:
        if r[1] != "OSPF" or r[6] != "Silent-interface":
            continue
        all_, on, off = rules.get(r[2], (False, set(), set()))
        t = r[7].split()
        if t[0] == "undo" and len(t) >= 3:
            if t[2] == "all":
                all_ = False
            else:
                off.add(t[2])
        elif len(t) >= 2:
            if t[1] == "all":
                all_ = True
            else:
                on.add(t[1])
        rules[r[2]] = (all_, on, off)
    return rules


def _ospf_networks(res):
    """[(process, vrf, réseau)] des commandes « network » OSPF."""
    out = []
    for r in res["routing"]:
        if r[1] == "OSPF" and r[6] == "Network":
            n = _net(r[7])
            if n:
                out.append((r[2], r[3], n))
    return out


def _subtree(res, results):
    path = res.get("path")
    if not path:
        return [res]
    return [r for r in results if r.get("path") and r["path"][:len(path)] == path]


def _endpoint_vlans(res):
    """VLAN -> ports d'extrémité (access explicites, hybrid, trunks edged) de l'équipement."""
    out = {}
    for it in res["interfaces"].values():
        if classify(it) != "L2" or it.get("agg_name", "").startswith("Bridge-Aggregation"):
            continue
        vl = set()
        if it["link_type"] == "access" and it["access_vlan"]:
            vl = {int(it["access_vlan"])}
        elif it["link_type"] == "hybrid":
            vl = set(it["hyb_tagged"]) | set(it["hyb_untagged"])
        elif it["link_type"] == "trunk" and it["stp_edge"]:
            vl = set(it["trunk_vlans"])
        for v in vl:
            out.setdefault(v, []).append(it["name"])
    return out


def _svi_vlan(name):
    m = re.match(r"Vlan-interface(\d+)$", name)
    return int(m.group(1)) if m else None


def classify_subnets(results):
    """Retourne la liste des subnets classés (dicts) et annote chaque résultat (res["_cat"])."""
    subnets = {}
    for res in results:
        for it in res["interfaces"].values():
            if classify(it) != "L3":
                continue
            vrf = it["vrf"] or "public"
            for a in it["ipv4"]:
                n = _net(a["network"]) if a["network"] else None
                if not n:
                    continue
                sn = subnets.setdefault((vrf, str(n)), {"vrf": vrf, "net": n, "members": [], "vips": []})
                sn["members"].append({"res": res, "it": it, "ip": a["ip"], "mask": a["mask"],
                                      "kind": "secondaire" if a["sub"] else "primaire"})
            for vrid, v in it["vrrp"].items():
                for vip in v["vip"]:
                    va = _addr(vip)
                    for a in it["ipv4"]:
                        n = _net(a["network"]) if a["network"] else None
                        if n and va and va in n:
                            subnets[(vrf, str(n))]["vips"].append(
                                {"res": res, "it": it, "ip": vip, "mask": a["mask"],
                                 "kind": f"VIP VRRP {vrid}"})
                            break

    # indices globaux
    endpoint = {id(r): _endpoint_vlans(r) for r in results}
    silent = {id(r): _ospf_silent_rules(r) for r in results}
    ospf_nets = {id(r): _ospf_networks(r) for r in results}
    vlan_names = {id(r): {v: (d["name"] + " " + d["desc"]).strip() for v, d in r["vlans"].items()}
                  for r in results}
    bgp_peers = [(r, p["id"], p["vrf"]) for r in results for p in r["bgp_peers"] if _addr(p["id"])]
    statics = [(r, s) for r in results for s in r["statics"]]
    dhcp_srv = {}                                   # (vrf, serveur) -> équipements relais
    for r in results:
        for it in r["interfaces"].values():
            for srv in it["dhcp_servers"]:
                dhcp_srv.setdefault((it["vrf"] or "public", srv), set()).add(r["device"])

    plan = []
    for key, sn in subnets.items():
        n, members = sn["net"], sn["members"]
        devs = sorted({m["res"]["device"] for m in members}, key=natural_key)
        names = [m["it"]["name"] for m in members]
        lan, inter, why = 0, 0, []

        def add(side, pts, reason):
            nonlocal lan, inter
            if side == "L":
                lan += pts
            else:
                inter += pts
            why.append(f"{'+' if pts > 0 else ''}{pts} {'LAN' if side == 'L' else 'INTERCO'} : {reason}")

        # catégories directes
        cat = conf = None
        if all(nm.startswith("LoopBack") for nm in names):
            cat, conf = CAT_LOOP, "haute"
            why.append("interface LoopBack")
        elif all(nm.startswith("M-GigabitEthernet") for nm in names):
            cat, conf = CAT_MGMT, "haute"
            why.append("port de management hors bande (M-GigabitEthernet)")
        else:
            labels = []
            for m in members:
                v = _svi_vlan(m["it"]["name"])
                labels.append(m["it"]["description"])
                if v is not None:
                    labels.append(vlan_names[id(m["res"])].get(v, ""))
            label = " ".join(dict.fromkeys(x for x in labels if x))     # sans doublon
            if sn["vrf"].lower() in ("mgmt", "management", "oob") or (MGMT_WORDS.search(label)
                                                                    and not LAN_WORDS.search(label)):
                cat, conf = CAT_MGMT, "moyenne"
                why.append(f"VRF ou nom de management ({sn['vrf']} / {label[:60]})")

        if cat is None:
            # --- taille du subnet
            if n.prefixlen >= 30:
                add("I", 3, f"masque /{n.prefixlen} (point à point)")
            elif n.prefixlen <= 26:
                add("L", 1, f"masque /{n.prefixlen} (plusieurs hôtes)")
            # --- nombre d'équipements du parc dans le subnet
            has_vrrp = bool(sn["vips"]) or any(m["it"]["vrrp"] for m in members)
            if len(devs) >= 2 and not has_vrrp:
                add("I", 2, f"partagé par {len(devs)} équipements du parc sans VRRP ({', '.join(devs)})")
            # --- services LAN
            if any(m["it"]["dhcp_mode"] in ("relay", "server") or m["it"]["dhcp_servers"] for m in members):
                add("L", 3, "DHCP relay / serveur DHCP sur l'interface")
            if has_vrrp:
                vips = ", ".join(sorted({v["ip"] for v in sn["vips"]}))
                add("L", 3, f"passerelle VRRP ({vips})" if vips else "VRRP configuré sur l'interface")
            hosted = sorted(srv for v, srv in dhcp_srv if v == sn["vrf"] and _addr(srv) and _addr(srv) in n)
            if hosted:
                add("L", 2, f"héberge les serveurs DHCP {', '.join(hosted)}")
            # --- ports d'accès dans le VLAN de la SVI (équipement + toute sa descendance)
            for m in members:
                v = _svi_vlan(m["it"]["name"])
                if v is None:
                    continue
                where = {}
                for r in _subtree(m["res"], results):
                    ports = endpoint[id(r)].get(v, [])
                    if ports:
                        where[r["device"]] = len(ports)
                if where:
                    add("L", 3, f"VLAN {v} : ports d'accès sur " +
                        ", ".join(f"{d} ({c})" for d, c in sorted(where.items(), key=lambda x: natural_key(x[0]))))
                    break
            # --- routage
            for m in members:
                it, r = m["it"], m["res"]
                ip = _addr(m["ip"])
                procs = {it["ospf_proc"]} if it["ospf_proc"] else set()
                procs |= {pr for pr, v, net in ospf_nets[id(r)] if v == sn["vrf"] and ip and ip in net}
                if not procs:
                    continue
                is_silent = False
                for pr in procs:
                    all_, on, off = silent[id(r)].get(pr, (False, set(), set()))
                    if it["name"] in on or (all_ and it["name"] not in off):
                        is_silent = True
                if it["ospf_nettype"] == "p2p":
                    add("I", 2, f"OSPF network-type p2p ({r['device']} {it['name']})")
                elif is_silent:
                    add("L", 1, f"OSPF passif / silent-interface ({r['device']} {it['name']})")
                else:
                    add("I", 1, f"OSPF actif, adjacence possible ({r['device']} {it['name']})")
                break
            peers = sorted({f"{p} ({r['device']})" for r, p, v in bgp_peers
                            if v == sn["vrf"] and _addr(p) in n})
            if peers:
                add("I", 3, f"voisin(s) BGP dans le subnet : {', '.join(peers)}")
            nhs = sorted({f"{s['nh']} ({r['device']})" for r, s in statics
                          if (s["nh_vrf"] or s["vrf"]) == sn["vrf"]
                          and s["nh"] and _addr(s["nh"]) and _addr(s["nh"]) in n})
            if nhs:
                add("I", 1, f"next-hop de routes statiques : {', '.join(nhs)}")
            # --- type d'interface et libellés
            if any(m["it"]["link_mode"] == "route" or "." in m["it"]["name"]
                   or m["it"]["name"].startswith(("Route-Aggregation", "Tunnel")) for m in members):
                add("I", 1, "port routé / sous-interface / Route-Aggregation")
            if INTERCO_WORDS.search(label):
                add("I", 1, f"libellé évoquant une interco (« {label[:50]} »)")
            elif LAN_WORDS.search(label):
                add("L", 1, f"libellé évoquant un LAN (« {label[:50]} »)")

            if inter >= 3 and inter > lan:
                cat = CAT_INTERCO
                conf = "haute" if inter >= 5 and lan == 0 else "moyenne"
            elif lan >= 3 and lan > inter:
                cat = CAT_LAN
                conf = "haute" if lan >= 5 and inter == 0 else "moyenne"
            else:
                cat, conf = CAT_CHECK, "basse"
                if not why:
                    why.append("aucun indice dans les configurations : à qualifier manuellement")
                else:
                    why.append("indices insuffisants ou contradictoires : à qualifier manuellement")
            if cat == CAT_INTERCO and len(devs) == 1:
                why.append("extrémité distante hors du parc analysé")

        # références à mettre à jour lors d'un réadressage
        refs = []
        for r in results:
            for pr, v, net in ospf_nets[id(r)]:
                if v == sn["vrf"] and net.overlaps(n):
                    refs.append(f"OSPF {pr} network {net} ({r['device']})")
            for row in r["routing"]:
                if row[1] == "BGP" and row[6] == "Network annoncé" and row[3] == sn["vrf"]:
                    bn = _net(row[7])
                    if bn and bn.overlaps(n):
                        refs.append(f"BGP network {bn} ({r['device']})")
        for r, s in statics:
            d = _net(s["dest"]) if s["af"] == "IPv4" else None
            if s["vrf"] == sn["vrf"] and d and d.overlaps(n) and d.prefixlen > 0:
                refs.append(f"route statique {s['dest']} ({r['device']})")
            if (s["nh_vrf"] or s["vrf"]) == sn["vrf"] and s["nh"] and _addr(s["nh"]) and _addr(s["nh"]) in n:
                refs.append(f"next-hop {s['nh']} de la route {s['dest']} ({r['device']})")
        for r, p, v in bgp_peers:
            if v == sn["vrf"] and _addr(p) in n:
                refs.append(f"voisin BGP {p} ({r['device']})")
        for (v, srv), users in dhcp_srv.items():
            if v == sn["vrf"] and _addr(srv) and _addr(srv) in n:
                refs.append(f"serveur DHCP relay {srv} utilisé par {', '.join(sorted(users, key=natural_key))}")

        vlans = sorted({v for v in (_svi_vlan(m["it"]["name"]) for m in members) if v is not None})
        shallow = min(members, key=lambda m: (len(m["res"].get("path") or []), m["res"].get("order", 0)))
        plan.append({"vrf": sn["vrf"], "net": n, "cat": cat, "conf": conf, "why": why,
                     "lan": lan, "inter": inter, "devs": devs, "members": members, "vips": sn["vips"],
                     "vlans": vlans, "refs": sorted(set(refs), key=natural_key), "anchor": shallow["res"]})

    plan.sort(key=lambda x: (x["anchor"].get("order", 10 ** 9), CAT_ORDER[x["cat"]],
                             x["vrf"], int(x["net"].network_address)))
    cats = {(x["vrf"], str(x["net"])): (x["cat"], x["conf"]) for x in plan}
    for res in results:
        res["_cat"] = cats
        res["_plan"] = plan
    return plan


PLAN_COLS = ["Équipement de référence", "VRF", "Subnet", "Masque", "Catégorie", "Confiance",
             "Points INTERCO", "Points LAN", "Raisons", "Nb équipements", "Équipements",
             "Adresses (équipement interface : IP)", "VLAN", "VIP VRRP",
             "Références à mettre à jour", "Nouveau subnet", "Commentaire intégrateur"]

READDR_COLS = ["Équipement", "Interface", "Description", "VRF", "Type d'adresse", "IP actuelle",
               "Masque actuel", "Préfixe", "Subnet", "Catégorie", "Confiance",
               "Nouvelle IP", "Nouveau masque"]


def plan_rows(plan, results, depth):
    """Lignes des onglets Plan d'adressage et Réadressage pour les équipements de `results`."""
    mine = {id(r) for r in results}
    prow, rrow = [], []
    for x in plan:
        if not any(id(m["res"]) in mine for m in x["members"]):
            continue
        addrs = "\n".join(f"{m['res']['device']} {m['it']['name']} : {m['ip']}"
                          + (" (secondaire)" if m["kind"] == "secondaire" else "") for m in x["members"])
        prow.append(hierarchy(x["anchor"], depth) + [
            x["anchor"]["device"], x["vrf"], str(x["net"]), str(x["net"].netmask), x["cat"], x["conf"],
            x["inter"], x["lan"], "\n".join(x["why"]), len(x["devs"]), ", ".join(x["devs"]), addrs,
            ", ".join(map(str, x["vlans"])), ", ".join(sorted({v["ip"] for v in x["vips"]})),
            "\n".join(x["refs"]), "", ""])
        for m in x["members"] + x["vips"]:
            if id(m["res"]) not in mine:
                continue
            rrow.append(hierarchy(m["res"], depth) + [
                m["res"]["device"], m["it"]["name"], m["it"]["description"], x["vrf"], m["kind"],
                m["ip"], m["mask"], f"/{x['net'].prefixlen}", str(x["net"]), x["cat"], x["conf"], "", ""])
    rrow.sort(key=lambda r: (results.index(next(x for x in results if x["device"] == r[depth])),
                             natural_key(r[depth + 1]), r[depth + 4] != "primaire"))
    return prow, rrow


CONFIG_EXTS = (".cfg", ".txt", ".log", ".conf")


def node_key(path):
    k = os.path.abspath(path)
    return k.lower() if os.name == "nt" else k


def is_link(path):
    """Lien symbolique ou jonction Windows (mklink /J) : non suivis dans l'arbre."""
    p = Path(path)
    isjunction = getattr(os.path, "isjunction", None)          # Python 3.12+
    if p.is_symlink() or (isjunction and isjunction(p)):
        return True
    # Python < 3.12 sous Windows : une jonction se repère à son chemin réel différent
    return os.path.normcase(os.path.realpath(p)) != os.path.normcase(
        os.path.join(os.path.realpath(p.parent), p.name))


def config_files_in(d):
    try:
        return sorted((f for f in Path(d).iterdir() if f.is_file() and f.suffix.lower() in CONFIG_EXTS),
                      key=lambda f: natural_key(f.name))
    except OSError as e:
        print(f"[ATTENTION] dossier illisible {d} : {e}", file=sys.stderr)
        return []


def new_topo():
    # dirs : clé dossier -> {"path", "parent" (clé dossier ou None), "level", "files"}
    return {"dirs": {}, "file_dir": {}, "warnings": []}


def walk_tree(root, topo, debug=False):
    """
    Parcourt l'arbre des nœuds : chaque dossier = un nœud (sa config .cfg/.txt),
    ses sous-dossiers = ses enfants. Les dossiers cachés (.git...) sont ignorés ;
    les liens symboliques / jonctions sont signalés et non suivis (pas de boucle possible).
    """
    files = []
    root = Path(root)
    root_key = node_key(root)
    for cur, dirnames, _ in os.walk(root, topdown=True, followlinks=False,
                                    onerror=lambda e: print(f"[ATTENTION] dossier illisible : {e}",
                                                            file=sys.stderr)):
        cur_p = Path(cur)
        key = node_key(cur_p)
        if key == root_key:
            parent, level = None, 0
        else:
            parent = node_key(cur_p.parent)
            level = topo["dirs"][parent]["level"] + 1
        keep = []
        for name in sorted(dirnames, key=natural_key):
            sub = cur_p / name
            if name.startswith("."):
                continue
            if is_link(sub):
                msg = f"{sub} est un lien / une jonction : non suivi (représentation en arbre)"
                print(f"[ATTENTION] {msg}", file=sys.stderr)
                topo["warnings"].append(msg)
                continue
            keep.append(name)
        dirnames[:] = keep                       # os.walk ne descend que dans ces dossiers
        cfgs = config_files_in(cur_p)
        if len(cfgs) > 1:
            dbg(debug, f"{len(cfgs)} fichiers de config dans {cur_p} : {', '.join(f.name for f in cfgs)}")
        topo["dirs"][key] = {"path": str(cur_p.resolve()), "parent": parent, "level": level,
                             "files": [node_key(f) for f in cfgs]}
        for f in cfgs:
            topo["file_dir"][node_key(f)] = key
        files.extend(cfgs)
    return files


def expand_inputs(items, recursive=False, topo=None, debug=False):
    """
    Windows (cmd / PowerShell) ne développe pas les jokers : on le fait ici.
    Accepte : fichiers, motifs (configs\\*.cfg) et dossiers (les .cfg/.txt/.log/.conf du dossier ;
    avec recursive=True, tout l'arbre des sous-dossiers, cf. walk_tree).
    """
    topo = topo if topo is not None else new_topo()
    files = []
    for item in items:
        p = Path(item)
        if p.is_dir():
            files.extend(walk_tree(p, topo, debug) if recursive else config_files_in(p))
        elif any(ch in str(item) for ch in "*?["):
            matches = sorted(Path(m) for m in glob.glob(str(item), recursive=recursive))
            if not matches:
                print(f"[ATTENTION] aucun fichier ne correspond à {item}", file=sys.stderr)
            files.extend(m for m in matches if m.is_file())
        elif p.is_file():
            files.append(p)
        else:
            print(f"[ERREUR] fichier introuvable : {item}", file=sys.stderr)
    seen, uniq = set(), []
    for f in files:
        k = node_key(Path(f).resolve())
        if k not in seen:
            seen.add(k)
            uniq.append(f)
    return uniq


def link_tree(results, topo):
    """
    Renseigne pour chaque équipement : nœud parent, enfants et niveau dans l'arbre.
    Un dossier sans config est « transparent » : ses enfants sont rattachés au premier
    ancêtre qui contient une config. Retourne les lignes de l'onglet Arborescence.
    """
    dirs = topo["dirs"]
    devs = {}                                   # clé dossier -> [sysnames]
    for res in results:
        d = topo["file_dir"].get(node_key(res["source"]))
        res["_dir"] = d
        if d:
            devs.setdefault(d, []).append(res["device"])

    def owner(d):
        """Premier dossier (lui-même ou ancêtre) qui contient un équipement."""
        while d is not None and d not in devs:
            d = dirs[d]["parent"]
        return d

    def parent_owner(d):
        return owner(dirs[d]["parent"]) if dirs[d]["parent"] is not None else None

    level, path = {}, {}
    for d in sorted(devs, key=lambda k: dirs[k]["level"]):
        po = parent_owner(d)
        level[d] = 0 if po is None else level[po] + 1
        path[d] = ([] if po is None else path[po] + [", ".join(devs[po])])

    children = {}
    for d in devs:
        po = parent_owner(d)
        if po is not None:
            children.setdefault(po, []).extend(devs[d])

    for res in results:
        d = res["_dir"]
        if d not in devs:
            continue
        po = parent_owner(d)
        res["parent"] = ", ".join(devs[po]) if po else ""
        res["children"] = sorted(children.get(d, []), key=natural_key)
        res["level"] = level[d]
        res["path"] = path[d] + [res["device"]]          # racine -> ... -> équipement

    # lignes de l'onglet Arborescence, dans l'ordre de l'arbre (parcours en profondeur)
    rows_out = []

    def visit(d):
        for dev in devs[d]:
            res = next(r for r in results if r["device"] == dev and r["_dir"] == d)
            res["order"] = len(rows_out)
            rows_out.append([res["level"], dev, res["parent"], ", ".join(res["children"]),
                             "    " * res["level"] + ("└─ " if res["level"] else "") + dev,
                             dirs[d]["path"]])
        kids = sorted((k for k in devs if parent_owner(k) == d),
                      key=lambda k: natural_key(devs[k][0]))
        for k in kids:
            visit(k)

    for d in sorted((k for k in devs if parent_owner(k) is None),
                    key=lambda k: natural_key(devs[k][0])):
        visit(d)
    return rows_out


DEFAULT_TREE_REPORT = "parc_hierarchique.xlsx"


def main():
    # Console Windows : évite les erreurs d'encodage sur les accents
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except AttributeError:
            pass

    ap = argparse.ArgumentParser(
        description="Parse des configs HPE 5940 (Comware 7) : un rapport Excel par équipement")
    ap.add_argument("configs", nargs="+", help="fichier(s), motif(s) ou dossier(s) de configuration")
    ap.add_argument("-r", "--recursive", action="store_true",
                    help="parcourt tous les sous-dossiers (arbre de nœuds : dossier = nœud)")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("-d", "--outdir",
                    help="dossier de sortie (défaut : à côté de chaque fichier de config)")
    ap.add_argument("--merge", metavar="FICHIER.xlsx",
                    help="en plus des rapports individuels, produit un rapport consolidé "
                         "(nom seul = écrit dans le dossier passé au script). "
                         f"Avec -r, il est produit d'office sous le nom {DEFAULT_TREE_REPORT}")
    ap.add_argument("--sans-consolide", action="store_true",
                    help="avec -r : ne pas produire le rapport consolidé hiérarchique")
    ap.add_argument("--open", action="store_true",
                    help="ouvre le dossier des rapports à la fin (Explorateur Windows)")
    ap.add_argument("--debug", action="store_true",
                    help="affiche le détail de la lecture de chaque fichier d'entrée")
    args = ap.parse_args()

    if args.debug:
        print(f"[DEBUG] hpe5940_parser {__version__}")
        print(f"[DEBUG] Python        : {sys.version.split()[0]}  ({sys.executable})")
        print(f"[DEBUG] Script        : {Path(__file__).resolve()}")
        print(f"[DEBUG] Dossier courant : {Path.cwd()}")
        print(f"[DEBUG] Arguments reçus : {args.configs}")
    topo = new_topo()
    files = expand_inputs(args.configs, recursive=args.recursive, topo=topo, debug=args.debug)
    if args.debug:
        print(f"[DEBUG] Fichiers retenus ({len(files)}) :")
        for f in files:
            print(f"[DEBUG]   {Path(f).resolve()}")
    if not files:
        sys.exit("Aucun fichier de configuration trouvé.")
    fixed_outdir = Path(args.outdir).resolve() if args.outdir else None

    # 1) lecture de toutes les configs (nécessaire pour connaître parents / enfants)
    parsed, skipped, errors = [], [], 0
    for f in files:
        try:
            if args.debug:
                print(f"\n--- {Path(f).name} ---")
            res = parse_config(f, debug=args.debug)
        except Exception as e:  # on continue sur les autres fichiers
            print(f"[ERREUR] {f}: {e}", file=sys.stderr)
            if args.debug:
                import traceback
                traceback.print_exc()
            errors += 1
            continue
        if not res["sysname_found"] and not res["interfaces"]:
            print(f"[IGNORÉ] {f} : pas une configuration Comware (ni sysname ni interface)")
            skipped.append(f)
            continue
        parsed.append((f, res))
    tree = link_tree([r for _, r in parsed], topo) if topo["dirs"] else []
    if parsed:
        # classification sur tout le parc : indices croisés entre équipements
        plan = classify_subnets([r for _, r in parsed])
        counts = {}
        for x in plan:
            counts[x["cat"]] = counts.get(x["cat"], 0) + 1
        print("Subnets IPv4 : " + ", ".join(f"{k}={counts[k]}" for k in sorted(counts, key=CAT_ORDER.get)))

    # 2) écriture d'un rapport par équipement, dans le dossier de sa config
    results, written = [], []
    for f, res in parsed:
        outdir = fixed_outdir or Path(f).resolve().parent
        outdir.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^\w.-]+", "_", res["device"]) or Path(f).stem
        out = outdir / f"{safe}.xlsx"
        if out in written:                    # deux configs avec le même sysname
            out = outdir / f"{safe}_{Path(f).stem}.xlsx"
            print(f"[ATTENTION] sysname {res['device']} déjà traité : rapport renommé {out.name}",
                  file=sys.stderr)
        try:
            stats = write_report([res], out)
        except PermissionError:
            print(f"[ERREUR] {out} est verrouillé (ouvert dans Excel ?) - fermez-le et relancez",
                  file=sys.stderr)
            errors += 1
            continue
        except Exception as e:
            print(f"[ERREUR] écriture de {out} impossible : {e}", file=sys.stderr)
            if args.debug:
                import traceback
                traceback.print_exc()
            errors += 1
            continue
        results.append(res)
        written.append(out)
        dbg(args.debug, f"Rapport écrit : {out}  (existe : {out.exists()}, {out.stat().st_size} octets)")
        print(f"[OK] {Path(f).name} -> {out}")
        print("     " + "  ".join(f"{k}={v}" for k, v in stats.items()))

    merge_name = args.merge
    if not merge_name and args.recursive and not args.sans_consolide:
        merge_name = DEFAULT_TREE_REPORT            # -r : fichier hiérarchique produit d'office
    if merge_name and not results:
        print(f"[ATTENTION] rapport consolidé non produit : aucune configuration valide traitée",
              file=sys.stderr)
    if merge_name and results:
        merge_path = Path(merge_name)
        if not merge_path.is_absolute():
            # nom seul : -d s'il est donné, sinon le dossier passé au script (racine de l'arbre),
            # sinon le dossier du premier rapport
            first_dir = next((Path(c) for c in args.configs if Path(c).is_dir()), None)
            base = fixed_outdir or (first_dir.resolve() if first_dir else written[0].parent)
            merge_path = base / merge_path
        try:
            write_report(results, merge_path, tree=tree)
            written.append(merge_path.resolve())
            print(f"[OK] Rapport consolidé{' hiérarchique' if tree else ''} -> {merge_path.resolve()}")
        except PermissionError:
            print(f"[ERREUR] {merge_path} est verrouillé (ouvert dans Excel ?)", file=sys.stderr)
        except Exception as e:
            print(f"[ERREUR] écriture de {merge_path} impossible : {e}", file=sys.stderr)

    print(f"\n{len(results)} rapport(s) généré(s), {len(skipped)} fichier(s) ignoré(s), "
          f"{errors} erreur(s)")
    if args.recursive:
        depth = max((r[0] for r in tree), default=0)
        print(f"Arbre : {len(topo['dirs'])} dossier(s) parcouru(s), {len(tree)} nœud(s), "
              f"profondeur {depth}, {len(topo['warnings'])} lien(s)/jonction(s) non suivi(s)")
    merged = [w for w in written if merge_name and w.name == Path(merge_name).name]
    folders = sorted({str(w.parent) for w in written if w not in merged})
    if len(folders) <= 5:
        for d in folders:
            print(f"Dossier des rapports : {d}")
    else:
        print(f"Rapports écrits dans {len(folders)} dossiers (chacun à côté de sa config)")
    if merged:
        print(f"Fichier consolidé : {merged[-1]}")
    if args.open and folders:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(merged[-1].parent) if merged else folders[0])
            elif sys.platform == "darwin":
                subprocess.run(["open", str(merged[-1].parent) if merged else folders[0]])
            else:
                subprocess.run(["xdg-open", str(merged[-1].parent) if merged else folders[0]])
        except Exception as e:
            print(f"[ATTENTION] impossible d'ouvrir le dossier : {e}", file=sys.stderr)
    sys.exit(1 if not results else 0)


if __name__ == "__main__":
    main()
