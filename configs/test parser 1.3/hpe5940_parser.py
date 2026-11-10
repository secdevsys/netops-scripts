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

Usage (Windows : "py" ou "python") :
  py hpe5940_parser.py config.cfg                         -> <sysname>_rapport.xlsx (à côté de config.cfg)
  py hpe5940_parser.py configs\\*.cfg -d rapports        -> un rapport par équipement
  py hpe5940_parser.py C:\\Backups\\5940 -d rapports      -> tous les .cfg/.txt/.log d'un dossier
  py hpe5940_parser.py configs -d rapports --merge tous.xlsx -> + un rapport consolidé (optionnel)
  py hpe5940_parser.py SW1.cfg --debug                    -> détail de la lecture (dépannage)
Les jokers (*.cfg) sont développés par le script : fonctionne sous cmd et PowerShell.

Entrée : sortie "display current-configuration" ou fichier startup.cfg.
Dépendance : openpyxl  (pip install openpyxl)
"""

__version__ = "1.3.0"

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
def clean_lines(text):
    """Nettoie une capture CLI (prompts, ---- More ----, caractères de contrôle)."""
    lines = []
    for raw in text.splitlines():
        line = re.sub(r"\s*-+\s*More\s*-+\s*", "", raw)
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
           "bgp_peers": [], "bgp_as": "", "statics": []}

    for b in blocks:
        if b["head"].startswith("sysname "):
            res["device"] = b["head"][len("sysname "):].strip()
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
           "Adresse IP", "Masque", "Préfixe", "Réseau", "IP secondaires", "IPv6",
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
                prim.get("prefix", ""), prim.get("network", ""), secs,
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


def write_sheet(wb, title, cols, rows, status_col=None):
    ws = wb.create_sheet(title)
    ws.append(cols)
    for c in ws[1]:
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
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
            longest = max((len(x) for x in str(c.value or "").split("\n")), default=0)
            widths[i] = max(widths[i], longest)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 8), 60)
    ws.freeze_panes = "C2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(rows) + 1}"
    ws.row_dimensions[1].height = 32
    return ws


def write_report(results, out):
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthèse"
    ws.append(["Équipement", "Version", "Router-ID", "AS BGP", "VRF (RD)", "Nb VLANs",
               "Interfaces L2", "Interfaces L3", "Interfaces shutdown", "Process OSPF",
               "Voisins BGP", "Routes statiques"])
    all_rows = [[] for _ in range(6)]
    for res in results:
        rows = build_rows(res)
        for acc, r in zip(all_rows, rows):
            acc.extend(r)
        l2, l3, rt, bgp, st, vl = rows
        ospf_procs = sorted({r[2] for r in rt if r[1] == "OSPF"}, key=natural_key)
        append_row(ws, [res["device"], res["version"], res["router_id"], res["bgp_as"],
                   ", ".join(f"{k} ({v})" if v else k for k, v in res["vrfs"].items()),
                   len(res["vlans"]), len(l2), len(l3),
                   sum(1 for r in l2 if r[4] == "shutdown") + sum(1 for r in l3 if r[4] == "shutdown"),
                   ", ".join(ospf_procs), sum(1 for p in bgp if p[4] == "peer"), len(st)])
    for c in ws[1]:
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for i in range(1, 13):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["E"].width = 34
    ws.freeze_panes = "B2"

    l2, l3, rt, bgp, st, vl = all_rows
    write_sheet(wb, "Interfaces L2", L2_COLS, l2, status_col=4)
    write_sheet(wb, "Interfaces L3", L3_COLS, l3, status_col=4)
    write_sheet(wb, "Routage dynamique", RT_COLS, rt)
    write_sheet(wb, "Voisins BGP", BGP_COLS, bgp)
    write_sheet(wb, "Routes statiques", ST_COLS, st)
    write_sheet(wb, "VLANs", VLAN_COLS, vl)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return {"L2": len(l2), "L3": len(l3), "Routage": len(rt), "BGP": len(bgp),
            "Statiques": len(st), "VLANs": len(vl)}


CONFIG_EXTS = (".cfg", ".txt", ".log", ".conf")


def expand_inputs(items):
    """
    Windows (cmd / PowerShell) ne développe pas les jokers : on le fait ici.
    Accepte : fichiers, motifs (configs\\*.cfg) et dossiers (tous les .cfg/.txt/.log/.conf).
    """
    files = []
    for item in items:
        p = Path(item)
        if p.is_dir():
            files.extend(sorted(f for f in p.iterdir()
                                if f.is_file() and f.suffix.lower() in CONFIG_EXTS))
        elif any(ch in item for ch in "*?["):
            matches = sorted(Path(m) for m in glob.glob(item))
            if not matches:
                print(f"[ATTENTION] aucun fichier ne correspond à {item}", file=sys.stderr)
            files.extend(m for m in matches if m.is_file())
        elif p.is_file():
            files.append(p)
        else:
            print(f"[ERREUR] fichier introuvable : {item}", file=sys.stderr)
    seen, uniq = set(), []
    for f in files:
        k = str(f.resolve()).lower()
        if k not in seen:
            seen.add(k)
            uniq.append(f)
    return uniq


def main():
    # Console Windows : évite les erreurs d'encodage sur les accents
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except AttributeError:
            pass

    ap = argparse.ArgumentParser(
        description="Parse des configs HPE 5940 (Comware 7) : un rapport Excel par équipement")
    ap.add_argument("configs", nargs="+", help="fichier(s) de configuration")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("-d", "--outdir",
                    help="dossier de sortie (défaut : à côté de chaque fichier de config)")
    ap.add_argument("--merge", metavar="FICHIER.xlsx",
                    help="en plus des rapports individuels, produit un rapport consolidé")
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
    files = expand_inputs(args.configs)
    if args.debug:
        print(f"[DEBUG] Fichiers retenus ({len(files)}) :")
        for f in files:
            print(f"[DEBUG]   {Path(f).resolve()}")
    if not files:
        sys.exit("Aucun fichier de configuration trouvé.")
    fixed_outdir = Path(args.outdir).resolve() if args.outdir else None
    results, written, errors = [], [], 0
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
        outdir = fixed_outdir or Path(f).resolve().parent
        outdir.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^\w.-]+", "_", res["device"]) or Path(f).stem
        out = outdir / f"{safe}_rapport.xlsx"
        if out in written:                    # deux configs avec le même sysname
            out = outdir / f"{safe}_{Path(f).stem}_rapport.xlsx"
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

    if args.merge and results:
        merge_path = Path(args.merge)
        if not merge_path.is_absolute():
            merge_path = (fixed_outdir or written[0].parent) / merge_path
        try:
            write_report(results, merge_path)
            written.append(merge_path.resolve())
            print(f"[OK] Rapport consolidé -> {merge_path.resolve()}")
        except PermissionError:
            print(f"[ERREUR] {merge_path} est verrouillé (ouvert dans Excel ?)", file=sys.stderr)
        except Exception as e:
            print(f"[ERREUR] écriture de {merge_path} impossible : {e}", file=sys.stderr)

    print(f"\n{len(results)} rapport(s) généré(s), {errors} erreur(s)")
    folders = sorted({str(w.parent) for w in written})
    for d in folders:
        print(f"Dossier des rapports : {d}")
    if args.open and folders:
        try:
            if sys.platform.startswith("win"):
                os.startfile(folders[0])
            elif sys.platform == "darwin":
                subprocess.run(["open", folders[0]])
            else:
                subprocess.run(["xdg-open", folders[0]])
        except Exception as e:
            print(f"[ATTENTION] impossible d'ouvrir le dossier : {e}", file=sys.stderr)
    sys.exit(1 if not results else 0)


if __name__ == "__main__":
    main()
