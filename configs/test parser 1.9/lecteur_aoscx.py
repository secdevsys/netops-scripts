#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lecteur_aoscx.py - Lecture des configurations ArubaOS-CX pour hpe5940_parser.py

Ce module est appelé automatiquement par hpe5940_parser.py quand un fichier est reconnu
comme une configuration ArubaOS-CX (show running-config). Il remplit EXACTEMENT la même
structure de données que la lecture Comware : tous les onglets du rapport, l'arbre, la
classification INTERCO/LAN et allouer_adresses.py fonctionnent donc sans changement.

Les noms d'interfaces AOS-CX sont conservés dans le rapport (1/1/1, lag 1, vlan 10...).
Un nom « équivalent Comware » interne (cw_name) sert uniquement à reconnaître le type
d'interface : vlan 10 -> SVI, loopback 0 -> Loopback, mgmt -> management, lag -> agrégat.

Correspondances principales
  hostname                          -> sysname
  vlan access / trunk native / trunk allowed   -> mode, VLAN access, PVID, VLANs autorisées
  routing / no routing / ip address -> port L3 / L2 (aucune valeur par défaut supposée :
                                       elle dépend de la plateforme)
  lag N (sur un port)               -> membre de l'agrégat "lag N" (config L2 héritée)
  interface lag N multi-chassis     -> agrégat VSX (noté dans « Autres commandes »)
  vrf attach X                      -> VRF
  active-gateway ip X / vrrp N      -> passerelle (comme VRRP)
  ip helper-address X               -> DHCP relay
  ip ospf N area A, ip ospf passive, passive-interface default
                                    -> OSPF de l'interface et interfaces passives
  router ospf / router bgp / ip route -> onglets Routage, Voisins BGP, Routes statiques

Écrit à partir de la documentation AOS-CX 10.x ; à valider sur des configurations réelles.
"""

__version__ = "1.0.0"

import re
from collections import OrderedDict

DETECT_RE = re.compile(r"^(!Version ArubaOS-CX|interface (lag|vlan|loopback) \d|interface \d+/\d+/\d+|"
                       r"interface mgmt$|vsx$)", re.M)


def is_aoscx(text):
    """Vrai si le texte ressemble à une configuration ArubaOS-CX."""
    if "ArubaOS-CX" in text[:2000]:
        return True
    hits = len(DETECT_RE.findall(text))
    return hits >= 2 and "sysname " not in text


# --------------------------------------------------------------------------
# Découpage en blocs (indentation de 4 espaces, lignes « ! » = commentaires)
# --------------------------------------------------------------------------
def blocks_of(lines):
    """[(en-tête, [(niveau, texte)])] ; niveau 1 = directement sous l'en-tête."""
    blocks, cur = [], None
    for line in lines:
        text = line.strip()
        if not text or text.startswith("!") or text in ("exit", "end"):
            continue
        indent = len(line) - len(line.lstrip(" "))
        if indent == 0:
            cur = (text, [])
            blocks.append(cur)
        elif cur is not None:
            cur[1].append((max(1, round(indent / 4)), text))
    return blocks


def vlan_list(text, P):
    """'10,20,30-35' / 'all' -> ensemble de VLANs."""
    text = text.strip()
    if text == "all":
        return set(P.ALL_VLANS)
    out = set()
    for part in text.replace(" ", ",").split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            if a.isdigit() and b.isdigit():
                out.update(range(int(a), int(b) + 1))
        elif part.isdigit():
            out.add(int(part))
    return out


def ip_prefix(cidr, P):
    """'10.1.1.1/24' -> (ip, masque, /24, réseau)"""
    if "/" not in cidr:
        return cidr, "", "", ""
    ip, plen = cidr.split("/", 1)
    return P.ip_info(ip, plen)


# --------------------------------------------------------------------------
# Interfaces
# --------------------------------------------------------------------------
def cw_name(name, routed):
    """Nom équivalent Comware, utilisé seulement pour reconnaître le type d'interface."""
    m = re.match(r"(vlan|loopback|lag|tunnel) ?(\d+)$", name)
    if m:
        kind, n = m.groups()
        return {"vlan": f"Vlan-interface{n}", "loopback": f"LoopBack{n}", "tunnel": f"Tunnel{n}",
                "lag": f"{'Route' if routed else 'Bridge'}-Aggregation{n}"}[kind]
    if name == "mgmt":
        return "M-GigabitEthernet0"
    return None


def parse_interface(name, body, P):
    it = P.new_interface(name)
    it["trunk_vlans"] = set()
    it["link_type"] = None
    routing = None              # True / False si écrit explicitement
    trunk_allowed = None
    vrrp_id, vrrp_af = None, None
    for level, line in body:
        tok = line.split()
        low = line.lower()
        # sous-bloc vrrp N address-family ipv4
        if level >= 2 and vrrp_id is not None:
            if tok[0] == "address" and len(tok) >= 2 and vrrp_af == "ipv4":
                it["vrrp"][vrrp_id]["vip"].append(tok[1])
            elif tok[0] == "priority" and len(tok) >= 2:
                it["vrrp"][vrrp_id]["prio"] = tok[1]
            else:
                it["others"].append(P.mask(f"vrrp {vrrp_id} : {line}"))
            continue
        vrrp_id = None
        if line.startswith("description "):
            it["description"] = line[len("description "):]
        elif line == "shutdown":
            it["shutdown"] = True
        elif line == "no shutdown":
            pass
        elif line == "routing":
            routing = True
        elif line == "no routing":
            routing = False
        elif line.startswith("vlan access ") and len(tok) >= 3:
            it["link_type"], it["access_vlan"] = "access", tok[2]
            routing = False if routing is None else routing
        elif line.startswith("vlan trunk native ") and len(tok) >= 4:
            it["link_type"], it["pvid"] = "trunk", tok[3]
            routing = False if routing is None else routing
        elif line.startswith("vlan trunk allowed ") and len(tok) >= 4:
            it["link_type"] = "trunk"
            routing = False if routing is None else routing
            trunk_allowed = (trunk_allowed or set()) | vlan_list(" ".join(tok[3:]), P)
        elif line.startswith("lag ") and len(tok) == 2 and tok[1].isdigit():
            it["agg_group"] = tok[1]
        elif line.startswith("lacp mode "):
            it["lacp_mode"] = tok[2]
        elif tok[0] == "mtu" and len(tok) > 1:
            it["mtu"] = tok[1]
        elif tok[0] == "speed" and len(tok) > 1:
            it["speed"] = " ".join(tok[1:])
        elif line.startswith("spanning-tree port-type admin-edge") or line == "spanning-tree port-type edge":
            it["stp_edge"] = True
        elif line.startswith("vrf attach ") and len(tok) >= 3:
            it["vrf"] = tok[2]
        elif line.startswith("ip address ") and len(tok) >= 3:
            ip, netmask, pfx, net = ip_prefix(tok[2], P)
            it["ipv4"].append({"ip": ip, "mask": netmask, "prefix": pfx, "network": net,
                               "sub": "secondary" in tok[3:]})
        elif line.startswith("ip static ") and len(tok) >= 3:            # interface mgmt
            ip, netmask, pfx, net = ip_prefix(tok[2], P)
            it["ipv4"].append({"ip": ip, "mask": netmask, "prefix": pfx, "network": net, "sub": False})
        elif line in ("ip dhcp", "ip address dhcp"):
            it["ipv4"].append({"ip": "dhcp", "mask": "", "prefix": "", "network": "", "sub": False})
        elif line.startswith("ipv6 address ") and len(tok) >= 3:
            it["ipv6"].append(tok[2])
        elif line.startswith("ip helper-address ") and len(tok) >= 3:
            it["dhcp_mode"] = "relay"
            it["dhcp_servers"].append(tok[2])
        elif line.startswith("ip mtu ") and len(tok) >= 3:
            it["ip_mtu"] = tok[2]
        elif line.startswith("active-gateway ip ") and len(tok) >= 3 and tok[2] != "mac":
            it["vrrp"].setdefault("active-gateway", {"vip": [], "prio": ""})["vip"].append(tok[2])
        elif line.startswith("active-gateway ip mac"):
            pass
        elif re.match(r"^vrrp \d+ address-family ", line):
            vrrp_id, vrrp_af = tok[1], tok[3]
            it["vrrp"].setdefault(vrrp_id, {"vip": [], "prio": ""})
        elif re.match(r"^ip ospf \d+ area ", line):
            it["ospf_proc"], it["ospf_area"] = tok[2], tok[4]
        elif line.startswith("ip ospf cost ") and len(tok) >= 4:
            it["ospf_cost"] = tok[3]
        elif low.startswith("ip ospf network point-to-point"):
            it["ospf_nettype"] = "p2p"
        elif line.startswith("ip ospf network ") and len(tok) >= 4:
            it["ospf_nettype"] = tok[3]
        elif line.startswith("ip ospf authentication") or line.startswith("ip ospf message-digest-key"):
            it["ospf_auth"] = True
        elif line == "ip ospf passive":
            it["_ospf_passive"] = True
        elif line == "no ip ospf passive":
            it["_ospf_passive"] = False
        elif line.startswith("encapsulation dot1q ") and len(tok) >= 3:      # sous-interface
            it["dot1q"] = tok[2]
        else:
            it["others"].append(P.mask(line))

    if trunk_allowed is not None:
        it["trunk_vlans"] = trunk_allowed
    elif it["link_type"] == "trunk":
        it["trunk_vlans"] = set(P.ALL_VLANS)      # trunk sans « allowed » : toutes les VLANs
    if it["ipv4"] or it["ipv6"]:
        routing = True if routing is None else routing
    it["_routing"] = routing
    it["_explicit_l2"] = it["link_type"] is not None
    if it["link_type"] is None:
        it["link_type"] = "access"
    cw = cw_name(name, bool(routing))
    if cw:
        it["cw_name"] = cw
    # « port routé » : seulement pour les ports physiques, sous-interfaces et LAG
    # (une SVI, une loopback ou mgmt sont L3 par nature, comme sous Comware)
    if not cw or cw.startswith(("Bridge-Aggregation", "Route-Aggregation")):
        it["link_mode"] = "route" if routing else ("bridge" if routing is False else None)
    return it


# --------------------------------------------------------------------------
# Routage
# --------------------------------------------------------------------------
def parse_router_ospf(head, body, dev, P):
    tok = head.split()
    proc = tok[2] if len(tok) > 2 else "1"
    vrf = tok[tok.index("vrf") + 1] if "vrf" in tok else ""
    rid = next((l.split()[1] for _, l in body if l.startswith("router-id ")), "")
    rows, passive_default = [], False

    def add(area, cat, val):
        rows.append([dev, "OSPF", proc, vrf or "public", rid, area, cat, val])

    for _, line in body:
        t = line.split()
        if line.startswith("router-id "):
            continue
        if t[0] == "area" and len(t) >= 2:
            add(t[1], "Zone", "déclarée")
            if len(t) > 2:
                add(t[1], "Type de zone", " ".join(t[2:]))
        elif line == "passive-interface default":
            passive_default = True
            add("", "Silent-interface", "silent-interface all")
        elif t[0] == "redistribute":
            add("", "Redistribution", " ".join(t[1:]))
        else:
            add("", "Paramètre", P.mask(line))
    return proc, vrf, rid, rows, passive_default


def parse_router_bgp(head, body, dev, P):
    tok = head.split()
    local_as = tok[2] if len(tok) > 2 else ""
    rows, peers = [], OrderedDict()
    vrf, af, af_level, rid = "", "", 0, ""

    def peer(pid):
        key = (vrf, pid)
        if key not in peers:
            peers[key] = {"vrf": vrf or "public", "id": pid, "kind": "peer", "group": "",
                          "as": "", "group_type": "", "desc": "", "conn_if": "", "auth": False,
                          "bfd": False, "multihop": "", "afs": [], "rp_in": [], "rp_out": [],
                          "others": []}
        return peers[key]

    def add(cat, val):
        rows.append([dev, "BGP", local_as, vrf or "public", rid, af, cat, val])

    for level, line in body:
        t = line.split()
        if af and level <= af_level:          # sortie de l'address-family
            af = ""
        if level == 1:                        # retour au niveau « router bgp »
            vrf = ""
        if level == 1 and t[0] == "vrf" and len(t) >= 2:
            vrf, af = t[1], ""
            add("VRF", f"vrf {vrf}")
            continue
        if line.startswith("address-family "):
            af, af_level = line[len("address-family "):], level
            add("Famille d'adresses", af)
            continue
        if line == "exit-address-family":
            af = ""
            continue
        if line.startswith("bgp router-id ") and len(t) >= 3:
            rid = t[2]
            add("Router-ID", rid)
        elif t[0] == "neighbor" and len(t) >= 3:
            p = peer(t[1])
            sub, rest = t[2], t[3:]
            if sub == "remote-as" and rest:
                p["as"] = rest[0]
            elif sub == "peer-group" and not rest:
                p["kind"], p["group_type"] = "group", "internal"
            elif sub == "peer-group" and rest:
                p["group"] = rest[0]
            elif sub == "description":
                p["desc"] = " ".join(rest)
            elif sub == "update-source" and rest:
                p["conn_if"] = " ".join(rest)
            elif sub == "password":
                p["auth"] = True
            elif sub in ("fall-over", "bfd") or (sub == "fall-over" and "bfd" in rest):
                p["bfd"] = True
            elif sub == "ebgp-multihop":
                p["multihop"] = rest[0] if rest else "255"
            elif sub == "activate":
                p["afs"].append(af or "ipv4 unicast")
            elif sub == "route-map" and len(rest) >= 2:
                (p["rp_in"] if rest[1] == "in" else p["rp_out"]).append(rest[0])
            else:
                p["others"].append(P.mask(f"[{af}] {' '.join(t[2:])}" if af else " ".join(t[2:])))
        elif t[0] == "network" and len(t) >= 2:
            add("Network annoncé", ip_prefix(t[1], P)[3] or t[1])
        elif t[0] == "redistribute":
            add("Redistribution", " ".join(t[1:]))
        else:
            add("Paramètre", P.mask(line))
    # héritage des peer-groups
    for (v, _), p in peers.items():
        if p["kind"] == "peer" and p["group"]:
            g = peers.get((v, p["group"]))
            if g:
                for k in ("as", "conn_if", "multihop"):
                    p[k] = p[k] or g[k]
                p["auth"] = p["auth"] or g["auth"]
                p["bfd"] = p["bfd"] or g["bfd"]
                p["afs"] = p["afs"] or [f"{a} (groupe)" for a in g["afs"]]
                p["rp_in"] = p["rp_in"] or g["rp_in"]
                p["rp_out"] = p["rp_out"] or g["rp_out"]
    for p in peers.values():
        if p["kind"] == "group" and p["as"] and p["as"] != local_as:
            p["group_type"] = "external"
    return local_as, rid, rows, peers


def parse_ip_route(line, dev):
    """ip route 10.0.0.0/8 10.1.1.1 [distance N] [tag N] [vrf V] / ... blackhole / ... 1/1/1"""
    t = line.split()
    v6 = t[0] == "ipv6"
    t = t[2:]
    if not t:
        return None
    r = {"device": dev, "af": "IPv6" if v6 else "IPv4", "vrf": "public", "dest": t[0],
         "nh_vrf": "", "out_if": "", "nh": "", "pref": "", "tag": "", "desc": "", "others": []}
    if not v6 and "/" in t[0]:
        import ipaddress
        try:
            n = ipaddress.ip_network(t[0], strict=False)
            r["dest"] = f"{n.network_address}/{n.prefixlen}"
        except ValueError:
            pass
    i, rest = 0, t[1:]
    while i < len(rest):
        k = rest[i]
        nxt = rest[i + 1] if i + 1 < len(rest) else ""
        if k == "vrf" and nxt:
            r["vrf"] = nxt
            i += 2
        elif k == "distance" and nxt:
            r["pref"] = nxt
            i += 2
        elif k == "tag" and nxt:
            r["tag"] = nxt
            i += 2
        elif k in ("blackhole", "reject", "nullroute"):
            r["out_if"] = k
            i += 1
        elif re.match(r"^\d{1,3}(\.\d{1,3}){3}$", k) or (v6 and ":" in k):
            r["nh"] = k
            i += 1
        elif re.match(r"^\d+/\d+/\d+", k) or k in ("lag", "vlan", "loopback"):
            r["out_if"] = f"{k} {nxt}" if k in ("lag", "vlan", "loopback") else k
            i += 2 if k in ("lag", "vlan", "loopback") else 1
        else:
            r["others"].append(k)
            i += 1
    return r


# --------------------------------------------------------------------------
# Point d'entrée appelé par hpe5940_parser.parse_config
# --------------------------------------------------------------------------
PAGER_RE = re.compile(r"-- MORE --, next page: Space, next line: Enter, quit: Control-C"
                      r"(?:[ \t\x08]|\x1b\[[0-9;]*[A-Za-z])*", re.I)


def parse(text, res, P, debug=False):
    text = PAGER_RE.sub("", text)            # pagination AOS-CX dans une capture terminal
    lines = P.clean_lines(text)
    # la ligne « !Version ArubaOS-CX ... » est un commentaire : on la lit avant de l'écarter
    m = re.search(r"^!\s*Version\s+ArubaOS-CX\s+(\S+)", text, re.M)
    if m:
        res["version"] = f"ArubaOS-CX {m.group(1)}"
    else:
        res["version"] = "ArubaOS-CX"
    blocks = blocks_of(lines)
    P.dbg(debug, f"OS détecté  : ArubaOS-CX ({len(blocks)} blocs)")
    for head, _ in blocks:
        if head.startswith("hostname "):
            res["device"] = head.split(None, 1)[1].strip()
            res["sysname_found"] = True
    dev = res["device"]
    ifs = res["interfaces"]
    passive_default = {}

    for head, body in blocks:
        tok = head.split()
        key = tok[0]
        if key == "interface" and len(tok) >= 2:
            name = " ".join(tok[1:])
            extra = ""
            if name.endswith(" multi-chassis"):
                name, extra = name[: -len(" multi-chassis")], "multi-chassis (VSX)"
            it = parse_interface(name, body, P)
            if extra:
                it["others"].insert(0, extra)
            ifs[name] = it
        elif key == "vlan" and len(tok) >= 2:
            ids = vlan_list(" ".join(tok[1:]), P)
            name = next((l[5:] for _, l in body if l.startswith("name ")), "")
            desc = next((l[12:] for _, l in body if l.startswith("description ")), "")
            for v in sorted(ids):
                res["vlans"][v] = {"name": name, "desc": desc}
        elif key == "vrf" and len(tok) >= 2:
            rd = next((l.split()[1] for _, l in body if l.startswith("rd ")), "")
            res["vrfs"][tok[1]] = rd
        elif head.startswith("router ospf"):
            proc, vrf, rid, rows, pdef = parse_router_ospf(head, body, dev, P)
            res["routing"].extend(rows)
            passive_default[(proc, vrf or "public")] = pdef
            if rid and not res["router_id"]:
                res["router_id"] = rid
        elif head.startswith("router bgp"):
            las, rid, rows, peers = parse_router_bgp(head, body, dev, P)
            res["bgp_as"] = las
            res["routing"].extend(rows)
            for p in peers.values():
                p["local_as"] = las
                res["bgp_peers"].append(p)
            if rid and not res["router_id"]:
                res["router_id"] = rid
        elif head.startswith(("ip route ", "ipv6 route ")):
            r = parse_ip_route(head, dev)
            if r:
                res["statics"].append(r)
        elif head.startswith("router-id ") and len(tok) >= 2:
            res["router_id"] = tok[1]

    # agrégats : membres + héritage de la config L2 du LAG
    for it in ifs.values():
        it["members"] = []
    for it in ifs.values():
        g = it["agg_group"]
        if not g:
            continue
        agg = f"lag {g}"
        it["agg_name"] = agg
        if agg in ifs:
            parent = ifs[agg]
            parent["members"].append(it["name"])
            it["agg_l2"] = not parent.get("_routing")
            if it["agg_l2"] and not it["_explicit_l2"]:
                for k in ("link_type", "access_vlan", "pvid", "trunk_vlans", "hyb_tagged", "hyb_untagged"):
                    it[k] = parent[k]
                it["inherited"] = True
            if parent.get("_routing"):
                it["link_mode"] = "route"

    # OSPF : interfaces passives -> mêmes lignes « Silent-interface » que Comware
    for it in ifs.values():
        if not it["ospf_proc"]:
            continue
        vrf = it["vrf"] or "public"
        pdef = passive_default.get((it["ospf_proc"], vrf), False)
        flag = it.get("_ospf_passive")
        if pdef and flag is False:
            res["routing"].append([dev, "OSPF", it["ospf_proc"], vrf, "", "", "Silent-interface",
                                   f"undo silent-interface {it['name']}"])
        elif not pdef and flag:
            res["routing"].append([dev, "OSPF", it["ospf_proc"], vrf, "", "", "Silent-interface",
                                   f"silent-interface {it['name']}"])
        det = [it["name"]] + ([f"cost {it['ospf_cost']}"] if it["ospf_cost"] else []) + \
              ([f"network-type {it['ospf_nettype']}"] if it["ospf_nettype"] else []) + \
              (["auth"] if it["ospf_auth"] else [])
        res["routing"].append([dev, "OSPF", it["ospf_proc"], vrf, "", it["ospf_area"], "Interface",
                               " | ".join(det)])
    P.dbg(debug, f"sysname      : {res['device']}  ({res['version']})")
    P.dbg(debug, f"Trouvé       : {len(ifs)} interfaces, {len(res['vlans'])} VLANs, "
                 f"{len(res['bgp_peers'])} voisins/groupes BGP, {len(res['statics'])} routes statiques")
    return res
