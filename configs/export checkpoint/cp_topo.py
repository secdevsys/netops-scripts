#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cp_topo.py - Analyse d'un export HTML de configuration Check Point (cluster)
             et génération de la topologie au format Visio (.vsdx).

Aucune dépendance externe : Python 3.8+ (bibliothèque standard uniquement).

Usage :
    python cp_topo.py export.html                 -> export_topologie.vsdx + rapport
    python cp_topo.py dossier_export/ -o topo.vsdx
    python cp_topo.py export.html --json modele.json --report rapport.md
    python cp_topo.py export.html --debug         -> liste les tableaux détectés
    python cp_topo.py export.html --cluster CL-FW-PARIS

Principe :
  1. Tous les tableaux HTML sont extraits avec leur contexte (titres h1-h6,
     légendes, lignes-titres internes aux tableaux).
  2. Les colonnes sont reconnues par alias (Name / IPv4 Address / Net Mask /
     Topology / Member / Destination / Next hop ...), en anglais et français.
  3. Les objets (cluster, membres, passerelles), leurs interfaces, les membres
     du cluster et les routes statiques sont reconstitués.
  4. La topologie est déduite : réseaux des VIP, réseaux de synchro/gestion
     propres aux membres, routeurs de prochain saut.
  5. Un fichier .vsdx (une page par cluster) + un rapport Markdown + un JSON
     sont produits. Les alias se complètent facilement dans FIELD_ALIASES.
"""

import argparse
import html
import ipaddress
import json
import os
import re
import sys
import zipfile
from collections import OrderedDict, deque
from datetime import datetime, timezone
from html.parser import HTMLParser

# --------------------------------------------------------------------------
# 1. Extraction des tableaux HTML
# --------------------------------------------------------------------------

HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}


class TableExtractor(HTMLParser):
    """Extrait chaque <table> avec : lignes, titre courant, légende, contexte."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.stack = []
        self.last_heading = ""
        self.recent = deque(maxlen=4)
        self._heading_buf = None
        self._caption_buf = None
        self._skip = 0  # <script>/<style>

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self._skip += 1
        elif tag in HEADING_TAGS and not self.stack:
            self._heading_buf = []
        elif tag == "table":
            self.stack.append({"rows": [], "heading": self.last_heading,
                               "context": list(self.recent), "caption": "",
                               "_row": None, "_cell": None})
        elif tag == "caption" and self.stack:
            self._caption_buf = []
        elif tag == "tr" and self.stack:
            self.stack[-1]["_row"] = []
        elif tag in ("td", "th") and self.stack:
            t = self.stack[-1]
            if t["_row"] is None:
                t["_row"] = []
            try:
                span = max(1, int(a.get("colspan", "1") or 1))
            except ValueError:
                span = 1
            t["_cell"] = {"text": [], "th": tag == "th", "span": span}
        elif tag == "br":
            self._add_text("\n")

    def _add_text(self, data):
        if self._skip:
            return
        if self.stack:
            t = self.stack[-1]
            if self._caption_buf is not None:
                self._caption_buf.append(data)
            elif t["_cell"] is not None:
                t["_cell"]["text"].append(data)
        elif self._heading_buf is not None:
            self._heading_buf.append(data)
        else:
            s = " ".join(data.split())
            if s:
                self.recent.append(s)

    def handle_data(self, data):
        self._add_text(data)

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in HEADING_TAGS and self._heading_buf is not None:
            txt = " ".join("".join(self._heading_buf).split())
            if txt:
                self.last_heading = txt
                self.recent.clear()
            self._heading_buf = None
        elif tag == "caption" and self.stack and self._caption_buf is not None:
            self.stack[-1]["caption"] = " ".join("".join(self._caption_buf).split())
            self._caption_buf = None
        elif tag in ("td", "th") and self.stack:
            self._close_cell(self.stack[-1])
        elif tag == "tr" and self.stack:
            self._close_row(self.stack[-1])
        elif tag == "table" and self.stack:
            t = self.stack.pop()
            self._close_row(t)
            if t["rows"]:
                del t["_row"], t["_cell"]
                self.tables.append(t)

    @staticmethod
    def _close_cell(t):
        c = t["_cell"]
        if c is None:
            return
        txt = "\n".join(" ".join(l.split()) for l in "".join(c["text"]).split("\n"))
        txt = txt.strip("\n ")
        t["_row"].append({"text": txt, "th": c["th"], "span": c["span"]})
        t["_cell"] = None

    def _close_row(self, t):
        self._close_cell(t)
        if t["_row"]:
            t["rows"].append(t["_row"])
        t["_row"] = None


# --------------------------------------------------------------------------
# 2. Reconnaissance des colonnes
# --------------------------------------------------------------------------

FIELD_ALIASES = OrderedDict([
    ("name", ["name", "interface", "interface name", "nom", "if name", "ifname",
              "iface", "nom interface", "nom de l'interface", "object name"]),
    ("ip", ["ip address", "ipv4 address", "ip", "ipv4", "address", "adresse ip",
            "adresse", "ip addr", "ipv4 addr", "ip/mask", "ipv4 address/mask"]),
    ("mask", ["net mask", "netmask", "mask", "subnet mask", "ipv4 net mask",
              "ipv4 mask", "network mask", "mask length", "ipv4 mask length",
              "prefix", "prefix length", "masque", "masque de sous-réseau",
              "ipv4 network mask"]),
    ("topology", ["topology", "leads to", "topologie", "network type",
                  "type de réseau", "network topology", "interface topology",
                  "topology type"]),
    ("member", ["member", "cluster member", "membre", "gateway", "device",
                "member name", "nom du membre", "object", "host"]),
    ("dest", ["destination", "destination network", "dest", "réseau destination",
              "route", "network destination"]),
    ("nexthop", ["next hop", "nexthop", "next-hop", "via", "passerelle",
                 "gateway", "next hop gateway", "prochain saut"]),
    ("antispoof", ["anti spoofing", "anti-spoofing", "antispoofing"]),
    ("vlan", ["vlan", "vlan id", "vlan tag"]),
])


def norm(s):
    s = (s or "").lower().replace("_", " ").replace("\n", " ")
    s = re.sub(r"[:*()\[\]]", " ", s)
    return " ".join(s.split())


def match_field(header):
    h = norm(header)
    if not h:
        return None
    for field, aliases in FIELD_ALIASES.items():
        if h in aliases:
            return field
    for field, aliases in FIELD_ALIASES.items():
        for al in aliases:
            if len(al) >= 6 and (h.startswith(al) or al in h):
                return field
    return None


RULE_ALIASES = OrderedDict([
    ("no", ["no", "no.", "#", "rule", "rule number", "rule no", "n°", "numéro", "num"]),
    ("rname", ["name", "rule name", "nom"]),
    ("src", ["source", "src", "sources"]),
    ("dst", ["destination", "dst", "destinations"]),
    ("vpn", ["vpn"]),
    ("svc", ["services & applications", "services and applications", "services",
             "service", "services & apps", "applications"]),
    ("content", ["content", "data", "contenu"]),
    ("action", ["action", "actions"]),
    ("track", ["track", "log", "logs", "journalisation", "suivi"]),
    ("install", ["install on", "installé sur", "install", "target", "targets"]),
    ("time", ["time", "heure"]),
    ("hits", ["hits", "hit count"]),
    ("comment", ["comments", "comment", "commentaires", "commentaire"]),
])


def analyse_rule_header(rows):
    for i, r in enumerate(rows[:3]):
        cells = expand_row(r)
        mapping = {}
        for j, cell in enumerate(cells):
            h = norm(cell)
            for f, al in RULE_ALIASES.items():
                if (h in al or (h and any(len(a) >= 6 and h.startswith(a) for a in al))) \
                        and f not in mapping.values():
                    mapping[j] = f
                    break
        if {"src", "dst", "action"} <= set(mapping.values()):
            return mapping, i + 1
    return {}, 0


def layer_from_text(t):
    """'Network Layer', 'Layer: Appli', 'Standard - Network layer' -> nom de couche."""
    if not t or not re.search(r"\b(layer|layers|couche)\b", t, re.I):
        return None
    s = re.split(r"\s+[-–|]\s+", t)[-1]
    s = re.sub(r"\b(access control|policy|politique|ordered|inline)\b", " ", s, flags=re.I)
    s = re.sub(r"\b(layers?|couches?)\b\s*[:\-]?", " ", s, flags=re.I)
    s = re.sub(r"^\s*(de la|du|de|of|:)\s+", "", " ".join(s.split()), flags=re.I)
    return s.strip(" :-\"'") or None


def expand_row(row):
    out = []
    for c in row:
        out.append(c["text"])
        out.extend([""] * (c["span"] - 1))
    return out


def split_subtables(t):
    """Découpe un tableau sur les lignes-titres (une seule cellule non vide)."""
    width = max(len(expand_row(r)) for r in t["rows"])
    parts, cur, sub = [], [], ""
    for r in t["rows"]:
        cells = expand_row(r)
        non_empty = [c for c in cells if c.strip()]
        if width >= 2 and len(non_empty) == 1 and (len(r) == 1 or r[0]["span"] > 1):
            if cur:
                parts.append((sub, cur))
            sub, cur = non_empty[0], []
        else:
            cur.append(r)
    if cur:
        parts.append((sub, cur))
    return parts


def analyse_header(rows):
    """Retourne (mapping {index: champ}, index_première_ligne_de_données)."""
    for i, r in enumerate(rows[:3]):
        cells = expand_row(r)
        mapping = {}
        for j, c in enumerate(cells):
            f = match_field(c)
            if f and f not in mapping.values():
                mapping[j] = f
        all_th = all(c["th"] for c in r)
        if len(mapping) >= 2 or (all_th and mapping):
            fields = set(mapping.values())
            # 'gateway' = prochain saut si le tableau contient une destination
            if "dest" in fields:
                for j, f in list(mapping.items()):
                    if f == "member" and norm(cells[j]) in ("gateway", "passerelle"):
                        mapping[j] = "nexthop"
            return mapping, i + 1
    return {}, 0


# --------------------------------------------------------------------------
# 3. Parsing des valeurs réseau
# --------------------------------------------------------------------------

IP_RE = re.compile(r"\b(\d{1,3}(?:\.\d{1,3}){3})(?:\s*/\s*(\d{1,2}))?\b")


def parse_prefix(mask_txt):
    if not mask_txt:
        return None
    m = IP_RE.search(mask_txt)
    if m:
        try:
            return ipaddress.IPv4Network("0.0.0.0/" + m.group(1)).prefixlen
        except ValueError:
            pass
    m = re.search(r"/?\b(\d{1,2})\b", mask_txt)
    if m and 0 <= int(m.group(1)) <= 32:
        return int(m.group(1))
    return None


def parse_ip(ip_txt, mask_txt=""):
    """'10.0.0.1/24' ou ('10.0.0.1', '255.255.255.0') -> (ip, prefix)"""
    m = IP_RE.search(ip_txt or "")
    if not m:
        return None, None
    ip = m.group(1)
    try:
        ipaddress.IPv4Address(ip)
    except ValueError:
        return None, None
    prefix = int(m.group(2)) if m.group(2) else parse_prefix(mask_txt)
    return ip, prefix


def classify_topology(txt, ifname=""):
    t = norm(txt) + " " + norm(ifname)
    if "external" in t or "externe" in t or "internet" in t:
        return "external"
    if "dmz" in t:
        return "dmz"
    if "sync" in t:
        return "sync"
    if "mgmt" in t or "management" in t or "gestion" in t:
        return "mgmt"
    if any(k in t for k in ("internal", "interne", "this network", "specific",
                            "leads to", "not defined")):
        return "internal"
    return "unknown"


# --------------------------------------------------------------------------
# 4. Reconstitution des objets
# --------------------------------------------------------------------------

OBJ_HEAD_RE = re.compile(
    r"(cluster\s*member|cluster|membre|member|gateway|passerelle|gw|firewall)"
    r"(?:\s+object)?\s*[:=]\s*[\"']?([A-Za-z0-9_.\-]+)", re.I)
OBJ_PAREN_RE = re.compile(r"^([A-Za-z0-9_.\-]+)\s*\(([^)]*)\)")

KV_NAME_KEYS = ("name", "object name", "nom", "nom de l'objet")
KV_TYPE_KEYS = ("type", "object type", "class", "classe", "object class")
KV_MEMBERS_KEYS = ("cluster members", "members", "membres", "membres du cluster")
KV_IP_KEYS = ("ipv4 address", "ip address", "main ip", "adresse ip", "ip")


def kind_from_text(t):
    t = norm(t)
    if "member" in t or "membre" in t:
        return "member"
    if "cluster" in t:
        return "cluster"
    if any(k in t for k in ("gateway", "passerelle", "gw", "firewall", "simple")):
        return "gateway"
    return None


def parse_obj_heading(text):
    if not text:
        return None, None
    m = OBJ_HEAD_RE.search(text)
    if m:
        return kind_from_text(m.group(1)), m.group(2)
    m = OBJ_PAREN_RE.match(text.strip())
    if m and kind_from_text(m.group(2)):
        return kind_from_text(m.group(2)), m.group(1)
    return None, None


class Model:
    def __init__(self):
        self.objects = OrderedDict()   # nom -> objet
        self.routes = []               # {owner, dest, prefix, nexthop}
        self.rules = []                # règles de politique (toutes couches)
        self.layers = OrderedDict()    # nom de couche -> {rules, sources}
        self.tables_log = []

    def obj(self, name):
        if name not in self.objects:
            self.objects[name] = {"name": name, "kind": None, "type": "",
                                  "ip": None, "cluster": None, "members": [],
                                  "interfaces": [], "props": OrderedDict()}
        return self.objects[name]

    def add_iface(self, owner, iface):
        o = self.obj(owner)
        for ex in o["interfaces"]:
            if ex["name"] == iface["name"] and ex["ip"] == iface["ip"]:
                for k, v in iface.items():
                    if v and not ex.get(k):
                        ex[k] = v
                return
        o["interfaces"].append(iface)


def find_known_name(text, model):
    best = None
    for n in model.objects:
        if n and re.search(r"(?<![\w.\-])" + re.escape(n) + r"(?![\w.\-])", text or ""):
            if best is None or len(n) > len(best):
                best = n
    return best


def build_model(html_texts, debug=False):
    model = Model()
    current_obj = None
    current_cluster = None
    current_layer = None

    for src, text in html_texts:
        p = TableExtractor()
        p.feed(text)
        p.close()
        for t in p.tables:
            for sub, rows in split_subtables(t):
                ctx_list = [t["heading"], t["caption"], sub] + t["context"]
                ctx = " | ".join(c for c in ctx_list if c)

                # Objet nommé par un titre ("Cluster: CL-FW", "fw01 (Cluster Member)")
                for c in (sub, t["caption"], t["heading"]):
                    kind, name = parse_obj_heading(c)
                    if name:
                        o = model.obj(name)
                        o["kind"] = o["kind"] or kind
                        if kind == "member" and current_cluster and not o["cluster"]:
                            o["cluster"] = current_cluster
                        current_obj = name
                        if kind == "cluster":
                            current_cluster = name
                        break

                width = max(len(expand_row(r)) for r in rows)

                # --- Couche (layer) courante ----------------------------------
                for cc in (sub, t["caption"], t["heading"]):
                    ln = layer_from_text(cc)
                    if ln:
                        current_layer = ln
                        break

                # --- Règles de politique (une ou plusieurs couches) ----------
                rmap, rstart = analyse_rule_header(rows)
                if not rmap and t.get("_rmap") and width == t["_rw"]:
                    rmap, rstart = t["_rmap"], 0      # suite après une ligne-section
                if rmap:
                    t["_rmap"], t["_rw"] = rmap, width
                    layer = current_layer or t["heading"] or "Couche 1"
                    section = sub if sub and not layer_from_text(sub) else ""
                    lay = model.layers.setdefault(layer, {"name": layer, "rules": 0,
                                                          "sources": set()})
                    lay["sources"].add(os.path.basename(src))
                    for r in rows[rstart:]:
                        cells = expand_row(r)
                        rec = {f: (cells[j] if j < len(cells) else "").strip()
                               for j, f in rmap.items()}
                        if not any(rec.get(k) for k in ("src", "dst", "action")):
                            continue
                        rec["layer"], rec["section"] = layer, section
                        rec["disabled"] = bool(re.search(r"disabled|désactiv", " ".join(cells), re.I))
                        model.rules.append(rec)
                        lay["rules"] += 1
                    model.tables_log.append({
                        "source": os.path.basename(src), "contexte": ctx[:120],
                        "colonnes": {str(k): v for k, v in rmap.items()},
                        "lignes": len(rows), "interprétation": "règles [%s]" % layer,
                        "entête": expand_row(rows[rstart - 1]) if rstart else []})
                    continue

                mapping, start = analyse_header(rows)
                fields = set(mapping.values())
                data = [expand_row(r) for r in rows[start:]]
                kind_tab = "?"

                # --- Tableau clé / valeur -----------------------------------
                if width == 2 and not ({"ip", "mask"} <= fields or "dest" in fields):
                    kind_tab = "clé/valeur"
                    kv = OrderedDict()
                    for r in [expand_row(r) for r in rows]:
                        if len(r) >= 2 and r[0].strip():
                            kv[norm(r[0])] = r[1].strip()
                    name = next((kv[k] for k in KV_NAME_KEYS if kv.get(k)), None)
                    if name:
                        name = name.split()[0]
                    elif current_obj:
                        name = current_obj
                    if name:
                        o = model.obj(name)
                        o["props"].update(kv)
                        typ = next((kv[k] for k in KV_TYPE_KEYS if kv.get(k)), "")
                        if typ:
                            o["type"] = typ
                            o["kind"] = kind_from_text(typ) or o["kind"]
                        ipv = next((kv[k] for k in KV_IP_KEYS if kv.get(k)), "")
                        if ipv and not o["ip"]:
                            o["ip"] = parse_ip(ipv)[0]
                        mem = next((kv[k] for k in KV_MEMBERS_KEYS if kv.get(k)), "")
                        if mem:
                            o["kind"] = o["kind"] or "cluster"
                            for mname in re.split(r"[,;\n]+|\s{2,}", mem):
                                mname = mname.strip().split(" ")[0]
                                if mname and mname not in o["members"]:
                                    o["members"].append(mname)
                        current_obj = name
                        if o["kind"] == "cluster":
                            current_cluster = name
                        elif o["kind"] == "member" and current_cluster and not o["cluster"]:
                            o["cluster"] = current_cluster

                # --- Routes statiques ----------------------------------------
                elif "dest" in fields and ("nexthop" in fields or "name" in fields):
                    kind_tab = "routes"
                    owner = find_known_name(ctx, model) or current_obj or "?"
                    for r in data:
                        rec = {f: (r[j] if j < len(r) else "") for j, f in mapping.items()}
                        dip, dpfx = parse_ip(rec.get("dest", ""), rec.get("mask", ""))
                        if "default" in norm(rec.get("dest", "")):
                            dip, dpfx = "0.0.0.0", 0
                        nh = parse_ip(rec.get("nexthop", ""))[0]
                        if dip is not None and nh:
                            model.routes.append({"owner": owner, "dest": dip,
                                                 "prefix": dpfx if dpfx is not None else 32,
                                                 "nexthop": nh})

                # --- Liste des membres (nom [+ IP], sans masque) ---------------
                elif ("name" in fields or "member" in fields) and "mask" not in fields and \
                        not any("/" in (r[k] if k < len(r) else "") for r in data
                                for k, f in mapping.items() if f == "ip") and \
                        re.search(r"member|membre", " ".join([sub, t["caption"], t["heading"]]), re.I):
                    kind_tab = "membres"
                    cl = current_cluster or current_obj
                    col = next((j for j, f in mapping.items() if f in ("member", "name")), 0)
                    ipcol = next((j for j, f in mapping.items() if f == "ip"), None)
                    for r in data:
                        mname = (r[col] if col < len(r) else "").strip().split(" ")[0]
                        if not mname:
                            continue
                        mo = model.obj(mname)
                        mo["kind"] = "member"
                        if cl:
                            mo["cluster"] = cl
                            model.obj(cl)["kind"] = model.obj(cl)["kind"] or "cluster"
                            if mname not in model.obj(cl)["members"]:
                                model.obj(cl)["members"].append(mname)
                        if ipcol is not None and ipcol < len(r):
                            mo["ip"] = mo["ip"] or parse_ip(r[ipcol])[0]

                # --- Tableau d'interfaces ------------------------------------
                elif "ip" in fields and ("name" in fields or "mask" in fields):
                    kind_tab = "interfaces"
                    ctx_owner = None
                    for c in (sub, t["caption"], t["heading"]):
                        ctx_owner = find_known_name(c, model)
                        if ctx_owner:
                            break
                    owner_default = ctx_owner or current_obj
                    for r in data:
                        rec = {f: (r[j] if j < len(r) else "") for j, f in mapping.items()}
                        ip, pfx = parse_ip(rec.get("ip", ""), rec.get("mask", ""))
                        if not ip:
                            continue
                        owner = (rec.get("member") or "").strip().split(" ")[0] or owner_default
                        if not owner:
                            owner = "FIREWALL"
                        if rec.get("member") and current_cluster and owner != current_cluster:
                            mo = model.obj(owner)
                            mo["kind"] = mo["kind"] or "member"
                            mo["cluster"] = mo["cluster"] or current_cluster
                        model.add_iface(owner, {
                            "name": (rec.get("name") or "").split("\n")[0] or "?",
                            "ip": ip, "prefix": pfx,
                            "topology_raw": rec.get("topology", ""),
                            "topology": classify_topology(rec.get("topology", ""),
                                                          rec.get("name", "")),
                            "vlan": rec.get("vlan", ""),
                        })

                model.tables_log.append({
                    "source": os.path.basename(src), "contexte": ctx[:120],
                    "colonnes": {str(k): v for k, v in mapping.items()},
                    "lignes": len(rows), "interprétation": kind_tab,
                    "entête": expand_row(rows[start - 1]) if start else expand_row(rows[0]),
                })
    return model


# --------------------------------------------------------------------------
# 5. Déduction de la topologie
# --------------------------------------------------------------------------

def net_of(iface):
    if iface.get("ip") and iface.get("prefix") is not None:
        return ipaddress.IPv4Interface("%s/%s" % (iface["ip"], iface["prefix"])).network
    return None


def build_topology(model, forced_cluster=None):
    objs = model.objects
    clusters = []
    if forced_cluster:
        model.obj(forced_cluster)["kind"] = "cluster"
    for o in objs.values():
        if o["kind"] == "cluster" or ("cluster" in norm(o["type"]) and "member" not in norm(o["type"])):
            clusters.append(o)
    if not clusters:
        # Pas de cluster explicite : on prend les objets qui ont des interfaces
        cands = [o for o in objs.values() if o["interfaces"] and o["kind"] != "member"]
        clusters = cands[:1] or [o for o in objs.values() if o["interfaces"]][:1]

    topologies = []
    for cl in clusters:
        members = []
        for n in cl["members"]:
            members.append(model.obj(n))
        for o in objs.values():
            if o["kind"] == "member" and o["name"] not in [m["name"] for m in members]:
                if o["cluster"] in (cl["name"], None) or len(clusters) == 1:
                    members.append(o)
        vip_nets = {net_of(i) for i in cl["interfaces"] if net_of(i)}
        if not members:   # passerelles partageant les réseaux des VIP
            for o in objs.values():
                if o is cl or not o["interfaces"] or o["kind"] == "cluster":
                    continue
                if any(net_of(i) in vip_nets for i in o["interfaces"]):
                    members.append(o)

        networks = OrderedDict()   # str(net) -> info

        def get_net(net, cls):
            k = str(net)
            if k not in networks:
                networks[k] = {"network": k, "class": cls, "vip": [], "members": [],
                               "routers": OrderedDict()}
            elif networks[k]["class"] in ("unknown",) and cls != "unknown":
                networks[k]["class"] = cls
            return networks[k]

        for i in cl["interfaces"]:
            n = net_of(i)
            if n is None:
                continue
            get_net(n, i["topology"])["vip"].append({"if": i["name"], "ip": i["ip"]})

        for m in members:
            for i in m["interfaces"]:
                n = net_of(i)
                if n is None:
                    continue
                cls = i["topology"]
                if str(n) not in networks and cls in ("unknown", "internal"):
                    cls = "mgmt" if "mgmt" in i["name"].lower() else "sync"
                get_net(n, cls)["members"].append({"member": m["name"], "if": i["name"],
                                                   "ip": i["ip"]})

        # Routes statiques -> routeurs de prochain saut
        owners = {cl["name"]} | {m["name"] for m in members} | {"?"}
        for r in model.routes:
            if r["owner"] not in owners:
                continue
            nh = ipaddress.IPv4Address(r["nexthop"])
            for info in networks.values():
                if nh in ipaddress.IPv4Network(info["network"]):
                    dest = "défaut" if r["prefix"] == 0 else "%s/%s" % (r["dest"], r["prefix"])
                    lst = info["routers"].setdefault(r["nexthop"], [])
                    if dest not in lst:
                        lst.append(dest)

        topologies.append({"cluster": cl, "members": members,
                           "networks": list(networks.values()),
                           "checks": run_checks(cl, members, networks)})
    return topologies


def run_checks(cl, members, networks):
    out = []
    seen = {}
    for o in [cl] + members:
        for i in o["interfaces"]:
            if i["prefix"] is None:
                out.append("Masque inconnu : %s %s (%s)" % (o["name"], i["name"], i["ip"]))
            seen.setdefault(i["ip"], []).append("%s/%s" % (o["name"], i["name"]))
    for ip, who in seen.items():
        if len(who) > 1:
            out.append("Adresse IP en double %s : %s" % (ip, ", ".join(who)))
    for n in networks.values():
        if n["vip"]:
            have = {x["member"] for x in n["members"]}
            for m in members:
                if m["name"] not in have:
                    out.append("Le membre %s n'a pas d'interface dans le réseau %s (VIP %s)"
                               % (m["name"], n["network"], n["vip"][0]["ip"]))
            for x in n["members"]:
                if x["if"] != n["vip"][0]["if"] and n["vip"][0]["if"] != "?":
                    out.append("Nom d'interface différent sur %s : VIP %s / %s %s"
                               % (n["network"], n["vip"][0]["if"], x["member"], x["if"]))
        if n["class"] == "unknown" and n["vip"]:
            out.append("Topologie non renseignée pour %s" % n["network"])
    if len(members) < 2:
        out.append("Moins de 2 membres détectés pour le cluster %s" % cl["name"])
    return out


# --------------------------------------------------------------------------
# 6. Écriture Visio (.vsdx) - Open Packaging Conventions, sans dépendance
# --------------------------------------------------------------------------

NS = "http://schemas.microsoft.com/office/visio/2012/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_VISIO = "http://schemas.microsoft.com/visio/2010/relationships/"

STYLE = {
    "external": ("#F8CECC", "#B85450"),
    "internal": ("#D5E8D4", "#82B366"),
    "dmz":      ("#FFE6CC", "#D79B00"),
    "sync":     ("#E6E6E6", "#666666"),
    "mgmt":     ("#E1D5E7", "#9673A6"),
    "unknown":  ("#F5F5F5", "#999999"),
    "cluster":  ("#DAE8FC", "#6C8EBF"),
    "member":   ("#FFFFFF", "#1F4E79"),
    "router":   ("#FFF2CC", "#B8860B"),
}
LABELS = {"external": "Externe", "internal": "Interne", "dmz": "DMZ",
          "sync": "Synchro", "mgmt": "Gestion", "unknown": "Topologie ?"}


def x(s):
    return html.escape(str(s), quote=True)


def c(n, v, f=None, u=None):
    extra = (' U="%s"' % u if u else "") + (' F="%s"' % x(f) if f else "")
    return '<Cell N="%s" V="%s"%s/>' % (n, x(v), extra)


class VisioPage:
    def __init__(self, name):
        self.name = name
        self.shapes_back, self.shapes_front, self.connects = [], [], []
        self.next_id = 1
        self.width, self.height = 11.0, 8.5

    def _id(self):
        i = self.next_id
        self.next_id += 1
        return i

    def box(self, cx, cy, w, h, text, fill, line, size_pt=9, bold=False,
            rounding=0.08, ellipse=False, valign=1, lw_pt=1.0, back=False, name=None):
        sid = self._id()
        cells = [c("PinX", cx), c("PinY", cy), c("Width", w), c("Height", h),
                 c("LocPinX", w / 2, "Width*0.5"), c("LocPinY", h / 2, "Height*0.5"),
                 c("Angle", 0), c("FillForegnd", fill), c("FillPattern", 1),
                 c("LineColor", line), c("LineWeight", lw_pt / 72.0), c("LinePattern", 1),
                 c("Rounding", 0 if ellipse else rounding), c("VerticalAlign", valign),
                 c("TopMargin", 0.05), c("BottomMargin", 0.05),
                 c("LeftMargin", 0.06), c("RightMargin", 0.06)]
        if ellipse:
            geom = ('<Section N="Geometry" IX="0">%s%s%s'
                    '<Row T="Ellipse" IX="1">%s%s%s%s%s%s</Row></Section>') % (
                c("NoFill", 0), c("NoLine", 0), c("NoShow", 0),
                c("X", w / 2, "Width*0.5"), c("Y", h / 2, "Height*0.5"),
                c("A", w, "Width*1"), c("B", h / 2, "Height*0.5"),
                c("C", w / 2, "Width*0.5"), c("D", h, "Height*1"))
        else:
            pts = [(0, 0), (1, 0), (1, 1), (0, 1), (0, 0)]
            rows = ['<Row T="RelMoveTo" IX="1">%s%s</Row>' % (c("X", 0), c("Y", 0))]
            for k, (px, py) in enumerate(pts[1:], start=2):
                rows.append('<Row T="RelLineTo" IX="%d">%s%s</Row>' % (k, c("X", px), c("Y", py)))
            geom = '<Section N="Geometry" IX="0">%s%s%s%s</Section>' % (
                c("NoFill", 0), c("NoLine", 0), c("NoShow", 0), "".join(rows))
        char = '<Section N="Character"><Row IX="0">%s%s%s</Row></Section>' % (
            c("Size", size_pt / 72.0, u="PT"), c("Style", 1 if bold else 0), c("Color", "#000000"))
        para = '<Section N="Paragraph"><Row IX="0">%s</Row></Section>' % c("HorzAlign", 1)
        xml = ('<Shape ID="%d" NameU="%s" Name="%s" Type="Shape" LineStyle="0" FillStyle="0" '
               'TextStyle="0">%s%s%s%s<Text>%s</Text></Shape>') % (
            sid, x(name or "Shape.%d" % sid), x(name or "Shape.%d" % sid),
            "".join(cells), char, para, geom, x(text))
        (self.shapes_back if back else self.shapes_front).append(xml)
        return sid, (cx, cy, w, h)

    def connector(self, a, b, text="", color="#404040", lw_pt=1.5, dash=False, arrow=False):
        """Connecteur collé (glue dynamique) entre deux formes a=(id,geo) b=(id,geo)."""
        (ida, (ax, ay, aw, ah)), (idb, (bx, by, bw, bh)) = a, b

        def edge(cx, cy, w, h, tx, ty):   # point de sortie sur le bord de la forme
            dx, dy = tx - cx, ty - cy
            if dx == 0 and dy == 0:
                return cx, cy
            sx = (w / 2) / abs(dx) if dx else float("inf")
            sy = (h / 2) / abs(dy) if dy else float("inf")
            s = min(sx, sy)
            return cx + dx * s, cy + dy * s

        bx0, by0 = edge(ax, ay, aw, ah, bx, by)
        ex0, ey0 = edge(bx, by, bw, bh, ax, ay)
        sid = self._id()
        w, h = ex0 - bx0, ey0 - by0
        cells = [c("PinX", (bx0 + ex0) / 2, "GUARD((BeginX+EndX)/2)"),
                 c("PinY", (by0 + ey0) / 2, "GUARD((BeginY+EndY)/2)"),
                 c("Width", w, "GUARD(EndX-BeginX)"), c("Height", h, "GUARD(EndY-BeginY)"),
                 c("LocPinX", w / 2, "GUARD(Width*0.5)"), c("LocPinY", h / 2, "GUARD(Height*0.5)"),
                 c("BeginX", bx0, "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)"),
                 c("BeginY", by0, "_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)"),
                 c("EndX", ex0, "_WALKGLUE(EndTrigger,BegTrigger,WalkPreference)"),
                 c("EndY", ey0, "_WALKGLUE(EndTrigger,BegTrigger,WalkPreference)"),
                 c("BegTrigger", 2, "_XFTRIGGER(Sheet.%d!EventXFMod)" % ida),
                 c("EndTrigger", 2, "_XFTRIGGER(Sheet.%d!EventXFMod)" % idb),
                 c("ObjType", 2), c("ShapeRouteStyle", 16), c("ConFixedCode", 6),
                 c("LineColor", color), c("LineWeight", lw_pt / 72.0),
                 c("LinePattern", 2 if dash else 1), c("BeginArrow", 0), c("EndArrow", 4 if arrow else 0),
                 c("EndArrowSize", 2),
                 c("FillPattern", 0), c("TextBkgnd", "#FFFFFF"),
                 c("TxtPinX", w / 2, "SETATREF(Controls.TextPosition)"),
                 c("TxtPinY", h / 2, "SETATREF(Controls.TextPosition.Y)"),
                 c("TxtWidth", 1.4), c("TxtHeight", 0.4),
                 c("TxtLocPinX", 0.7, "TxtWidth*0.5"), c("TxtLocPinY", 0.2, "TxtHeight*0.5")]
        ctrl = ('<Section N="Control"><Row N="TextPosition">%s%s%s%s%s%s</Row></Section>') % (
            c("X", w / 2), c("Y", h / 2), c("XDyn", w / 2, "Controls.TextPosition"),
            c("YDyn", h / 2, "Controls.TextPosition.Y"), c("XCon", 0), c("YCon", 0))
        geom = ('<Section N="Geometry" IX="0">%s%s%s'
                '<Row T="MoveTo" IX="1">%s%s</Row><Row T="LineTo" IX="2">%s%s</Row></Section>') % (
            c("NoFill", 1), c("NoLine", 0), c("NoShow", 0),
            c("X", 0), c("Y", 0), c("X", w, "Width"), c("Y", h, "Height"))
        char = '<Section N="Character"><Row IX="0">%s%s</Row></Section>' % (
            c("Size", 7.5 / 72.0, u="PT"), c("Color", "#202020"))
        xml = ('<Shape ID="%d" NameU="Dynamic connector.%d" Name="Connecteur.%d" Type="Shape" '
               'LineStyle="0" FillStyle="0" TextStyle="0">%s%s%s%s<Text>%s</Text></Shape>') % (
            sid, sid, sid, "".join(cells), ctrl, char, geom, x(text))
        self.shapes_back.append(xml)   # sous les formes
        self.connects.append('<Connect FromSheet="%d" FromCell="BeginX" FromPart="9" '
                             'ToSheet="%d" ToCell="PinX" ToPart="3"/>' % (sid, ida))
        self.connects.append('<Connect FromSheet="%d" FromCell="EndX" FromPart="12" '
                             'ToSheet="%d" ToCell="PinX" ToPart="3"/>' % (sid, idb))
        return sid

    def xml(self):
        return ('<?xml version="1.0" encoding="utf-8"?>'
                '<PageContents xmlns="%s" xmlns:r="%s" xml:space="preserve"><Shapes>%s</Shapes>'
                '<Connects>%s</Connects></PageContents>') % (
            NS, NS_R, "".join(self.shapes_back + self.shapes_front), "".join(self.connects))


def write_vsdx(pages, path, title="Topologie Check Point"):
    ct = ('<?xml version="1.0" encoding="utf-8"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>'
          '<Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>'
          + "".join('<Override PartName="/visio/pages/page%d.xml" '
                    'ContentType="application/vnd.ms-visio.page+xml"/>' % (i + 1)
                    for i in range(len(pages))) +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="utf-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="%sdocument" Target="visio/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>') % REL_VISIO
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="utf-8"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>cp_topo.py</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '</cp:coreProperties>') % (x(title), now)
    app = ('<?xml version="1.0" encoding="utf-8"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>Microsoft Visio</Application><AppVersion>16.0000</AppVersion></Properties>')
    style0 = ("".join([c("EnableLineProps", 1), c("EnableFillProps", 1), c("EnableTextProps", 1),
                       c("HideForApply", 0), c("LineWeight", 0.01041666666666667),
                       c("LineColor", 0), c("LinePattern", 1), c("Rounding", 0),
                       c("EndArrowSize", 2), c("BeginArrow", 0), c("EndArrow", 0),
                       c("LineCap", 0), c("BeginArrowSize", 2), c("LineColorTrans", 0),
                       c("FillForegnd", 1), c("FillBkgnd", 0), c("FillPattern", 1),
                       c("ShdwForegnd", 0), c("ShdwPattern", 0), c("FillForegndTrans", 0),
                       c("FillBkgndTrans", 0), c("ShdwForegndTrans", 0),
                       c("LeftMargin", 0), c("RightMargin", 0), c("TopMargin", 0),
                       c("BottomMargin", 0), c("VerticalAlign", 1), c("TextBkgnd", 0),
                       c("TextDirection", 0), c("TextBkgndTrans", 0)])
              + '<Section N="Character"><Row IX="0">%s%s%s%s</Row></Section>' % (
                  c("Font", "Calibri"), c("Color", 0), c("Style", 0), c("Size", 0.1666666666666667))
              + '<Section N="Paragraph"><Row IX="0">%s</Row></Section>' % c("HorzAlign", 1))
    doc = ('<?xml version="1.0" encoding="utf-8"?>'
           '<VisioDocument xmlns="%s" xmlns:r="%s" xml:space="preserve">'
           '<DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" '
           'DefaultFillStyle="0" DefaultGuideStyle="0"><GlueSettings>9</GlueSettings>'
           '<SnapSettings>65847</SnapSettings><SnapExtensions>34</SnapExtensions>'
           '<DynamicGridEnabled>0</DynamicGridEnabled><ProtectStyles>0</ProtectStyles>'
           '<ProtectShapes>0</ProtectShapes><ProtectMasters>0</ProtectMasters>'
           '<ProtectBkgnds>0</ProtectBkgnds></DocumentSettings>'
           '<Colors><ColorEntry IX="0" RGB="#000000"/><ColorEntry IX="1" RGB="#FFFFFF"/></Colors>'
           '<FaceNames><FaceName NameU="Calibri" UnicodeRanges="-536859905 -1073732485 9 0" '
           'CharSets="536871423 0" Panose="2 15 5 2 2 2 4 3 2 4" Flags="325"/></FaceNames>'
           '<StyleSheets><StyleSheet ID="0" NameU="No Style" Name="No Style">%s</StyleSheet>'
           '</StyleSheets></VisioDocument>') % (NS, NS_R, style0)
    doc_rels = ('<?xml version="1.0" encoding="utf-8"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="%spages" Target="pages/pages.xml"/></Relationships>'
                ) % REL_VISIO
    pages_xml = ['<?xml version="1.0" encoding="utf-8"?><Pages xmlns="%s" xmlns:r="%s" '
                 'xml:space="preserve">' % (NS, NS_R)]
    pages_rels = ['<?xml version="1.0" encoding="utf-8"?>'
                  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, p in enumerate(pages):
        pages_xml.append(
            '<Page ID="%d" NameU="%s" Name="%s" ViewScale="-1" ViewCenterX="%s" ViewCenterY="%s">'
            '<PageSheet LineStyle="0" FillStyle="0" TextStyle="0">%s%s%s%s%s%s%s</PageSheet>'
            '<Rel r:id="rId%d"/></Page>' % (
                i, x(p.name), x(p.name), p.width / 2, p.height / 2,
                c("PageWidth", p.width, u="IN"), c("PageHeight", p.height, u="IN"),
                c("PageScale", 1, u="IN"), c("DrawingScale", 1, u="IN"),
                c("DrawingSizeType", 0), c("DrawingScaleType", 0),
                c("PrintPageOrientation", 2), i + 1))
        pages_rels.append('<Relationship Id="rId%d" Type="%spage" Target="page%d.xml"/>'
                          % (i + 1, REL_VISIO, i + 1))
    pages_xml.append("</Pages>")
    pages_rels.append("</Relationships>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
        z.writestr("visio/document.xml", doc)
        z.writestr("visio/_rels/document.xml.rels", doc_rels)
        z.writestr("visio/pages/pages.xml", "".join(pages_xml))
        z.writestr("visio/pages/_rels/pages.xml.rels", "".join(pages_rels))
        for i, p in enumerate(pages):
            z.writestr("visio/pages/page%d.xml" % (i + 1), p.xml())


# --------------------------------------------------------------------------
# 7. Mise en page de la topologie
# --------------------------------------------------------------------------

def short(ip):
    return "." + ip.split(".")[-1]


def draw_topology(topo):
    cl, members, nets = topo["cluster"], topo["members"], topo["networks"]
    page = VisioPage(cl["name"][:30])
    NET_W, NET_H, GAP = 2.3, 0.85, 0.55
    MEM_W = 2.3
    margin = 0.8

    top = [n for n in nets if n["class"] == "external"]
    side = [n for n in nets if n["class"] in ("sync", "mgmt") and not n["vip"]]
    bottom = [n for n in nets if n not in top and n not in side]

    def mem_lines(m):
        ls = [m["name"]]
        if m.get("ip"):
            ls.append("Principale : %s" % m["ip"])
        for i in m["interfaces"][:12]:
            ls.append("%s  %s/%s" % (i["name"], i["ip"], i["prefix"] if i["prefix"] is not None else "?"))
        if len(m["interfaces"]) > 12:
            ls.append("... +%d" % (len(m["interfaces"]) - 12))
        return ls

    mem_h = max([0.4 + 0.15 * len(mem_lines(m)) for m in members] + [1.0])
    n_mem = max(len(members), 1)
    cl_w = max(n_mem * (MEM_W + 0.4) + 0.4, 4.0)
    cl_h = mem_h + 1.0 + (NET_H + 0.9 if side else 0)
    cl_w = max(cl_w, len(side) * (NET_W + 0.4) + 0.4)
    row_w = lambda k: k * NET_W + max(k - 1, 0) * GAP
    side_w = 0
    content_w = max(row_w(len(top)), row_w(len(bottom)), cl_w) + side_w
    has_rt_top = any(n["routers"] for n in top)
    has_rt_bot = any(n["routers"] for n in bottom)

    page.width = content_w + 2 * margin
    y = margin
    y_rt_bot = None
    if has_rt_bot:
        y_rt_bot = y + 0.5
        y += 1.6
    y_bot = y + NET_H / 2
    y += NET_H + 1.6
    y_cl = y + cl_h / 2
    y += cl_h + 1.6
    y_top = y + NET_H / 2
    y += NET_H
    y_rt_top = None
    if has_rt_top:
        y_rt_top = y + 1.1
        y += 1.6
    y += 1.3   # titre + légende
    page.height = y + margin * 0.5

    cx_mid = margin + (content_w - side_w) / 2
    # Titre
    page.box(page.width / 2, page.height - 0.55, min(page.width - 1, 9), 0.5,
             "Topologie du cluster %s — %d membre(s), %d réseau(x)" % (cl["name"], len(members), len(nets)),
             "#FFFFFF", "#FFFFFF", size_pt=14, bold=True, lw_pt=0.1)

    # Cluster (cadre) + membres
    fill, line = STYLE["cluster"]
    mode = next((v for k, v in cl["props"].items() if "mode" in k or "high availability" in k), "")
    cl_text = cl["name"] + (" (%s)" % cl["type"] if cl["type"] else "") + (("\n" + mode) if mode else "")
    cl_shape = page.box(cx_mid, y_cl, cl_w, cl_h, cl_text, fill, line, size_pt=11, bold=True,
                        valign=0, lw_pt=2, back=True, name="Cluster " + cl["name"])
    mem_shapes = {}
    x0 = cx_mid - (n_mem * (MEM_W + 0.4) - 0.4) / 2 + MEM_W / 2
    for k, m in enumerate(members):
        f2, l2 = STYLE["member"]
        mem_y = y_cl + cl_h / 2 - 0.75 - mem_h / 2
        mem_shapes[m["name"]] = page.box(x0 + k * (MEM_W + 0.4), mem_y, MEM_W, mem_h,
                                         "\n".join(mem_lines(m)), f2, l2, size_pt=7.5,
                                         lw_pt=1.25, name=m["name"])

    def net_text(n):
        lines = [n["network"], LABELS.get(n["class"], n["class"])]
        for v in n["vip"][:2]:
            lines.append("VIP %s (%s)" % (v["ip"], v["if"]))
        return "\n".join(lines)

    def place_row(row, yy):
        shapes = []
        x = cx_mid - row_w(len(row)) / 2 + NET_W / 2
        for n in row:
            f, l = STYLE.get(n["class"], STYLE["unknown"])
            s = page.box(x, yy, NET_W, NET_H, net_text(n), f, l, size_pt=8.5, name=n["network"])
            shapes.append((n, s))
            x += NET_W + GAP
        return shapes

    def label(n):
        ifs = sorted({v["if"] for v in n["vip"]} | {mm["if"] for mm in n["members"]})
        mem = " ".join("%s" % short(mm["ip"]) for mm in n["members"])
        return "%s\n%s" % ("/".join(ifs), mem) if mem else "/".join(ifs)

    for (n, s), yy in [(p, y_top) for p in place_row(top, y_top)] + \
                      [(p, y_bot) for p in place_row(bottom, y_bot)]:
        f, l = STYLE.get(n["class"], STYLE["unknown"])
        target = cl_shape if n["vip"] or not n["members"] else mem_shapes.get(n["members"][0]["member"], cl_shape)
        page.connector(s, target, label(n), color=l, lw_pt=2)
        if n["routers"]:
            ry = y_rt_top if yy == y_top else y_rt_bot
            k = len(n["routers"])
            for j, (nh, dests) in enumerate(n["routers"].items()):
                rx = s[1][0] + (j - (k - 1) / 2) * 1.5
                txt = "Routeur %s\n%s" % (nh, "\n".join(dests[:5]) + ("\n..." if len(dests) > 5 else ""))
                fr, lr = STYLE["router"]
                r = page.box(rx, ry, 1.4, 0.75, txt, fr, lr, size_pt=7, ellipse=True, name="GW " + nh)
                page.connector(s, r, "", color=lr, lw_pt=1)

    # Réseaux sync / gestion dans le cadre du cluster, sous les membres
    xs = cx_mid - (len(side) * (NET_W + 0.4) - 0.4) / 2 + NET_W / 2
    yy = y_cl - cl_h / 2 + 0.25 + NET_H / 2
    for n in side:
        f, l = STYLE[n["class"]]
        s = page.box(xs, yy, NET_W, NET_H, net_text(n), f, l, size_pt=8.5, name=n["network"])
        for mm in n["members"]:
            if mm["member"] in mem_shapes:
                page.connector(s, mem_shapes[mm["member"]], "%s %s" % (mm["if"], short(mm["ip"])),
                               color=l, lw_pt=1.25, dash=True)
        xs += NET_W + 0.4

    # Légende (en haut à gauche)
    for k, cls in enumerate(["external", "internal", "dmz", "sync", "mgmt"]):
        f, l = STYLE[cls]
        page.box(margin * 0.5 + 0.45 + k * 1.0, page.height - 1.15, 0.9, 0.26, LABELS[cls], f, l, size_pt=7)
    return page



# --------------------------------------------------------------------------
# 7b. Analyse des couches (layers) et de leurs règles
# --------------------------------------------------------------------------

ANY_RE = re.compile(r"^\s*(any|\*any|tout)\s*$", re.I)


def is_any(v):
    return bool(ANY_RE.match((v or "").split("\n")[0])) and "\n" not in (v or "").strip()


def action_kind(a):
    a = norm(a)
    if any(k in a for k in ("drop", "reject", "block", "refus", "bloqu")):
        return "drop"
    if any(k in a for k in ("accept", "allow", "autoris")):
        return "accept"
    if "inline" in a or "apply layer" in a or "layer" in a:
        return "inline"
    return "other"


def analyse_layers(model, cluster_names=()):
    out = []
    names = list(model.layers)
    for ln in names:
        rules = [r for r in model.rules if r["layer"] == ln]
        st = {"name": ln, "rules": len(rules), "accept": 0, "drop": 0, "inline": 0,
              "other": 0, "disabled": 0, "sections": [], "children": [], "checks": []}
        for r in rules:
            k = action_kind(r.get("action", ""))
            child = next((n for n in names if n != ln and
                          re.search(r"(?<!\w)" + re.escape(norm(n)) + r"(?!\w)", norm(r.get("action", "")))), None)
            if child:
                k = "inline"
                st["children"].append((child, r.get("no", "?")))
            st[k] += 1
            st["disabled"] += r["disabled"]
            if r["section"] and r["section"] not in st["sections"]:
                st["sections"].append(r["section"])
            if r["disabled"]:
                continue
            tag = "règle %s%s" % (r.get("no", "?"), (" (%s)" % r["rname"]) if r.get("rname") else "")
            if k == "accept" and is_any(r.get("src")) and is_any(r.get("dst")) and \
                    (is_any(r.get("svc")) or not r.get("svc")):
                st["checks"].append("%s : Any → Any → Any en Accept" % tag)
            elif k == "accept" and (is_any(r.get("src")) or is_any(r.get("dst"))) and \
                    is_any(r.get("svc")):
                st["checks"].append("%s : Accept avec source ou destination Any et service Any" % tag)
            if "track" in r and k in ("accept", "drop") and norm(r["track"]) in ("", "none", "aucun", "-"):
                st["checks"].append("%s : aucune journalisation" % tag)
            if "hits" in r and re.fullmatch(r"\s*0\s*(\(.*\))?\s*", r["hits"] or "x"):
                st["checks"].append("%s : 0 hit (règle inutilisée ?)" % tag)
        if st["disabled"]:
            st["checks"].append("%d règle(s) désactivée(s)" % st["disabled"])
        enabled = [r for r in rules if not r["disabled"]]
        if enabled:
            last = enabled[-1]
            if action_kind(last.get("action", "")) != "drop" or not (
                    is_any(last.get("src")) and is_any(last.get("dst"))):
                st["checks"].append("Pas de règle de nettoyage (Any/Any/Drop) en fin de couche "
                                    "(dernière : règle %s, %s)" % (last.get("no", "?"), last.get("action", "")))
        out.append(st)
    child_names = {c for st in out for c, _ in st["children"]}
    for st in out:
        st["inline_layer"] = st["name"] in child_names
    return out


def draw_layers(layer_stats, title):
    page = VisioPage("Couches")
    BW, BH, VG = 3.0, 1.15, 0.8
    ordered = [s for s in layer_stats if not s["inline_layer"]]
    inline = [s for s in layer_stats if s["inline_layer"]]
    rows = max(len(ordered), 1)
    page.width = 2 * BW + 4.5
    page.height = rows * (BH + VG) + 2.4 + max(0, len(inline) - rows) * (BH + 0.4)
    page.box(page.width / 2, page.height - 0.5, page.width - 1, 0.5,
             "Couches de la politique — %s" % title, "#FFFFFF", "#FFFFFF",
             size_pt=14, bold=True, lw_pt=0.1)

    def txt(s):
        t = "%s\n%d règles : %d accept / %d drop" % (s["name"], s["rules"], s["accept"], s["drop"])
        if s["inline"] or s["children"]:
            t += " / %d inline" % max(s["inline"], len(s["children"]))
        if s["disabled"]:
            t += "\n%d désactivée(s)" % s["disabled"]
        if s["sections"]:
            t += "\n%d section(s)" % len(s["sections"])
        if s["checks"]:
            t += "\n⚠ %d point(s) d'attention" % len(s["checks"])
        return t

    shapes = {}
    x0, y = 0.9 + BW / 2, page.height - 1.4 - BH / 2
    prev = None
    for k, s in enumerate(ordered):
        sh = page.box(x0, y, BW, BH, ("%d. " % (k + 1)) + txt(s), "#DAE8FC", "#6C8EBF",
                      size_pt=9, bold=False, name="Layer " + s["name"])
        shapes[s["name"]] = sh
        if prev:
            page.connector(prev, sh, "Accept", color="#6C8EBF", lw_pt=1.5, arrow=True)
        prev, y = sh, y - (BH + VG)
    xi = x0 + BW + 1.8
    yi = page.height - 1.4 - BH / 2
    for s in inline:
        sh = page.box(xi, yi, BW, BH, "Inline : " + txt(s), "#FFF2CC", "#D6B656",
                      size_pt=9, name="Layer " + s["name"])
        shapes[s["name"]] = sh
        yi -= BH + 0.4
    for s in layer_stats:
        for child, no in s["children"]:
            if s["name"] in shapes and child in shapes:
                page.connector(shapes[s["name"]], shapes[child], "règle %s" % no,
                               color="#D6B656", lw_pt=1.25, dash=True, arrow=True)
    return page


def write_rules_csv(model, path):
    import csv
    cols = ["layer", "section", "no", "rname", "src", "dst", "vpn", "svc", "content",
            "action", "track", "install", "time", "hits", "comment", "disabled"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Couche", "Section", "N°", "Nom", "Source", "Destination", "VPN",
                    "Services", "Contenu", "Action", "Track", "Installé sur", "Horaire",
                    "Hits", "Commentaire", "Désactivée"])
        for r in model.rules:
            w.writerow([str(r.get(c2, "")).replace("\n", ", ") for c2 in cols])

# --------------------------------------------------------------------------
# 8. Rapport
# --------------------------------------------------------------------------

def report_md(topos, model, layer_stats=()):
    L = ["# Analyse de configuration Check Point", ""]
    if layer_stats:
        L += ["## Couches de politique", "",
              "| # | Couche | Type | Règles | Accept | Drop | Inline | Désactivées | Sections |",
              "|---|---|---|---|---|---|---|---|---|"]
        for k, st in enumerate(layer_stats, 1):
            L.append("| %d | %s | %s | %d | %d | %d | %d | %d | %d |" % (
                k, st["name"], "inline" if st["inline_layer"] else "ordonnée", st["rules"],
                st["accept"], st["drop"], len(st["children"]) or st["inline"],
                st["disabled"], len(st["sections"])))
        for st in layer_stats:
            L += ["", "### Couche %s" % st["name"], ""]
            if st["sections"]:
                L.append("Sections : " + ", ".join(st["sections"]))
                L.append("")
            L += ["- " + ch for ch in st["checks"]] or ["- Aucun point d'attention"]
        L.append("")
    for t in topos:
        cl = t["cluster"]
        L += ["## Cluster %s" % cl["name"], ""]
        if cl["type"]:
            L.append("- Type : %s" % cl["type"])
        if cl["ip"]:
            L.append("- Adresse principale : %s" % cl["ip"])
        L.append("- Membres : %s" % (", ".join(m["name"] for m in t["members"]) or "aucun détecté"))
        L += ["", "### Interfaces virtuelles (VIP)", "",
              "| Interface | VIP | Réseau | Topologie |", "|---|---|---|---|"]
        for i in cl["interfaces"]:
            L.append("| %s | %s | %s | %s |" % (i["name"], i["ip"], net_of(i) or "?",
                                               i["topology_raw"] or LABELS.get(i["topology"], "")))
        for m in t["members"]:
            L += ["", "### Membre %s" % m["name"], "",
                  "| Interface | IP | Réseau |", "|---|---|---|"]
            for i in m["interfaces"]:
                L.append("| %s | %s | %s |" % (i["name"], i["ip"], net_of(i) or "?"))
        L += ["", "### Réseaux", "", "| Réseau | Rôle | VIP | Membres | Routeurs |",
              "|---|---|---|---|---|"]
        for n in t["networks"]:
            L.append("| %s | %s | %s | %s | %s |" % (
                n["network"], LABELS.get(n["class"], n["class"]),
                ", ".join(v["ip"] for v in n["vip"]),
                ", ".join("%s %s" % (mm["member"], mm["ip"]) for mm in n["members"]),
                "; ".join("%s → %s" % (k, ", ".join(v)) for k, v in n["routers"].items())))
        L += ["", "### Points d'attention", ""]
        L += ["- " + ch for ch in t["checks"]] or ["- Aucun"]
        L.append("")
    return "\n".join(L)


def serialise(topos):
    out = []
    for t in topos:
        out.append({"cluster": {k: v for k, v in t["cluster"].items()},
                    "members": t["members"], "networks": t["networks"],
                    "checks": t["checks"]})
    return out


# --------------------------------------------------------------------------
# 9. Programme principal
# --------------------------------------------------------------------------

def read_html(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def collect_inputs(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, names in os.walk(p):
                files += [os.path.join(root, n) for n in sorted(names)
                          if n.lower().endswith((".html", ".htm"))]
        else:
            files.append(p)
    return files


def main(argv=None):
    ap = argparse.ArgumentParser(description="Export HTML Check Point -> topologie Visio")
    ap.add_argument("inputs", nargs="+", help="fichier(s) HTML ou dossier d'export")
    ap.add_argument("-o", "--output", help="fichier .vsdx de sortie")
    ap.add_argument("--report", help="rapport Markdown (défaut : <sortie>.md)")
    ap.add_argument("--json", help="modèle extrait au format JSON")
    ap.add_argument("--cluster", help="forcer le nom de l'objet cluster")
    ap.add_argument("--debug", action="store_true", help="afficher les tableaux détectés")
    a = ap.parse_args(argv)

    files = collect_inputs(a.inputs)
    if not files:
        sys.exit("Aucun fichier HTML trouvé.")
    model = build_model([(f, read_html(f)) for f in files])

    if a.debug:
        print("=== Tableaux détectés ===")
        for t in model.tables_log:
            print("[%s] %-11s %3d lignes | %s\n      entête=%s colonnes=%s" % (
                t["source"], t["interprétation"], t["lignes"], t["contexte"],
                t["entête"][:8], t["colonnes"]))
        print("=== Objets ===")
        for o in model.objects.values():
            print("  %-25s kind=%-8s type=%-25s itf=%d membres=%s" % (
                o["name"], o["kind"], o["type"][:25], len(o["interfaces"]), o["members"]))

    topos = build_topology(model, a.cluster)
    topos = [t for t in topos if t["cluster"]["interfaces"] or t["members"]]
    layer_stats = analyse_layers(model, [t["cluster"]["name"] for t in topos])
    if not topos and not layer_stats:
        sys.exit("Ni topologie ni règles trouvées. Relancez avec --debug pour voir "
                 "les tableaux détectés et ajustez FIELD_ALIASES / RULE_ALIASES.")
    if not topos:
        print("Attention : aucune interface de cluster trouvée (seules les couches seront dessinées).")

    out = a.output or os.path.splitext(files[0])[0] + "_topologie.vsdx"
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    pages = [draw_topology(t) for t in topos]
    if layer_stats:
        pages.append(draw_layers(layer_stats, ", ".join(t["cluster"]["name"] for t in topos)))
    write_vsdx(pages, out)
    rep = a.report or os.path.splitext(out)[0] + ".md"
    with open(rep, "w", encoding="utf-8") as f:
        f.write(report_md(topos, model, layer_stats))
    if model.rules:
        csv_path = os.path.splitext(out)[0] + "_regles.csv"
        write_rules_csv(model, csv_path)
        print("Règles  : %s (%d règles, %d couche(s))" % (csv_path, len(model.rules), len(layer_stats)))
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            data = {"clusters": serialise(topos), "layers": layer_stats, "rules": model.rules}
            json.dump(data, f, indent=2, ensure_ascii=False, default=str)

    for t in topos:
        print("Cluster %s : %d membre(s), %d réseau(x), %d point(s) d'attention" % (
            t["cluster"]["name"], len(t["members"]), len(t["networks"]), len(t["checks"])))
    print("Visio   : %s\nRapport : %s" % (out, rep))


if __name__ == "__main__":
    main()
