#!/usr/bin/env python3
"""
topologie.py : reconstitue la topologie d'un réseau HP Comware à partir d'une
arborescence de dossiers (un dossier = un équipement, un sous-dossier = un
équipement câblé physiquement sous son parent).

Chaque dossier contient :
  - <SYSNAME>.txt   : la configuration (utilisée seulement pour contrôler les dates)
  - <SYSNAME>.xlsx  : le rapport du parseur (Synthese, Interfaces_L2, Interfaces_L3, Management)

Sorties :
  - topologie.xlsx  : Lignée, Équipements, Liens, Subnets, VLANs, East-West, Anomalies
  - topologie.drawio : page Synthèse + une page par bloc (+ vue complète si le parc est petit),
                       calques Base / Ports / Subnets / VLANs / East-West

Usage :
  python3 topologie.py RACINE [-o DOSSIER_SORTIE] [--profondeur-bloc 1]

Principes :
  - l'arbre des dossiers fait foi pour les liens physiques (graphe des équipements = arbre) ;
  - les subnets ajoutés par-dessus donnent un DAG orienté du cœur vers les feuilles ;
  - le script ne devine jamais : toute règle qui échoue est consignée dans l'onglet Anomalies.

Dépendance : openpyxl.
"""
from __future__ import annotations

import argparse
import ipaddress
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===========================================================================
# PARAMÈTRES : à adapter au format réel du parseur
# ===========================================================================

# Noms des onglets du rapport
FEUILLES = {
    "synthese": "Synthese",
    "l2": "Interfaces_L2",
    "l3": "Interfaces_L3",
    "mgmt": "Management",
}

# Champs de l'onglet Synthese (format "Champ | Valeur", ou une ligne d'en-têtes)
SYNTHESE = {
    "sysname": "Sysname",
    "modele": "Modèle",
    "irf": "Membres IRF",
    "date_collecte": "Date de collecte",
}

# Colonnes des onglets d'interfaces (clé interne -> en-tête dans le rapport)
COLONNES = {
    "l2": {
        "interface": "Interface",
        "description": "Description",
        "link_type": "Link-type",
        "vlan_access": "VLAN access",
        "vlans": "VLANs autorisés",
        "groupe": "Groupe agrégat",
    },
    "l3": {
        "interface": "Interface",
        "description": "Description",
        "vlan": "VLAN",
        "ip": "Adresse IP",
        "masque": "Masque",
        "vrf": "VPN-instance",
        "vrrp_vrid": "VRRP VRID",
        "vrrp_vip": "VRRP VIP",
        "groupe": "Groupe agrégat",
    },
    "mgmt": {
        "interface": "Interface",
        "vlan": "VLAN",
        "ip": "Adresse IP",
        "masque": "Masque",
        "passerelle": "Passerelle par défaut",
    },
}

# Libellés des niveaux selon la profondeur dans l'arbre (au-delà : "Niveau N")
NIVEAUX = ["Cœur", "Distribution", "Accès"]

# draw.io : au-delà de ce nombre d'équipements, pas de page "Vue complète"
SEUIL_VUE_COMPLETE = 150
# draw.io : nombre max de subnets affichés à côté d'un équipement
MAX_SUBNETS_AFFICHES = 12

# ===========================================================================
# Modèle de données
# ===========================================================================


@dataclass
class Iface:
    name: str
    desc: str = ""
    kind: str = "phys"            # phys | bagg | ragg | svi | loop
    layer: str = "L2"             # L2 | L3
    link_type: str = ""           # access | trunk | hybrid
    access_vlan: int | None = None
    vlans: set[int] | None = None  # VLANs autorisés (trunk / hybrid)
    group: int | None = None      # numéro de groupe d'agrégation
    ip: str = ""
    network: ipaddress.IPv4Network | None = None
    vrf: str = ""
    vlan: int | None = None       # VLAN d'une Vlan-interface
    vrrp_vrid: str = ""
    vrrp_vip: str = ""


@dataclass
class Device:
    id: str                       # chemin relatif du dossier (unique)
    name: str                     # nom du dossier (= sysname attendu)
    folder: Path
    parent: str | None = None
    children: list[str] = field(default_factory=list)
    depth: int = 0
    path: list[str] = field(default_factory=list)   # noms de la racine jusqu'à lui
    statut: str = "OK"            # OK | Opaque
    sysname: str = ""
    modele: str = ""
    irf: str = ""
    date_collecte: str = ""
    date_txt: str = ""
    l2: list[Iface] = field(default_factory=list)
    l3: list[Iface] = field(default_factory=list)
    mgmt: list[dict] = field(default_factory=list)

    @property
    def role(self) -> str:
        if self.statut != "OK":
            return "Opaque"
        return "L3" if any(i.network for i in self.l3) else "L2"


@dataclass
class Port:
    name: str = ""
    members: list[str] = field(default_factory=list)
    vlans: set[int] | None = None
    methode: str = ""


@dataclass
class Link:
    parent: str
    child: str
    type: str = "L2"
    intercos: list[tuple] = field(default_factory=list)  # (clé subnet, iface parent, iface enfant)
    p_port: Port = field(default_factory=Port)
    c_port: Port = field(default_factory=Port)
    vlans_eff: set[int] | None = None
    vlans_verifies: bool = True
    statut: str = "OK"


class Anomalies:
    def __init__(self):
        self.rows: list[list] = []

    def add(self, gravite, categorie, equipement, objet, detail):
        self.rows.append([gravite, categorie, equipement, objet, detail])


# ===========================================================================
# Utilitaires
# ===========================================================================

ABREVIATIONS = [
    ("Ten-GigabitEthernet", "XGE"), ("Twenty-FiveGigE", "WGE"), ("FortyGigE", "FGE"),
    ("HundredGigE", "HGE"), ("M-GigabitEthernet", "MGE"), ("GigabitEthernet", "GE"),
    ("Bridge-Aggregation", "BAGG"), ("Route-Aggregation", "RAGG"),
    ("Vlan-interface", "Vlan"), ("LoopBack", "Loop"),
]


def court(nom: str) -> str:
    for long_, abr in ABREVIATIONS:
        if nom.startswith(long_):
            return abr + nom[len(long_):]
    return nom


def kind_of(nom: str) -> str:
    n = nom.lower()
    if n.startswith("bridge-aggregation"):
        return "bagg"
    if n.startswith("route-aggregation"):
        return "ragg"
    if n.startswith("vlan-interface"):
        return "svi"
    if n.startswith("loopback"):
        return "loop"
    return "phys"


def num_suffix(nom: str) -> int | None:
    m = re.search(r"(\d+)$", nom)
    return int(m.group(1)) if m else None


def to_int(v) -> int | None:
    if v is None or str(v).strip() == "":
        return None
    try:
        return int(float(str(v).strip()))
    except ValueError:
        return None


def txt(v) -> str:
    return "" if v is None else str(v).strip()


def parse_vlans(v) -> set[int]:
    """'110 120 130' / '2 to 10' / '2-10' / '1,5'"""
    s = txt(v).replace(",", " ")
    s = re.sub(r"\s+to\s+", "-", s, flags=re.I)
    out: set[int] = set()
    for tok in s.split():
        if "-" in tok:
            a, _, b = tok.partition("-")
            if a.isdigit() and b.isdigit():
                out.update(range(int(a), int(b) + 1))
        elif tok.isdigit():
            out.add(int(tok))
    return out


def fmt_vlans(vs: set[int] | None) -> str:
    if not vs:
        return ""
    s = sorted(vs)
    parts, start, prev = [], s[0], s[0]
    for v in s[1:] + [None]:
        if v is not None and v == prev + 1:
            prev = v
            continue
        parts.append(str(start) if start == prev else f"{start}-{prev}")
        if v is not None:
            start = prev = v
    return " ".join(parts)


def mentionne(desc: str, nom: str) -> bool:
    """La description cite-t-elle ce nom d'équipement (avec frontières de mot) ?"""
    if not desc or not nom:
        return False
    return re.search(r"(?<![A-Za-z0-9])" + re.escape(nom) + r"(?![A-Za-z0-9])", desc, re.I) is not None


def cle_txt(cle) -> str:
    vrf, net = cle
    return f"{net}" + (f" [{vrf}]" if vrf else "")


# ===========================================================================
# 1. Parcours de l'arborescence
# ===========================================================================


def parcourir(racine: Path, anomalies: Anomalies) -> dict[str, Device]:
    devices: dict[str, Device] = {}

    def est_ignore(p: Path) -> bool:
        return p.name.startswith((".", "_", "~"))

    def visiter(dossier: Path, parent: str | None, depth: int, chemin: list[str]):
        dev_id = str(dossier.relative_to(racine)).replace("\\", "/")
        dev = Device(id=dev_id, name=dossier.name, folder=dossier, parent=parent,
                     depth=depth, path=chemin + [dossier.name])
        devices[dev_id] = dev
        if parent:
            devices[parent].children.append(dev_id)
        for sous in sorted(p for p in dossier.iterdir() if p.is_dir() and not est_ignore(p)):
            visiter(sous, dev_id, depth + 1, dev.path)

    # La racine est un équipement si elle contient elle-même une config,
    # sinon chacun de ses sous-dossiers est une racine (forêt possible).
    if list(racine.glob("*.txt")) or list(racine.glob("*.xlsx")):
        visiter(racine, None, 0, [])
    else:
        for sous in sorted(p for p in racine.iterdir() if p.is_dir() and not est_ignore(p)):
            visiter(sous, None, 0, [])

    # Doublons de noms (ambiguïté pour les descriptions)
    par_nom = defaultdict(list)
    for d in devices.values():
        par_nom[d.name.lower()].append(d.id)
    for nom, ids in par_nom.items():
        if len(ids) > 1:
            anomalies.add("Erreur", "Arborescence", devices[ids[0]].name, "Nom en double",
                          "Même nom de dossier à plusieurs endroits : " + " ; ".join(ids))
    return devices


# ===========================================================================
# 2. Lecture des rapports
# ===========================================================================

_colonnes_manquantes: set[tuple[str, str]] = set()


def lire_table(ws, onglet: str) -> list[dict]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [txt(c).lower() for c in rows[0]]
    mapping = {}
    for cle, entete in COLONNES[onglet].items():
        if entete.lower() in header:
            mapping[cle] = header.index(entete.lower())
        else:
            _colonnes_manquantes.add((FEUILLES[onglet], entete))
    out = []
    for r in rows[1:]:
        if r is None or all(v in (None, "") for v in r):
            continue
        out.append({cle: (r[i] if i < len(r) else None) for cle, i in mapping.items()})
    return out


def lire_synthese(ws) -> dict:
    rows = [r for r in ws.iter_rows(values_only=True) if r]
    if not rows:
        return {}
    header = [txt(c).lower() for c in rows[0]]
    # Format en ligne : en-têtes = champs
    if SYNTHESE["sysname"].lower() in header and len(rows) > 1:
        return {k: rows[1][header.index(v.lower())] for k, v in SYNTHESE.items()
                if v.lower() in header}
    # Format clé / valeur
    kv = {txt(r[0]).lower(): (r[1] if len(r) > 1 else None) for r in rows}
    return {k: kv.get(v.lower()) for k, v in SYNTHESE.items()}


def construire_iface(row: dict, couche: str) -> Iface:
    nom = txt(row.get("interface"))
    it = Iface(name=nom, desc=txt(row.get("description")), kind=kind_of(nom), layer=couche,
               group=to_int(row.get("groupe")))
    if it.kind in ("bagg", "ragg") and it.group is None:
        it.group = num_suffix(nom)
    if couche == "L2":
        it.link_type = txt(row.get("link_type")).lower()
        it.access_vlan = to_int(row.get("vlan_access"))
        if it.link_type in ("trunk", "hybrid"):
            it.vlans = parse_vlans(row.get("vlans"))
    else:
        it.vlan = to_int(row.get("vlan"))
        if it.kind == "svi" and it.vlan is None:
            it.vlan = num_suffix(nom)
        it.vrf = txt(row.get("vrf"))
        it.vrrp_vrid = txt(row.get("vrrp_vrid"))
        it.vrrp_vip = txt(row.get("vrrp_vip"))
        ip, masque = txt(row.get("ip")), txt(row.get("masque"))
        if ip:
            try:
                itf = ipaddress.ip_interface(ip if "/" in ip else f"{ip}/{masque or 32}")
                it.ip, it.network = str(itf.ip), itf.network
            except ValueError:
                it.ip = ip
    return it


def lire_rapports(devices: dict[str, Device], anomalies: Anomalies):
    for dev in devices.values():
        txts = sorted(dev.folder.glob("*.txt"))
        xlsxs = sorted(p for p in dev.folder.glob("*.xlsx") if not p.name.startswith("~$"))
        txt_f = next((p for p in txts if p.stem.lower() == dev.name.lower()), txts[0] if txts else None)
        xls_f = next((p for p in xlsxs if p.stem.lower() == dev.name.lower()), xlsxs[0] if xlsxs else None)

        if len(txts) > 1:
            anomalies.add("Avertissement", "Fichiers", dev.name, "Plusieurs .txt",
                          f"Fichier retenu : {txt_f.name}")
        if txt_f:
            dev.date_txt = datetime.fromtimestamp(txt_f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        else:
            anomalies.add("Avertissement", "Fichiers", dev.name, "Config absente",
                          "Aucun fichier .txt dans le dossier")
        if not xls_f:
            dev.statut = "Opaque"
            anomalies.add("Erreur", "Fichiers", dev.name, "Rapport absent",
                          "Aucun .xlsx : équipement conservé comme nœud opaque")
            continue
        if txt_f and txt_f.stat().st_mtime > xls_f.stat().st_mtime + 1:
            anomalies.add("Avertissement", "Fichiers", dev.name, "Rapport périmé",
                          f"{txt_f.name} est plus récent que {xls_f.name} : relancer le parseur")
        try:
            wb = load_workbook(xls_f, read_only=True, data_only=True)
        except Exception as e:  # noqa: BLE001
            dev.statut = "Opaque"
            anomalies.add("Erreur", "Fichiers", dev.name, "Rapport illisible", f"{xls_f.name} : {e}")
            continue

        manquants = [f for f in FEUILLES.values() if f not in wb.sheetnames]
        if manquants:
            anomalies.add("Erreur", "Fichiers", dev.name, "Onglet absent", ", ".join(manquants))

        if FEUILLES["synthese"] in wb.sheetnames:
            s = lire_synthese(wb[FEUILLES["synthese"]])
            dev.sysname = txt(s.get("sysname"))
            dev.modele = txt(s.get("modele"))
            dev.irf = txt(s.get("irf"))
            dev.date_collecte = txt(s.get("date_collecte"))
        if FEUILLES["l2"] in wb.sheetnames:
            dev.l2 = [construire_iface(r, "L2") for r in lire_table(wb[FEUILLES["l2"]], "l2")
                      if txt(r.get("interface"))]
        if FEUILLES["l3"] in wb.sheetnames:
            dev.l3 = [construire_iface(r, "L3") for r in lire_table(wb[FEUILLES["l3"]], "l3")
                      if txt(r.get("interface"))]
            for i in dev.l3:
                if i.ip and not i.network:
                    anomalies.add("Avertissement", "Rapport", dev.name, i.name,
                                  f"Adresse IP invalide : {i.ip}")
        if FEUILLES["mgmt"] in wb.sheetnames:
            dev.mgmt = lire_table(wb[FEUILLES["mgmt"]], "mgmt")
        wb.close()

        if dev.sysname and dev.sysname.lower() != dev.name.lower():
            anomalies.add("Avertissement", "Arborescence", dev.name, "Sysname différent",
                          f"Dossier « {dev.name} », sysname du rapport « {dev.sysname} »")

    for onglet, entete in sorted(_colonnes_manquantes):
        anomalies.add("Avertissement", "Rapport", "(tous)", f"Colonne absente : {onglet}",
                      f"« {entete} » introuvable : vérifier COLONNES en tête de script")


# ===========================================================================
# 3. Analyse : subnets, liens, ports, VLANs, passerelles
# ===========================================================================


def membres(dev: Device, port: Iface) -> list[str]:
    if port.kind not in ("bagg", "ragg") or port.group is None:
        return []
    pool = dev.l2 if port.kind == "bagg" else dev.l3
    return [i.name for i in pool if i.kind == "phys" and i.group == port.group]


def vlans_port(dev: Device, port: Iface) -> set[int] | None:
    if port.vlans is not None:
        return port.vlans
    if port.kind == "bagg":
        vs = [i.vlans for i in dev.l2 if i.kind == "phys" and i.group == port.group and i.vlans is not None]
        if vs:
            return set().union(*vs)
    return None


def ports_logiques_l2(dev: Device) -> list[Iface]:
    """Agrégats L2 + ports physiques L2 hors agrégat."""
    baggs = {i.group for i in dev.l2 if i.kind == "bagg"}
    return [i for i in dev.l2 if i.kind == "bagg" or (i.kind == "phys" and (i.group is None or i.group not in baggs))]


def trouver_port_l2(dev: Device, voisin: str, cote: str, anomalies: Anomalies) -> tuple[Iface | None, str]:
    """Port L2 de `dev` vers `voisin`. cote = 'parent' ou 'enfant'."""
    baggs = {i.group: i for i in dev.l2 if i.kind == "bagg"}
    cands: dict[str, Iface] = {}
    for i in dev.l2:
        if mentionne(i.desc, voisin):
            logique = baggs.get(i.group, i) if i.kind == "phys" and i.group in baggs else i
            cands[logique.name] = logique
    if len(cands) == 1:
        return next(iter(cands.values())), "description"
    if len(cands) > 1:
        anomalies.add("Avertissement", "Ports", dev.name, f"Port vers {voisin}",
                      "Plusieurs ports candidats par description : " + ", ".join(court(n) for n in cands))
        return None, "ambigu"
    if cote == "enfant":
        uplinks = [i for i in ports_logiques_l2(dev) if i.link_type in ("trunk", "hybrid")]
        if len(uplinks) == 1:
            return uplinks[0], "uplink unique"
    return None, "non identifié"


def analyser(devices: dict[str, Device], anomalies: Anomalies) -> dict:
    by_name = {d.name: d for d in devices.values()}

    # ---- Index des subnets (management exclu : il est dans un onglet à part) ----
    index: dict[tuple, list[tuple[str, Iface]]] = defaultdict(list)
    for d in devices.values():
        for i in d.l3:
            if i.network:
                index[(i.vrf, i.network)].append((d.id, i))

    def sur(cle, dev_id):
        return [i for did, i in index[cle] if did == dev_id]

    vrrp_cles = {cle for cle, lst in index.items() if any(i.vrrp_vip or i.vrrp_vrid for _, i in lst)}

    # ---- Liens parent -> enfant ----
    links: dict[str, Link] = {}
    cles_interco: set[tuple] = set()
    for d in devices.values():
        if not d.parent:
            continue
        p = devices[d.parent]
        lk = Link(parent=p.id, child=d.id)
        links[d.id] = lk
        if p.statut != "OK" or d.statut != "OK":
            lk.statut = "Non vérifiable (rapport manquant)"
            lk.type = "?"
            continue

        cles_p = {cle for cle, lst in index.items() if any(did == p.id for did, _ in lst)}
        cles_c = {cle for cle, lst in index.items() if any(did == d.id for did, _ in lst)}
        for cle in sorted(cles_p & cles_c - vrrp_cles, key=lambda c: (c[0], c[1])):
            lk.intercos.append((cle, sur(cle, p.id)[0], sur(cle, d.id)[0]))
            cles_interco.add(cle)

        transit_svi = lk.intercos and all(ip.kind == "svi" and ic.kind == "svi" for _, ip, ic in lk.intercos)
        if lk.intercos and not transit_svi:
            # Interco L3 routée : ports = interfaces portant le subnet
            lk.type = "L3"
            _, ip_, ic_ = lk.intercos[0]
            lk.p_port = Port(ip_.name, membres(p, ip_), None, "subnet commun")
            lk.c_port = Port(ic_.name, membres(d, ic_), None, "subnet commun")
            continue

        lk.type = "L3 (VLAN de transit)" if transit_svi else "L2"
        pp, mp = trouver_port_l2(p, d.name, "parent", anomalies)
        pc, mc = trouver_port_l2(d, p.name, "enfant", anomalies)
        if pp:
            lk.p_port = Port(pp.name, membres(p, pp), vlans_port(p, pp), mp)
        else:
            lk.p_port = Port("", [], None, mp)
            anomalies.add("Avertissement", "Ports", p.name, f"Downlink vers {d.name}",
                          f"Port côté parent {mp} (aucune description citant {d.name})")
        if pc:
            lk.c_port = Port(pc.name, membres(d, pc), vlans_port(d, pc), mc)
        else:
            lk.c_port = Port("", [], None, mc)
            anomalies.add("Avertissement", "Ports", d.name, f"Uplink vers {p.name}",
                          f"Port côté enfant {mc}")
        if not pp and not pc:
            lk.statut = "Lien non retrouvé dans les configs"
            anomalies.add("Erreur", "Liens", d.name, f"Lien {p.name} → {d.name}",
                          "Ni subnet commun, ni port L2 identifiable des deux côtés")

        vp, vc = lk.p_port.vlans, lk.c_port.vlans
        if vp is not None and vc is not None:
            lk.vlans_eff = vp & vc
            if vp - vc:
                anomalies.add("Avertissement", "VLANs", d.name, f"Lien {p.name} → {d.name}",
                              f"VLANs autorisés côté parent ({court(lk.p_port.name)}) mais pas côté enfant : {fmt_vlans(vp - vc)}")
            if vc - vp:
                anomalies.add("Avertissement", "VLANs", d.name, f"Lien {p.name} → {d.name}",
                              f"VLANs autorisés côté enfant ({court(lk.c_port.name)}) mais pas côté parent : {fmt_vlans(vc - vp)}")
        else:
            lk.vlans_eff = vp if vp is not None else vc
            lk.vlans_verifies = False

    # ---- Classification des subnets et liens east-west ----
    east_west = []
    types_subnet: dict[tuple, str] = {}
    for cle, lst in index.items():
        devs = {did for did, _ in lst}
        if cle in vrrp_cles:
            types_subnet[cle] = "User (VRRP)"
        elif cle in cles_interco:
            types_subnet[cle] = "Interco"
        elif len(devs) == 1:
            i = lst[0][1]
            types_subnet[cle] = "Loopback" if i.kind == "loop" else "User"
        else:
            types_subnet[cle] = "East-West"
        if cle in vrrp_cles or len(devs) < 2:
            continue
        ids = sorted(devs)
        for a_i in range(len(ids)):
            for b_i in range(a_i + 1, len(ids)):
                a, b = devices[ids[a_i]], devices[ids[b_i]]
                if a.parent == b.id or b.parent == a.id:
                    continue
                ia, ib = sur(cle, a.id)[0], sur(cle, b.id)[0]
                east_west.append([a.name, court(ia.name), ia.ip, b.name, court(ib.name), ib.ip,
                                  str(cle[1]), cle[0]])
                anomalies.add("Info", "East-West", a.name, f"{a.name} ↔ {b.name}",
                              f"Subnet {cle_txt(cle)} partagé hors arbre : signalé, exclu du graphe")
        if types_subnet[cle] == "Interco" and len(devs) > 2:
            anomalies.add("Avertissement", "Subnets", "", cle_txt(cle),
                          f"Interco présente sur {len(devs)} équipements")

    # ---- Présence des VLANs et recherche de passerelle ----
    svi_par_dev: dict[str, dict[int, list[Iface]]] = defaultdict(lambda: defaultdict(list))
    for d in devices.values():
        for i in d.l3:
            if i.kind == "svi" and i.vlan is not None and i.network and types_subnet.get((i.vrf, i.network)) != "Interco":
                svi_par_dev[d.id][i.vlan].append(i)

    vlan_rows = []
    services = []  # (dev_id servi, vlan, gw_id, [ifaces SVI], chemin L2, vérifié)
    for d in devices.values():
        acces: dict[int, list[str]] = defaultdict(list)
        trunks: dict[int, list[str]] = defaultdict(list)
        for i in d.l2:
            if i.link_type == "access" and i.access_vlan:
                acces[i.access_vlan].append(court(i.name))
        for i in ports_logiques_l2(d):
            for v in (vlans_port(d, i) or set()):
                trunks[v].append(court(i.name))
        for v, ifs in sorted(acces.items()):
            vlan_rows.append([d.name, v, "Accès", ", ".join(ifs)])
        for v, ifs in sorted(trunks.items()):
            vlan_rows.append([d.name, v, "Trunk", ", ".join(ifs)])
        for v, ifs in sorted(svi_par_dev[d.id].items()):
            vlan_rows.append([d.name, v, "SVI", ", ".join(court(i.name) for i in ifs)])

        for v in sorted(acces):
            cur, chemin, verifie = d, [d.name], True
            while True:
                if svi_par_dev[cur.id].get(v):
                    services.append((d.id, v, cur.id, svi_par_dev[cur.id][v], chemin, verifie))
                    break
                if not cur.parent:
                    anomalies.add("Erreur", "Passerelle", d.name, f"VLAN {v}",
                                  "Servi en accès mais aucune passerelle trouvée jusqu'à la racine")
                    break
                lk = links[cur.id]
                par = devices[cur.parent]
                if lk.type == "L3":
                    anomalies.add("Erreur", "Passerelle", d.name, f"VLAN {v}",
                                  f"Servi en accès, aucune SVI avant le lien L3 {cur.name} → {par.name}")
                    break
                if lk.vlans_eff is None:
                    verifie = False
                elif v not in lk.vlans_eff:
                    anc = []
                    a = par
                    while a:
                        if svi_par_dev[a.id].get(v):
                            anc.append(a.name)
                        a = devices[a.parent] if a.parent else None
                    suite = (f" (une SVI {v} existe plus haut sur {', '.join(anc)})" if anc
                             else f" ; aucune SVI {v} chez les ancêtres")
                    anomalies.add("Erreur", "Passerelle", d.name, f"VLAN {v}",
                                  f"Servi en accès mais bloqué sur le lien {cur.name} → {par.name}{suite}")
                    break
                cur = par
                chemin.append(cur.name)

    return dict(index=index, types=types_subnet, links=links, east_west=east_west,
                vlan_rows=vlan_rows, services=services, by_name=by_name)


# ===========================================================================
# 4. Export Excel
# ===========================================================================

HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="305496")
FILL_GRAV = {"Erreur": PatternFill("solid", fgColor="F8CBAD"),
             "Avertissement": PatternFill("solid", fgColor="FFE699"),
             "Info": PatternFill("solid", fgColor="DDEBF7")}


def write_sheet(ws, headers, rows, fill_col=None):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font, cell.fill = HEAD_FONT, HEAD_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    widths = [len(str(h)) for h in headers]
    for r in rows:
        ws.append(r)
        for k, v in enumerate(r):
            widths[k] = max(widths[k], len(str(v)) if v is not None else 0)
        if fill_col is not None and r[fill_col] in FILL_GRAV:
            for c in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=c).fill = FILL_GRAV[r[fill_col]]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
    for k, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(k)].width = min(max(w + 2, 8), 70)


def niveau(depth: int) -> str:
    return NIVEAUX[depth] if depth < len(NIVEAUX) else f"Niveau {depth + 1}"


def export_excel(devices, res, anomalies: Anomalies, out: Path):
    links, types, index = res["links"], res["types"], res["index"]
    prof_max = max((d.depth for d in devices.values()), default=0) + 1
    wb = Workbook()

    # ---------------- Lignée ----------------
    lig_rows = []

    def ligne(dev: Device, rattachement, cle, iface: Iface | None, passerelle, type_):
        niveaux = dev.path + [""] * (prof_max - len(dev.path))
        lig_rows.append(niveaux + [
            " > ".join(dev.path), dev.name, rattachement, type_, str(cle[1]), cle[0],
            (iface.vlan if iface and iface.vlan is not None else ""),
            court(iface.name) if iface else "", passerelle,
        ])

    servi_par_cle = defaultdict(list)
    for dev_id, v, gw_id, svis, chemin, verifie in res["services"]:
        for s in svis:
            servi_par_cle[(s.vrf, s.network)].append((dev_id, v, gw_id, chemin, verifie))

    for cle, lst in index.items():
        t = types[cle]
        if t in ("Interco", "East-West"):   # intercos ajoutées plus bas, east-west hors graphe
            continue
        for dev_id, i in lst:
            gw = devices[dev_id].name + (f" (VIP {i.vrrp_vip})" if i.vrrp_vip else "")
            ligne(devices[dev_id], "Porteur", cle, i, gw, t)
        for dev_id, v, gw_id, chemin, verifie in servi_par_cle.get(cle, []):
            if dev_id == gw_id:
                continue
            svi = next(i for did, i in lst if did == gw_id)
            gw = devices[gw_id].name + (f" (VIP {svi.vrrp_vip})" if svi.vrrp_vip else "")
            fake = Iface(name="", vlan=v)
            ligne(devices[dev_id], "Servi" + ("" if verifie else " (non vérifié)"), cle, fake, gw, t)

    for lk in links.values():
        for cle, ip_, ic_ in lk.intercos:
            ligne(devices[lk.child], f"Interco vers {devices[lk.parent].name}", cle, ic_, "", "Interco")

    lig_rows.sort(key=lambda r: (r[prof_max], {"P": 0, "S": 1, "I": 2}.get(r[prof_max + 2][:1], 3), r[prof_max + 4]))
    write_sheet(wb.active, [niveau(k) for k in range(prof_max)] +
                ["Chemin", "Équipement", "Rattachement", "Type", "Subnet", "VRF", "VLAN", "Interface", "Passerelle"],
                lig_rows)
    wb.active.title = "Lignée"

    # ---------------- Équipements ----------------
    nb_users = defaultdict(int)
    for cle, lst in index.items():
        if types[cle].startswith("User"):
            for did, _ in lst:
                nb_users[did] += 1
    eq_rows = []
    for d in sorted(devices.values(), key=lambda d: d.path):
        m = d.mgmt[0] if d.mgmt else {}
        eq_rows.append([
            d.name, devices[d.parent].name if d.parent else "", d.depth + 1, niveau(d.depth),
            " > ".join(d.path), d.role, d.modele, d.irf, txt(m.get("ip")),
            txt(m.get("vlan")), txt(m.get("passerelle")), len(d.children), nb_users[d.id],
            d.date_txt, d.date_collecte, d.statut, d.id,
        ])
    write_sheet(wb.create_sheet("Équipements"),
                ["Équipement", "Parent", "Profondeur", "Niveau", "Chemin", "Rôle", "Modèle", "IRF",
                 "IP management", "VLAN management", "Passerelle management", "Nb enfants",
                 "Nb subnets user", "Date .txt", "Date de collecte", "Statut rapport", "Dossier"], eq_rows)

    # ---------------- Liens ----------------
    lk_rows = []
    for lk in sorted(links.values(), key=lambda l: devices[l.child].path):
        p, c = devices[lk.parent], devices[lk.child]
        inter = " ; ".join(cle_txt(k) for k, _, _ in lk.intercos)
        ipp = " ; ".join(i.ip for _, i, _ in lk.intercos)
        ipc = " ; ".join(i.ip for _, _, i in lk.intercos)
        lk_rows.append([
            p.name, c.name, lk.type, court(lk.p_port.name), ", ".join(court(m) for m in lk.p_port.members),
            court(lk.c_port.name), ", ".join(court(m) for m in lk.c_port.members),
            inter, ipp, ipc, fmt_vlans(lk.p_port.vlans), fmt_vlans(lk.c_port.vlans),
            fmt_vlans(lk.vlans_eff) + ("" if lk.vlans_verifies or lk.vlans_eff is None else " (un seul côté)"),
            f"{lk.p_port.methode} / {lk.c_port.methode}" if lk.p_port.methode or lk.c_port.methode else "",
            lk.statut,
        ])
    write_sheet(wb.create_sheet("Liens"),
                ["Parent", "Enfant", "Type", "Port parent", "Membres parent", "Port enfant", "Membres enfant",
                 "Subnet interco", "IP parent", "IP enfant", "VLANs parent", "VLANs enfant", "VLANs effectifs",
                 "Méthode ports (parent / enfant)", "Statut"], lk_rows)

    # ---------------- Subnets ----------------
    sn_rows = []
    for cle in sorted(index, key=lambda c: (c[0], c[1])):
        lst = index[cle]
        servis = sorted({devices[did].name for did, _, gw, _, _ in servi_par_cle.get(cle, []) if did != gw})
        vlan = next((i.vlan for _, i in lst if i.vlan is not None), "")
        sn_rows.append([
            str(cle[1]), cle[0], types[cle], ", ".join(devices[did].name for did, _ in lst),
            ", ".join(court(i.name) for _, i in lst), ", ".join(i.ip for _, i in lst), vlan,
            ", ".join(i.vrrp_vip for _, i in lst if i.vrrp_vip), ", ".join(servis),
        ])
    write_sheet(wb.create_sheet("Subnets"),
                ["Subnet", "VRF", "Type", "Équipement(s)", "Interface(s)", "IP", "VLAN", "VIP VRRP",
                 "Servi sur"], sn_rows)

    # ---------------- VLANs ----------------
    gw_par = {(dev_id, v): devices[gw].name for dev_id, v, gw, _, _, _ in res["services"]}
    vl_rows = [r + [gw_par.get((res["by_name"][r[0]].id, r[1]), "") if r[2] == "Accès" else ""]
               for r in res["vlan_rows"]]
    write_sheet(wb.create_sheet("VLANs"), ["Équipement", "VLAN", "Présence", "Interfaces", "Passerelle"], vl_rows)

    # ---------------- East-West ----------------
    write_sheet(wb.create_sheet("East-West"),
                ["Équipement A", "Interface A", "IP A", "Équipement B", "Interface B", "IP B", "Subnet", "VRF"],
                res["east_west"])

    # ---------------- Anomalies ----------------
    ordre = {"Erreur": 0, "Avertissement": 1, "Info": 2}
    an = sorted(anomalies.rows, key=lambda r: (ordre.get(r[0], 9), r[1], r[2]))
    write_sheet(wb.create_sheet("Anomalies"), ["Gravité", "Catégorie", "Équipement", "Objet", "Détail"], an,
                fill_col=0)

    wb.save(out)


# ===========================================================================
# 5. Export draw.io
# ===========================================================================

DEV_W, DEV_H = 170, 64
SUB_W, SUB_H, SUB_GAP = 130, 32, 6
H_GAP, V_GAP = 50, 150
LBL_W, LBL_H = 200, 14

STYLE_DEV = {
    "L3": "rounded=1;whiteSpace=wrap;html=1;fillColor=#DAE8FC;strokeColor=#6C8EBF;fontSize=11;",
    "L2": "rounded=1;whiteSpace=wrap;html=1;fillColor=#D5E8D4;strokeColor=#82B366;fontSize=11;",
    "Opaque": "rounded=1;whiteSpace=wrap;html=1;fillColor=#F5F5F5;strokeColor=#666666;dashed=1;fontSize=11;",
}
STYLE_ERR = "strokeColor=#FF0000;strokeWidth=2;"
STYLE_SUB = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF2CC;strokeColor=#D6B656;fontSize=9;"
STYLE_EDGE = {
    "L3": "endArrow=none;html=1;strokeWidth=3;strokeColor=#1F4E79;rounded=0;",
    "L2": "endArrow=none;html=1;strokeWidth=2;strokeColor=#548235;rounded=0;",
    "?": "endArrow=none;html=1;strokeWidth=1;strokeColor=#999999;dashed=1;rounded=0;",
}
STYLE_SUB_EDGE = "endArrow=none;html=1;strokeWidth=1;strokeColor=#B0A060;dashed=1;dashPattern=1 2;" \
                 "edgeStyle=orthogonalEdgeStyle;rounded=0;exitX=1;exitY=0.5;entryX=0;entryY=0.5;"
STYLE_EW = "endArrow=none;html=1;strokeWidth=2;strokeColor=#C00000;dashed=1;rounded=0;" \
           "exitX=0.75;exitY=0;entryX=0.75;entryY=0;"
STYLE_LBL = "text;html=1;align=left;verticalAlign=middle;fontSize=9;spacing=0;spacingLeft=2;" \
            "labelBackgroundColor=#FFFFFF;"


class Page:
    def __init__(self, nom: str):
        self.nom = nom
        self.cells: list[str] = ['<mxCell id="0"/>']
        self.n = 0
        self.layers: dict[str, str] = {}

    def _id(self):
        self.n += 1
        return f"c{self.n}"

    def layer(self, nom):
        if nom not in self.layers:
            lid = "1" if not self.layers else self._id()
            self.layers[nom] = lid
            self.cells.append(f'<mxCell id="{lid}" value={quoteattr(nom)} parent="0"/>')
        return self.layers[nom]

    def vertex(self, layer, value, style, x, y, w, h):
        cid = self._id()
        self.cells.append(
            f'<mxCell id="{cid}" value={quoteattr(value)} style={quoteattr(style)} vertex="1" '
            f'parent="{self.layer(layer)}"><mxGeometry x="{x:.0f}" y="{y:.0f}" width="{w}" height="{h}" '
            f'as="geometry"/></mxCell>')
        return cid

    def edge(self, layer, src, tgt, style, points=(), value=""):
        cid = self._id()
        pts = "".join(f'<mxPoint x="{x:.0f}" y="{y:.0f}"/>' for x, y in points)
        pts = f'<Array as="points">{pts}</Array>' if pts else ""
        self.cells.append(
            f'<mxCell id="{cid}" value={quoteattr(value)} style={quoteattr(style)} edge="1" '
            f'parent="{self.layer(layer)}" source="{src}" target="{tgt}">'
            f'<mxGeometry relative="1" as="geometry">{pts}</mxGeometry></mxCell>')
        return cid

    def label(self, layer, text, x, y, w=LBL_W):
        if text:
            self.vertex(layer, text, STYLE_LBL, x, y - LBL_H / 2, w, LBL_H)

    def xml(self):
        return (f'<diagram id={quoteattr(self.nom)} name={quoteattr(self.nom)}>'
                f'<mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" '
                f'connect="1" arrows="1" fold="1" page="0" pageScale="1" math="0" shadow="0">'
                f'<root>{"".join(self.cells)}</root></mxGraphModel></diagram>')


def subnets_affiches(dev_id, res) -> list[str]:
    out = []
    for cle, lst in res["index"].items():
        t = res["types"][cle]
        if t not in ("User", "User (VRRP)", "Loopback"):
            continue
        for did, i in lst:
            if did == dev_id:
                lab = f"<b>{escape(str(cle[1]))}</b>"
                extra = []
                if i.vlan is not None:
                    extra.append(f"VLAN {i.vlan}")
                if cle[0]:
                    extra.append(escape(cle[0]))
                if i.vrrp_vip:
                    extra.append(f"VIP {escape(i.vrrp_vip)}")
                if t == "Loopback":
                    extra.append("loopback")
                out.append(lab + ("<br>" + " · ".join(extra) if extra else ""))
    out.sort()
    if len(out) > MAX_SUBNETS_AFFICHES:
        reste = len(out) - (MAX_SUBNETS_AFFICHES - 1)
        out = out[:MAX_SUBNETS_AFFICHES - 1] + [f"… +{reste} autres (voir Excel)"]
    return out


def dessiner(page: Page, racines: list[str], devices, res, inclure, gravites):
    """Dessine les sous-arbres `racines` (filtrés par `inclure`) sur la page."""
    subs = {did: subnets_affiches(did, res) for did in devices if inclure(did)}

    def enfants(did):
        return [c for c in devices[did].children if inclure(c)]

    def own_w(did):
        return DEV_W + (SUB_W + 20 if subs[did] else 0)

    largeur: dict[str, float] = {}

    def calc_w(did):
        ch = enfants(did)
        wc = sum(calc_w(c) for c in ch) + H_GAP * max(len(ch) - 1, 0)
        largeur[did] = max(own_w(did), wc)
        return largeur[did]

    for r in racines:
        calc_w(r)

    base_depth = min(devices[r].depth for r in racines)
    row_h: dict[int, float] = defaultdict(lambda: DEV_H)
    for did in subs:
        k = devices[did].depth - base_depth
        row_h[k] = max(row_h[k], DEV_H, len(subs[did]) * (SUB_H + SUB_GAP))
    max_k = max((devices[d].depth - base_depth for d in subs), default=0)
    row_y = {0: 0.0}
    for k in range(1, max_k + 1):
        row_y[k] = row_y[k - 1] + row_h[k - 1] + V_GAP

    pos: dict[str, tuple[float, float]] = {}
    cell: dict[str, str] = {}

    def placer(did, x0):
        k = devices[did].depth - base_depth
        w = largeur[did]
        x = x0 + (w - own_w(did)) / 2
        pos[did] = (x, row_y[k])
        ch = enfants(did)
        wc = sum(largeur[c] for c in ch) + H_GAP * max(len(ch) - 1, 0)
        cx = x0 + (w - wc) / 2
        for c in ch:
            placer(c, cx)
            cx += largeur[c] + H_GAP

    x0 = 0.0
    for r in racines:
        placer(r, x0)
        x0 += largeur[r] + H_GAP * 2

    # Légende
    page.vertex("Base", "<b>Légende</b> : trait bleu épais = lien L3 · trait vert = lien L2 · "
                "pointillé rouge = east-west (hors arbre) · bordure rouge = anomalie bloquante<br>"
                "Sur chaque lien : étiquette haute = port côté parent, étiquette basse = port côté enfant",
                "text;html=1;align=left;verticalAlign=top;fontSize=11;", 0, -70, 900, 40)

    # Équipements + subnets
    for did, (x, y) in pos.items():
        d = devices[did]
        mg = txt(d.mgmt[0].get("ip")) if d.mgmt else ""
        lab = f"<b>{escape(d.name)}</b>"
        if d.modele:
            lab += f"<br><font style='font-size:9px'>{escape(d.modele)}</font>"
        if mg:
            lab += f"<br><font style='font-size:9px'>mgmt {escape(mg)}</font>"
        style = STYLE_DEV[d.role] + (STYLE_ERR if gravites.get(d.name) == "Erreur" else "")
        cell[did] = page.vertex("Base", lab, style, x, y, DEV_W, DEV_H)
        for k, s in enumerate(subs[did]):
            sid = page.vertex("Subnets", s, STYLE_SUB, x + DEV_W + 20, y + k * (SUB_H + SUB_GAP), SUB_W, SUB_H)
            page.edge("Subnets", cell[did], sid, STYLE_SUB_EDGE)

    # Liens parent -> enfant
    for did in pos:
        d = devices[did]
        if not d.parent or d.parent not in pos:
            continue
        lk = res["links"][did]
        px, py = pos[d.parent]
        cx, cy = pos[did]
        pcx, ccx = px + DEV_W / 2, cx + DEV_W / 2
        k = d.depth - base_depth - 1
        bend = row_y[k] + row_h[k] + V_GAP * 0.35
        style = STYLE_EDGE["L3" if lk.type.startswith("L3") else ("?" if lk.type == "?" else "L2")]
        style += "exitX=0.5;exitY=1;entryX=0.5;entryY=0;"
        page.edge("Base", cell[d.parent], cell[did], style, [(pcx, bend), (ccx, bend)])

        def port_txt(p: Port):
            if not p.name:
                return "port ?"
            return court(p.name) + (f" ({', '.join(court(m) for m in p.members)})" if p.members else "")

        page.label("Ports", port_txt(lk.p_port), ccx + 3, bend + 12)
        page.label("Ports", port_txt(lk.c_port), ccx + 3, cy - 12)
        y_mid = bend + 32
        if lk.intercos:
            page.label("Base", "<b>" + " ; ".join(escape(cle_txt(c)) for c, _, _ in lk.intercos) + "</b>",
                       ccx + 3, y_mid)
            y_mid += 16
        if lk.vlans_eff:
            page.label("VLANs", "VLANs " + fmt_vlans(lk.vlans_eff), ccx + 3, y_mid)

    # East-west (si les deux extrémités sont sur la page)
    vus = set()
    for a, _, _, b, _, _, net, vrf in res["east_west"]:
        da, db = res["by_name"].get(a), res["by_name"].get(b)
        if not da or not db or da.id not in pos or db.id not in pos or (a, b, net) in vus:
            continue
        vus.add((a, b, net))
        (ax, ay), (bx, by) = pos[da.id], pos[db.id]
        y = min(ay, by) - 30
        page.edge("East-West", cell[da.id], cell[db.id], STYLE_EW,
                  [(ax + DEV_W * 0.75, y), (bx + DEV_W * 0.75, y)])
        page.label("East-West", f"<font color='#C00000'>east-west {escape(net)}</font>",
                   (ax + bx) / 2 + DEV_W * 0.75, y - 9, 160)


def export_drawio(devices, res, anomalies: Anomalies, out: Path, prof_bloc: int):
    gravites = {}
    for g, _, eq, _, _ in anomalies.rows:
        if g == "Erreur":
            gravites[eq] = "Erreur"
    racines = [d.id for d in devices.values() if not d.parent]
    pages = []

    def nouvelle(nom):
        p = Page(nom)
        for l in ("Base", "Ports", "Subnets", "VLANs", "East-West"):
            p.layer(l)
        pages.append(p)
        return p

    # Synthèse : niveaux 0..prof_bloc
    dessiner(nouvelle("Synthèse"), racines, devices, res,
             lambda did: devices[did].depth <= prof_bloc, gravites)

    # Une page par bloc (équipement à la profondeur de découpe ayant des enfants)
    for d in sorted(devices.values(), key=lambda d: d.path):
        if d.depth == prof_bloc and d.children:
            def dans_bloc(did, root=d):
                cur = devices[did]
                while cur:
                    if cur.id == root.id:
                        return True
                    cur = devices[cur.parent] if cur.parent else None
                return False
            dessiner(nouvelle(f"Bloc {d.name}"), [d.id], devices, res, dans_bloc, gravites)

    if len(devices) <= SEUIL_VUE_COMPLETE:
        dessiner(nouvelle("Vue complète"), racines, devices, res, lambda did: True, gravites)

    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n'
           f'<mxfile host="topologie.py" modified="{datetime.now().isoformat(timespec="seconds")}" '
           f'type="device">' + "".join(p.xml() for p in pages) + "</mxfile>\n")
    out.write_text(xml, encoding="utf-8")
    return [p.nom for p in pages]


# ===========================================================================
# Main
# ===========================================================================


def main(argv=None):
    ap = argparse.ArgumentParser(description="Arborescence Comware -> Excel plat + draw.io")
    ap.add_argument("racine", type=Path, help="Dossier racine de l'arborescence")
    ap.add_argument("-o", "--sortie", type=Path, default=Path("."), help="Dossier de sortie (défaut : courant)")
    ap.add_argument("--profondeur-bloc", type=int, default=1,
                    help="Profondeur de découpe des pages draw.io (défaut : 1 = un bloc par distribution)")
    args = ap.parse_args(argv)

    if not args.racine.is_dir():
        sys.exit(f"Dossier introuvable : {args.racine}")
    args.sortie.mkdir(parents=True, exist_ok=True)

    anomalies = Anomalies()
    devices = parcourir(args.racine, anomalies)
    if not devices:
        sys.exit("Aucun équipement trouvé.")
    lire_rapports(devices, anomalies)
    res = analyser(devices, anomalies)

    xlsx = args.sortie / "topologie.xlsx"
    drawio = args.sortie / "topologie.drawio"
    export_excel(devices, res, anomalies, xlsx)
    pages = export_drawio(devices, res, anomalies, drawio, args.profondeur_bloc)

    nb = defaultdict(int)
    for r in anomalies.rows:
        nb[r[0]] += 1
    print(f"{len(devices)} équipements, {len(res['links'])} liens, {len(res['index'])} subnets, "
          f"{len(res['east_west'])} lien(s) east-west")
    print("Anomalies : " + ", ".join(f"{k} {v}" for k, v in nb.items()) if nb else "Aucune anomalie")
    print(f"-> {xlsx}")
    print(f"-> {drawio} (pages : {', '.join(pages)})")


if __name__ == "__main__":
    main()
