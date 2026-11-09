"""
Assigns free /30 subnets from a source workbook to interconnects listed in a target workbook,
highlights the taken rows in the source and keeps a log of every allocation.

IDEMPOTENT: running the script any number of times gives the same result.
  - Each allocation is keyed on the interconnect NAME: an interconnect that already has a
    subnet (in the target OR recorded in the source "Assigned to" column) keeps it, never twice.
  - Missing cells (VLAN, IPs) are filled in from the subnet already assigned.
  - If source and target disagree, the source is realigned to the target
    (or a conflict is reported) without reassigning anything.
  - If there is nothing to do, no file is modified and the log is left untouched.
  - Files are modified IN PLACE, with a timestamped backup taken before writing.

Requirements: pip install openpyxl
"""
from datetime import datetime
import getpass
import ipaddress
import os
import shutil
import sys

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ===== CONFIGURATION =====
SOURCE = "subnets_192.168.10.0_30.xlsx"
SOURCE_SHEET = "Subnets"
SOURCE_COL_VLAN = "A"          # VLAN ID column ("" if none)
SOURCE_COL_SUBNET = "B"        # available subnets
SOURCE_COL_ASSIGNED_TO = "D"   # who the subnet is assigned to
SOURCE_COL_DATE = "E"          # assignment date
SOURCE_FIRST_ROW = 2

TARGET = "intercos_wan.xlsx"
TARGET_SHEET = "Intercos"
TARGET_COL_NAME = "A"          # unique interconnect name (allocation key)
TARGET_COL_SUBNET = "F"        # /30 subnet
TARGET_COL_VLAN = "G"          # VLAN ID ("" to skip)
TARGET_COL_IP_A = "H"          # 1st usable IP (site router) ("" to skip)
TARGET_COL_IP_B = "I"          # 2nd usable IP (WAN router) ("" to skip)
TARGET_FIRST_ROW = 2

TAKEN_COLOR = "FFC7CE"         # light red = subnet taken
LOG_SHEET = "History"
BACKUP_DIR = "backups"
# =========================

TAKEN_FILL = PatternFill(start_color=TAKEN_COLOR, end_color=TAKEN_COLOR, fill_type="solid")
DATE_FORMAT = "DD/MM/YYYY HH:MM"


def normalize(subnet):
    """Canonical subnet form ('192.168.10.1/30 ' -> '192.168.10.0/30'), None if invalid."""
    try:
        return str(ipaddress.ip_network(str(subnet).strip(), strict=False))
    except ValueError:
        return None


def text(v):
    return str(v).strip() if v is not None else ""


def is_highlighted(cell):
    c = cell.fill.start_color
    return cell.fill.fill_type == "solid" and c.rgb and str(c.rgb)[-6:].upper() == TAKEN_COLOR


def open_workbook(path):
    if not os.path.exists(path):
        sys.exit(f"File not found: {path}")
    return load_workbook(path)


wb_src, wb_tgt = open_workbook(SOURCE), open_workbook(TARGET)
ws_src, ws_tgt = wb_src[SOURCE_SHEET], wb_tgt[TARGET_SHEET]
changes = 0     # number of cells actually changed
events = []     # rows to append to the log
warnings = []


def write(ws, ref, value, fmt=None):
    """Only writes if the value differs -> no needless changes."""
    global changes
    cell = ws[ref]
    if cell.value != value:
        cell.value = value
        if fmt:
            cell.number_format = fmt
        changes += 1


# ---------- 1. Source state ----------
src_by_subnet = {}         # normalized subnet -> source row
src_by_name = {}           # interconnect name -> source row
manually_reserved = set()  # highlighted rows with no name: reserved by hand, left alone
for r in range(SOURCE_FIRST_ROW, ws_src.max_row + 1):
    s = normalize(ws_src[f"{SOURCE_COL_SUBNET}{r}"].value)
    if not s:
        continue
    if s in src_by_subnet:
        warnings.append(f"Subnet {s} duplicated in source (rows {src_by_subnet[s]} and {r}).")
        continue
    src_by_subnet[s] = r
    name = text(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{r}"].value)
    if name:
        if name in src_by_name:
            warnings.append(f"'{name}' has several subnets in source (rows {src_by_name[name]} and {r}).")
        else:
            src_by_name[name] = r
    elif is_highlighted(ws_src[f"{SOURCE_COL_SUBNET}{r}"]):
        manually_reserved.add(r)

# ---------- 2. Target state ----------
target_rows = []
seen = set()
target_subnets = {}   # normalized subnet -> interconnect using it
for r in range(TARGET_FIRST_ROW, ws_tgt.max_row + 1):
    name = text(ws_tgt[f"{TARGET_COL_NAME}{r}"].value)
    if not name:
        continue
    if name in seen:
        warnings.append(f"Interconnect '{name}' duplicated in target (row {r}): skipped.")
        continue
    seen.add(name)
    target_rows.append((r, name))
    s = normalize(ws_tgt[f"{TARGET_COL_SUBNET}{r}"].value)
    if s:
        if s in target_subnets:
            warnings.append(f"Subnet {s} used by both '{target_subnets[s]}' and '{name}' in target.")
        target_subnets[s] = name

# Free subnets: not named, not manually reserved, not already used in the target
free = [r for s, r in sorted(src_by_subnet.items(), key=lambda x: x[1])
        if not text(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{r}"].value)
        and r not in manually_reserved and s not in target_subnets]

now = datetime.now()
operator = getpass.getuser()


def mark_source(src_row, name):
    write(ws_src, f"{SOURCE_COL_ASSIGNED_TO}{src_row}", name)
    if not ws_src[f"{SOURCE_COL_DATE}{src_row}"].value:   # keep the original date
        write(ws_src, f"{SOURCE_COL_DATE}{src_row}", now, DATE_FORMAT)


def fill_target(row, subnet, src_row):
    write(ws_tgt, f"{TARGET_COL_SUBNET}{row}", subnet)
    if TARGET_COL_VLAN and SOURCE_COL_VLAN and src_row:
        write(ws_tgt, f"{TARGET_COL_VLAN}{row}", ws_src[f"{SOURCE_COL_VLAN}{src_row}"].value)
    hosts = list(ipaddress.ip_network(subnet).hosts())
    if TARGET_COL_IP_A and len(hosts) >= 1:
        write(ws_tgt, f"{TARGET_COL_IP_A}{row}", str(hosts[0]))
    if TARGET_COL_IP_B and len(hosts) >= 2:
        write(ws_tgt, f"{TARGET_COL_IP_B}{row}", str(hosts[1]))


# ---------- 3. Process each interconnect ----------
new_count = 0
for row, name in target_rows:
    raw = ws_tgt[f"{TARGET_COL_SUBNET}{row}"].value
    subnet = normalize(raw) if raw else None
    if raw and not subnet:
        warnings.append(f"'{name}' (row {row}): invalid subnet '{raw}', row skipped.")
        continue

    if subnet:
        # Case A: target already has a subnet -> align the source with it
        src_row = src_by_subnet.get(subnet)
        if src_row is None:
            warnings.append(f"'{name}': subnet {subnet} not found in source (entered by hand?).")
        else:
            owner = text(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{src_row}"].value)
            if owner and owner != name:
                warnings.append(f"CONFLICT: {subnet} assigned to '{owner}' in source "
                                f"but used by '{name}' in target. Nothing changed.")
                continue
            if not owner:
                mark_source(src_row, name)
                events.append(("Realigned", subnet, name, src_row, row))
    elif name in src_by_name:
        # Case B: source already knows this interconnect's subnet -> reuse it
        src_row = src_by_name[name]
        subnet = normalize(ws_src[f"{SOURCE_COL_SUBNET}{src_row}"].value)
        if subnet in target_subnets:
            warnings.append(f"CONFLICT: {subnet} reserved for '{name}' but used by "
                            f"'{target_subnets[subnet]}' in target. Nothing changed.")
            continue
        events.append(("Restored", subnet, name, src_row, row))
    else:
        # Case C: new interconnect -> next free subnet
        if not free:
            warnings.append(f"No free subnet left for '{name}' (row {row}).")
            continue
        src_row = free.pop(0)
        subnet = normalize(ws_src[f"{SOURCE_COL_SUBNET}{src_row}"].value)
        mark_source(src_row, name)
        events.append(("Assigned", subnet, name, src_row, row))
        new_count += 1

    fill_target(row, subnet, src_row)

# ---------- 4. Highlight = exact mirror of "Assigned to" ----------
for r in src_by_subnet.values():
    if text(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{r}"].value):
        for c in ws_src[r]:
            if not is_highlighted(c):
                c.fill = TAKEN_FILL
                changes += 1

# ---------- 5. Log ----------
if events:
    if LOG_SHEET in wb_src.sheetnames:
        ws_log = wb_src[LOG_SHEET]
    else:
        ws_log = wb_src.create_sheet(LOG_SHEET)
        ws_log.append(["Date", "Action", "Subnet", "Assigned to", "Source row",
                       "Target file", "Target row", "Operator"])
        for c in ws_log[1]:
            c.font = Font(bold=True)
    for action, subnet, name, s_row, t_row in events:
        ws_log.append([now, action, subnet, name, s_row, TARGET, t_row, operator])
        ws_log.cell(row=ws_log.max_row, column=1).number_format = DATE_FORMAT

# ---------- 6. Duplicate check on the final state ----------
def check_duplicates():
    """Checks uniqueness of names, subnets, VLANs and IPs, overlaps and source/target
    consistency. Returns the list of problems found (empty = all good)."""
    problems = []

    def dupes(label, items):
        seen_at = {}
        for value, where in items:
            if value in (None, ""):
                continue
            key = str(value).strip()
            if key in seen_at:
                problems.append(f"Duplicate {label} '{key}': {seen_at[key]} and {where}.")
            else:
                seen_at[key] = where

    def overlaps(label, items):
        nets = sorted(((ipaddress.ip_network(s), w) for s, w in items),
                      key=lambda x: (x[0].network_address, x[0].prefixlen))
        for (n1, w1), (n2, w2) in zip(nets, nets[1:]):
            if n1 != n2 and n1.overlaps(n2):
                problems.append(f"Overlapping {label} {n1} ({w1}) and {n2} ({w2}).")

    # Source
    src_subnets = [(normalize(ws_src[f"{SOURCE_COL_SUBNET}{r}"].value), f"source row {r}")
                   for r in range(SOURCE_FIRST_ROW, ws_src.max_row + 1)]
    src_subnets = [x for x in src_subnets if x[0]]
    dupes("subnet in source", src_subnets)
    overlaps("subnets in source", src_subnets)
    if SOURCE_COL_VLAN:
        dupes("VLAN ID in source", [(ws_src[f"{SOURCE_COL_VLAN}{r}"].value, f"source row {r}")
                                    for r in range(SOURCE_FIRST_ROW, ws_src.max_row + 1)])
    dupes("assignee in source", [(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{r}"].value, f"source row {r}")
                                 for r in range(SOURCE_FIRST_ROW, ws_src.max_row + 1)])

    # Target
    rows = [r for r in range(TARGET_FIRST_ROW, ws_tgt.max_row + 1)
            if text(ws_tgt[f"{TARGET_COL_NAME}{r}"].value)]
    dupes("interconnect name in target", [(ws_tgt[f"{TARGET_COL_NAME}{r}"].value, f"target row {r}") for r in rows])
    tgt_subnets = [(normalize(ws_tgt[f"{TARGET_COL_SUBNET}{r}"].value), f"target row {r}") for r in rows]
    tgt_subnets = [x for x in tgt_subnets if x[0]]
    dupes("subnet in target", tgt_subnets)
    overlaps("subnets in target", tgt_subnets)
    if TARGET_COL_VLAN:
        dupes("VLAN ID in target", [(ws_tgt[f"{TARGET_COL_VLAN}{r}"].value, f"target row {r}") for r in rows])
    ips = []
    for col in (TARGET_COL_IP_A, TARGET_COL_IP_B):
        if col:
            ips += [(ws_tgt[f"{col}{r}"].value, f"target {col}{r}") for r in rows]
    dupes("IP address in target", ips)

    # Consistency: each target subnet must be assigned to the same name in the source
    for r in rows:
        s = normalize(ws_tgt[f"{TARGET_COL_SUBNET}{r}"].value)
        if not s:
            continue
        name = text(ws_tgt[f"{TARGET_COL_NAME}{r}"].value)
        src_row = next((rr for v, w in src_subnets if v == s
                        for rr in [int(w.split()[-1])]), None)
        if src_row is None:
            problems.append(f"Target row {r}: subnet {s} not found in source.")
        else:
            owner = text(ws_src[f"{SOURCE_COL_ASSIGNED_TO}{src_row}"].value)
            if owner != name:
                problems.append(f"Target row {r}: {s} belongs to '{owner or '(nobody)'}' "
                                f"in source, not '{name}'.")
    return problems


# ---------- 7. Save (only if something changed) ----------
for w in warnings:
    print("WARNING:", w)

problems = check_duplicates()
if problems:
    print(f"DUPLICATE CHECK: {len(problems)} problem(s) found:")
    for p in problems:
        print("  -", p)
else:
    print("DUPLICATE CHECK: OK (no duplicate name, subnet, VLAN or IP; no overlap).")

if changes == 0:
    print("Nothing to do: everything is up to date. No file modified.")
    sys.exit(1 if problems else 0)

stamp = now.strftime("%Y%m%d_%H%M%S")
os.makedirs(BACKUP_DIR, exist_ok=True)
for path in (SOURCE, TARGET):
    base, ext = os.path.splitext(os.path.basename(path))
    shutil.copy2(path, os.path.join(BACKUP_DIR, f"{base}_{stamp}{ext}"))


def save(wb, path):
    """Write to a temp file then replace: never a half-written file."""
    tmp = path + ".tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except PermissionError:
        if os.path.exists(tmp):
            os.remove(tmp)
        sys.exit(f"Cannot write {path}: close it in Excel and run again. "
                 f"Re-running is safe (idempotent script).")


# Target first: if the source save then fails, the next run realigns the source
save(wb_tgt, TARGET)
save(wb_src, SOURCE)
print(f"{new_count} new assignment(s), {len(events) - new_count} realigned/restored, "
      f"{len(free)} free subnet(s) left. Backups in '{BACKUP_DIR}'.")
sys.exit(1 if problems else 0)