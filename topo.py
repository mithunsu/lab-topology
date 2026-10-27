#!/usr/bin/env python3
"""
topo.py - draw a lab topology diagram from an Excel workbook

Workbook sheets
  Devices : name, kind, model, role, emp_ip, tier, notes
            kind = switch (default) | ixia | server | mgmt
            tier = row in the diagram (1 = top); order inside a tier = sheet order
  Links   : device_a, port_a, patch_a, device_b, port_b, patch_b, cable, speed,
            status, date_changed, notes, loop_side
            cable  = fiber | eth | dac         (what the physical cable is)
            speed  = 1G, 10G, 25G, 40G, 100G ...  printed on the line for cable types whose
                     show_speed = yes in the Cables sheet (fiber and dac by default)
            patch_a / patch_b = patch-panel port (P1, P5 ...) at that end, printed inside the box;
                     neighbouring links with the same P# get one bracket + label
            status = active (default) | planned | removed
            loop_side = top | bottom | left | right  (optional, loopbacks only; a side that faces
                     another device is ignored - loopbacks always go on an outer side)
  The CONNECTION kind is never typed in - it is worked out from the two devices:
            switch-switch = dut-dut, switch-server = traffic (arrow to the server),
            switch-ixia = ixia, anything-mgmt = mgmt, same device on both ends = loopback
  Roles   : role, color, description          colour per connection kind (edit colours only)
  Cables  : cable, width, line, show_speed, description
            line = solid | double | dashdot

  Line rules: colour = connection kind, thickness/style = cable type,
  dotted = a switch port is not recorded, dashed = planned.

Run
  python topo.py                              # redraws every workbook in setups/
  python topo.py setups/jupiter_setup1.xlsx   # just one workbook
  python topo.py --show-planned        # also draw "planned" links (dashed)
  python topo.py --title "Jupiter AOSX - Setup 1"

Outputs next to the workbook: <name>.svg and <name>.drawio
(+ <name>.png / .pdf if the optional "cairosvg" package is installed)

Requires Python 3.8+ and openpyxl  (pip install openpyxl)
"""
import argparse
import datetime as dt
import html
import math
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("This script needs openpyxl:  pip install openpyxl")

# ---------------------------------------------------------------- settings
HGAP = 240            # horizontal gap between devices in a tier
MIN_VGAP = 170        # minimum vertical gap between tiers
SIDE_PITCH = 30       # spacing of links on a left/right edge
TB_PITCH = 28         # spacing of links on a top/bottom edge
LABEL_ZONE = 62       # room for port labels next to a device
CH_PITCH = 12         # spacing of parallel routing channels
MARGIN, TITLE_H = 50, 80
FONT = "Arial, Helvetica, sans-serif"

MIN_SIZE = {"switch": (220, 120), "ixia": (150, 72), "server": (150, 56), "mgmt": (170, 56)}
SHAPE = {  # fill, stroke, text, sub-text
    "switch": ("#fff6f0", "#d9632b", "#7a2e0e", "#8a5a44"),
    "ixia": ("#34a165", "#23784a", "#ffffff", "#e8fff1"),
    "server": ("#eef2f7", "#64748b", "#1e293b", "#475569"),
    "mgmt": ("#f3f4f6", "#9ca3af", "#374151", "#6b7280"),
}
DEFAULT_ROLES = {  # color, description  (role is derived from the two devices, never typed)
    "dut-dut": ("#1f3a5f", "Switch <-> switch"),
    "traffic": ("#2a6fbd", "Switch -> server NIC (arrow points at the server)"),
    "ixia": ("#2e9e5b", "Switch <-> Ixia port"),
    "mgmt": ("#8b95a5", "Anything <-> management network"),
    "loopback": ("#7e57c2", "Port to port on the same device"),
    "other": ("#6b7280", "Any other pair (e.g. server <-> Ixia)"),
}
DEFAULT_CABLES = {  # width, line, show_speed, description
    "eth": (2, "solid", False, "Copper Ethernet (RJ45)"),
    "fiber": (3, "solid", True, "Optical fiber"),
    "dac": (4, "double", True, "Direct-attach copper (twinax)"),
    "": (2, "solid", False, "Cable type not recorded"),
}
CABLE_ALIASES = {"ethernet": "eth", "copper": "eth", "rj45": "eth", "fibre": "fiber",
                 "optic": "fiber", "optical": "fiber", "twinax": "dac"}
DOTTED, DASHED = "2 4", "7 5"      # port not recorded / planned
YES = ("yes", "y", "true", "1")
SIDES = ("right", "left", "bottom", "top")   # auto loop side: fewest links, ties in this order
OUT = {"top": (0, -1), "bottom": (0, 1), "left": (-1, 0), "right": (1, 0)}
LOOP_D = 26           # how far a loopback sticks out of its device
BADGE = "#e8692e"     # switch name badge


# ---------------------------------------------------------------- reading
def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def read_sheet(wb, name, required=True):
    if name not in wb.sheetnames:
        if required:
            sys.exit(f"Workbook has no '{name}' sheet")
        return []
    rows = list(wb[name].iter_rows(values_only=True))
    if not rows:
        return []
    hdr = [norm(h).lower().replace(" ", "_") for h in rows[0]]
    out = []
    for i, r in enumerate(rows[1:], start=2):
        d = {hdr[j]: norm(r[j]) for j in range(min(len(hdr), len(r))) if hdr[j]}
        if not any(d.values()):
            continue
        d["_row"] = i
        out.append(d)
    return out


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    devices = read_sheet(wb, "Devices")
    links = read_sheet(wb, "Links")
    roles = {k: v[0] for k, v in DEFAULT_ROLES.items()}
    for r in read_sheet(wb, "Roles", required=False):
        if r.get("role", "").lower() in roles and r.get("color"):
            roles[r["role"].lower()] = r["color"]
    cables = {k: v[:3] for k, v in DEFAULT_CABLES.items()}
    for c in read_sheet(wb, "Cables", required=False):
        key = cable_key(c.get("cable"))
        if not key:
            continue
        try:
            width = float(c.get("width") or 2)
        except ValueError:
            width = 2
        line = (c.get("line") or "solid").lower()
        cables[key] = (width, line if line in ("solid", "double", "dashdot") else "solid",
                       (c.get("show_speed") or "").lower() in YES)
    return devices, links, {"roles": roles, "cables": cables}


def cable_key(v):
    c = norm(v).lower()
    return CABLE_ALIASES.get(c, c)


def role_of(l, names):
    if l["device_a"] == l["device_b"]:
        return "loopback"
    ks = {names[l["device_a"]]["_kind"], names[l["device_b"]]["_kind"]}
    if "mgmt" in ks:
        return "mgmt"
    if ks == {"switch"}:
        return "dut-dut"
    if ks == {"switch", "ixia"}:
        return "ixia"
    if ks == {"switch", "server"}:
        return "traffic"
    return "other"


SPEED_COLORS = [(10, None), (25, "#0f766e"), (40, "#e07b22"), (50, "#b7791f"), (100, "#c8352b"),
                (200, "#9d174d"), (400, "#0e7490"), (10 ** 9, "#4338ca")]


def speed_gbps(s):
    nums = re.findall(r"\d+(?:\.\d+)?", norm(s))
    return float(nums[-1]) if nums else None          # "2X200G" -> 200 (per lane)


def speed_class(s):
    """(label, colour) for DUT-DUT links faster than 10G, else None."""
    g = speed_gbps(s)
    if g is None:
        return None
    for lim, color in SPEED_COLORS:
        if g <= lim:
            return (f"{lim:g}G" if lim < 10 ** 9 else ">400G", color) if color else None
    return None


def style_for(l, cfg):
    """Colour from the connection kind (DUT-DUT also by speed), width/line from the cable."""
    width, line, _ = cfg["cables"].get(l["_cable"], cfg["cables"][""])
    dash = DASHED if l.get("_planned") else DOTTED if l.get("_unk_port") else ("10 4 2 4" if line == "dashdot" else "")
    color = cfg["roles"].get(l["_role"], DEFAULT_ROLES["other"][0])
    sc = speed_class(l.get("speed")) if l["_role"] == "dut-dut" else None
    if sc:
        color = sc[1]
    return {"role": l["_role"], "color": color, "speed_class": sc[0] if sc else None,
            "width": width, "double": line == "double", "dash": dash, "arrow": l.get("_arrow")}


def kind_of(d):
    k = (d.get("kind") or "switch").lower()
    if k.startswith("ixia"):
        return "ixia"
    if k in ("server", "host", "nic", "pc"):
        return "server"
    if k in ("mgmt", "management", "oob"):
        return "mgmt"
    return "switch"


# ---------------------------------------------------------------- checks
def validate(devices, links, show_planned, cfg):
    errors, warnings = [], []
    names = {}
    for d in devices:
        n = d.get("name", "")
        if not n:
            errors.append(f"Devices row {d['_row']}: missing name")
        elif n in names:
            errors.append(f"Devices row {d['_row']}: duplicate name '{n}'")
        else:
            d["_kind"] = kind_of(d)
            names[n] = d

    drawn = []
    port_use = defaultdict(list)
    for l in links:
        status = (l.get("status") or "active").lower()
        if status == "removed" or (status == "planned" and not show_planned):
            continue
        if status not in ("active", "planned"):
            warnings.append(f"Links row {l['_row']}: unknown status '{status}', drawing it as active")
        a, b = l.get("device_a", ""), l.get("device_b", "")
        bad = [x for x in (a, b) if x not in names]
        if bad:
            errors.append(f"Links row {l['_row']}: unknown device {', '.join(repr(x) for x in bad)}")
            continue
        l["_loop"] = a == b
        if l["_loop"]:
            pa, pb = l.get("port_a"), l.get("port_b")
            if pa and pa == pb:
                errors.append(f"Links row {l['_row']}: loopback uses the same port {pa} on both ends")
                continue
            ls = (l.get("loop_side") or "").lower()
            if ls and ls not in SIDES:
                warnings.append(f"Links row {l['_row']}: loop_side '{ls}' must be top/bottom/left/right, choosing automatically")
                l["loop_side"] = ""
        l["_cable"] = cable_key(l.get("cable") or l.get("media"))
        if l["_cable"] not in cfg["cables"]:
            warnings.append(f"Links row {l['_row']}: cable '{l['_cable']}' not in the Cables sheet, drawing a plain line")
            l["_cable"] = ""
        l["_show_speed"] = cfg["cables"][l["_cable"]][2]
        if l["_show_speed"] and not l.get("speed"):
            warnings.append(f"Links row {l['_row']}: {l['_cable']} cable without a speed")
        l["_role"] = role_of(l, names)
        l["_unk_port"] = any(names[l["device_" + e]]["_kind"] == "switch" and not l.get("port_" + e) for e in "ab")
        l["_arrow"] = None
        if l["_role"] == "traffic":
            l["_arrow"] = "b" if names[b]["_kind"] == "server" else "a"
        for dev, port in ((a, l.get("port_a")), (b, l.get("port_b"))):
            if port:
                port_use[(dev, port)].append(l["_row"])
        l["_planned"] = status == "planned"
        drawn.append(l)

    for (dev, port), rows in port_use.items():
        if len(rows) > 1:
            warnings.append(f"Port {dev} {port} is used by more than one link (Links rows {rows})")
    return names, drawn, errors, warnings


# ---------------------------------------------------------------- layout
# Look: DUTs are big boxes on a grid; Ixia ports and server NICs are drawn next to the DUT they
# belong to (Ixia = green ellipse, server NIC = blue "Eth.." text with an arrow); loopbacks are arcs.
GAP_X = 360           # minimum gap between DUTs side by side
GAP_Y = 230           # minimum gap between rows
SIDE_PITCH = 48       # spacing of links on a left/right edge
TB_PITCH = 46         # spacing of links on a top/bottom edge
LOOP_PITCH = 34       # spacing of the two legs of one loopback
SIDE_START = 56       # first link on a left/right edge, measured from the top of the box
CH_PITCH = 14         # spacing of parallel routing channels
LABEL_ZONE = 72       # room for port labels next to a box
SW_MIN = (300, 250)   # minimum DUT box size
STUB_T = 160          # traffic line length (DUT edge -> arrow tip)
STUB_I = 110          # DUT edge -> Ixia ellipse
IXIA_W, IXIA_H = 140, 86
PFS, CW, PH = 12, 7.0, 20     # pill font size, char width, height
TEXT_BLUE = "#2a6fbd"
CORE_KINDS = ("switch", "mgmt")


def pw(text):
    return CW * len(text) + 16


def tier_of(d):
    try:
        return int(float(d.get("tier") or 1))
    except ValueError:
        return 1


def assign_tiers(names, links):
    """Rows of devices. A device with a blank tier is placed automatically:
    next to its neighbour when all its neighbours share one tier, otherwise on a row between them."""
    tier = {}
    for d in names.values():
        try:
            tier[d["name"]] = float(d["tier"]) if norm(d.get("tier")) else None
        except ValueError:
            tier[d["name"]] = None
    nb = defaultdict(list)
    for l in links:
        a, b = l["device_a"], l["device_b"]
        if a != b and a in names and b in names:
            nb[a].append(b)
            nb[b].append(a)
    tiers = defaultdict(list)
    pending = []
    for d in names.values():
        (pending.append(d) if tier[d["name"]] is None else tiers[tier[d["name"]]].append(d))
    if not tiers:
        return [list(names.values())]
    for _ in range(len(pending) + 1):
        left = []
        for d in pending:
            ts = sorted({tier[n] for n in nb[d["name"]] if tier.get(n) is not None})
            if not ts:
                left.append(d)
                continue
            if len(ts) == 1:
                t = ts[0]
                row = tiers[t]
                idx = [i for i, x in enumerate(row) if x["name"] in nb[d["name"]]]
                row.insert(idx[0] if idx[0] < len(row) / 2 else idx[-1] + 1, d)
            else:
                t = (ts[0] + ts[-1]) / 2
                tiers[t].append(d)
            tier[d["name"]] = t
        if len(left) == len(pending):
            break
        pending = left
    if pending:
        tiers[max(tiers) + 1].extend(pending)
    return [tiers[t] for t in sorted(tiers)]


def plan(names, links, cfg):
    """Work out where everything goes. Returns a scene dict used by svg() and drawio()."""
    # ---- 1. which devices are boxes (core) and which are drawn next to a DUT (satellites)
    kinds = {n: d["_kind"] for n, d in names.items()}
    core = {n for n, k in kinds.items() if k in CORE_KINDS}
    for l in links:                                   # Ixia/server linked to another Ixia/server: box
        a, b = l["device_a"], l["device_b"]
        if a != b and a not in core and b not in core:
            core |= {a, b}
    for n in names:                                   # unconnected devices still get a box
        if not any(n in (l["device_a"], l["device_b"]) for l in links):
            core.add(n)
    cores = {n: names[n] for n in names if n in core}
    for l in links:
        a, b = l["device_a"], l["device_b"]
        l["_core_ends"] = [e for e in "ab" if l["device_" + e] in core]
        if a == b:
            l["_route"] = "loop"
        elif a in core and b in core:
            l["_route"] = "core"
        else:
            l["_route"] = "sat"
            l["_host"] = "a" if a in core else "b"
            l["_far"] = "b" if l["_host"] == "a" else "a"
            l["_satkind"] = "ixia" if kinds[l["device_" + l["_far"]]] == "ixia" else "traffic"
    core_links = [l for l in links if l["_route"] == "core"]

    rows = assign_tiers(cores, core_links)
    for ri, row in enumerate(rows):
        for ci, d in enumerate(row):
            d["_r"], d["_c"] = ri, ci

    def outer(d, side):
        r, c = d["_r"], d["_c"]
        return {"left": c == 0, "right": c == len(rows[r]) - 1,
                "top": r == 0, "bottom": r == len(rows) - 1}[side]

    # ---- 2. edges for DUT-DUT links
    gap_links = defaultdict(list)
    for l in core_links:
        A, B = names[l["device_a"]], names[l["device_b"]]
        if A["_r"] != B["_r"]:
            l["_sa"] = "bottom" if A["_r"] < B["_r"] else "top"
            l["_sb"] = "top" if l["_sa"] == "bottom" else "bottom"
            l["_kind"] = "vert"
            gap_links[max(A["_r"], B["_r"])].append(l)
        elif abs(A["_c"] - B["_c"]) == 1:
            l["_sa"] = "right" if B["_c"] > A["_c"] else "left"
            l["_sb"] = "left" if l["_sa"] == "right" else "right"
            l["_kind"] = "horiz"
        else:
            l["_sa"] = l["_sb"] = "top"
            l["_kind"] = "over"
            gap_links[A["_r"]].append(l)
    busy = Counter()
    for l in core_links:
        busy[(l["device_a"], l["_sa"])] += 1
        busy[(l["device_b"], l["_sb"])] += 1

    # ---- 3. sides for Ixia / server (satellites) then loopbacks: outer sides first
    sats_on, loops_on = Counter(), Counter()
    notes = []
    for l in links:
        if l["_route"] != "sat":
            continue
        host = l["device_" + l["_host"]]
        pref = {"top": 0, "bottom": 0, "left": 2, "right": 2} if l["_satkind"] == "ixia" else \
               {"left": 0, "right": 0, "bottom": 3, "top": 3}
        order = ("bottom", "top", "left", "right") if l["_satkind"] == "ixia" else ("left", "right", "bottom", "top")
        if not any(outer(names[host], s) for s in order):
            notes.append(f"{host} has no outer side free; {l['device_' + l['_far']]} drawn on its least busy side")
        side = min(order, key=lambda s: ((0 if outer(names[host], s) else 1000) + 6 * busy[(host, s)]
                                         + 0.5 * sats_on[(host, s)] + pref[s], order.index(s)))
        l["_s" + l["_host"]] = side
        sats_on[(host, side)] += 1
    for l in links:
        if l["_route"] != "loop":
            continue
        n = l["device_a"]
        forced = (l.get("loop_side") or "").lower()
        if forced and not outer(names[n], forced):        # never point a loopback at another switch
            notes.append(f"Links row {l['_row']}: loop_side '{forced}' on {n} faces another device, "
                         f"drawing the loopback on an outer side instead")
            forced = ""
        order = ("top", "bottom", "right", "left")
        if not forced and not any(outer(names[n], s) for s in order):
            notes.append(f"{n} has no outer side free; loopback drawn on its least busy side")
        side = forced or min(order, key=lambda s: ((0 if outer(names[n], s) else 1000) + 6 * busy[(n, s)]
                                                   + 20 * (sats_on[(n, s)] > 0) + 3 * loops_on[(n, s)]
                                                   + (0 if s in ("top", "bottom") else 2), order.index(s)))
        l["_sa"] = l["_sb"] = side
        loops_on[(n, side)] += 1

    # ---- 4. items on every edge, in drawing order, with their spacing
    edge = defaultdict(list)                           # (device, side) -> [(sortkey, item)]
    for li, l in enumerate(links):
        if l["_route"] == "core":
            for e, o in (("a", "b"), ("b", "a")):
                other = names[l["device_" + o]]
                key = (0, li) if l["_kind"] == "horiz" else (0, other["_c"] * 1000 + other["_r"], li)
                edge[(l["device_" + e], l["_s" + e])].append((key, ("end", li, e)))
        elif l["_route"] == "sat":
            group = 1 if l["_satkind"] == "traffic" else 2
            far = l["device_" + l["_far"]] if l["_satkind"] == "ixia" else ""
            edge[(l["device_" + l["_host"]], l["_s" + l["_host"]])].append(((group, far, li), ("end", li, l["_host"])))
        else:
            edge[(l["device_a"], l["_sa"])].append(((3, li, 0), ("end", li, "a")))
            edge[(l["device_a"], l["_sa"])].append(((3, li, 1), ("end", li, "b")))
    spans = {}
    for k, items in edge.items():
        items.sort(key=lambda t: t[0])
        pitch = SIDE_PITCH if k[1] in ("left", "right") else TB_PITCH
        offs, cur, prev = [], 0, None
        for _, it in items:
            li = it[1]
            if prev is not None:
                cur += LOOP_PITCH if (prev[1] == li and links[li]["_route"] == "loop") else pitch
                pl, cl = links[prev[1]], links[li]
                if (pl["_route"] == cl["_route"] == "sat" and pl["_satkind"] == cl["_satkind"] == "ixia"
                        and pl["device_" + pl["_far"]] != cl["device_" + cl["_far"]]):
                    cur += (IXIA_W if k[1] in ("top", "bottom") else IXIA_H) + 20 - pitch
            offs.append(cur)
            prev = it
        edge[k] = [(it, o) for (_, it), o in zip(items, offs)]
        spans[k] = cur

    # ---- 5. box sizes (all DUTs the same width, same height within a row)
    patched = {(l["device_" + e], l["_s" + e]) for l in links for e in l["_core_ends"] if l.get("patch_" + e)}
    for d in cores.values():
        n = d["name"]
        if d["_kind"] == "switch":
            mw, mh = SW_MIN
        else:
            mw, mh = MIN_SIZE.get(d["_kind"], (170, 60))
        d["w"] = max(mw, max(spans.get((n, "top"), 0), spans.get((n, "bottom"), 0)) + 110)
        d["_tpad"] = 34 if (n, "top") in patched else 0
        d["_bpad"] = 34 if (n, "bottom") in patched else 0
        lr = max(spans.get((n, "left"), -1), spans.get((n, "right"), -1))
        need = (SIDE_START + d["_tpad"] + lr + 60 + d["_bpad"]) if lr >= 0 else 0
        d["h"] = max(mh, need, 130 + d["_tpad"] + d["_bpad"])
    sw = [d for d in cores.values() if d["_kind"] == "switch"]
    if sw:
        W = max(d["w"] for d in sw)
        for d in sw:
            d["w"] = W
    for row in rows:
        hs = [d["h"] for d in row if d["_kind"] == "switch"]
        for d in row:
            if d["_kind"] == "switch":
                d["h"] = max(hs)

    # ---- 6. how far things stick out of each edge (labels, Ixia, arrows, loops)
    loop_leg = max([56] + [pw(l.get("port_" + e) or "") + 22 for l in links if l["_route"] == "loop" for e in "ab"])
    ix = [l for l in links if l["_route"] == "sat" and l["_satkind"] == "ixia"]
    ix_groups = Counter((l["device_" + l["_host"]], l["_s" + l["_host"]]) for l in ix)
    far_pills = any(ix_groups[(l["device_" + l["_host"]], l["_s" + l["_host"]])] > 1 and l.get("port_" + l["_far"]) for l in ix)
    stub_i = STUB_I + (60 if any(show_speed(l) for l in ix) else 0) + (60 if far_pills else 0)
    stub_t = STUB_T + (40 if any(l["_route"] == "sat" and l["_satkind"] == "traffic" and show_speed(l) for l in links) else 0)
    ext = defaultdict(float)
    for l in links:
        if l["_route"] == "sat":
            k = (l["device_" + l["_host"]], l["_s" + l["_host"]])
            vertical = k[1] in ("top", "bottom")
            if l["_satkind"] == "traffic":
                t = endpoint_text(l)
                e = (stub_t * 0.8 + 30) if vertical else stub_t + CW * 1.3 * len(t) + 30
            else:
                e = stub_i + (IXIA_H if vertical else IXIA_W) + 10
            ext[k] = max(ext[k], e)
        elif l["_route"] == "loop":
            k = (l["device_a"], l["_sa"])
            sp = pw(str(l["speed"])) if show_speed(l) else 0
            e = loop_leg + LOOP_PITCH / 2 + (PH + 14 if k[1] in ("top", "bottom") else sp + 14)
            ext[k] = max(ext[k], e)

    # ---- 7. positions
    ncols = max(len(r) for r in rows)
    gaps_c = [GAP_X] * max(ncols - 1, 0)
    for row in rows:
        if len(row) == ncols:
            for c in range(ncols - 1):
                gaps_c[c] = max(gaps_c[c], ext[(row[c]["name"], "right")] + ext[(row[c + 1]["name"], "left")] + 60)
    widest = [max((rows[r][c]["w"] for r in range(len(rows)) if len(rows[r]) == ncols), default=0) for c in range(ncols)]
    grid_w = sum(widest) + sum(gaps_c)
    extl = max(ext[(row[0]["name"], "left")] for row in rows)
    extr = max(ext[(row[-1]["name"], "right")] for row in rows)
    width = max(MARGIN * 2 + extl + grid_w + extr, 900)
    x_grid = MARGIN + extl + (width - 2 * MARGIN - extl - grid_w - extr) / 2
    for row in rows:
        if len(row) == ncols:
            x = x_grid
            for c, d in enumerate(row):
                d["x"] = x + (widest[c] - d["w"]) / 2
                x += widest[c] + (gaps_c[c] if c < ncols - 1 else 0)
        else:
            gaps = [max(GAP_X, ext[(a["name"], "right")] + ext[(b["name"], "left")] + 60) for a, b in zip(row, row[1:])]
            span = sum(d["w"] for d in row) + sum(gaps)
            x = x_grid + (grid_w - span) / 2
            for i, d in enumerate(row):
                d["x"] = x
                x += d["w"] + (gaps[i] if i < len(gaps) else 0)
    y = MARGIN + TITLE_H + max([ext[(d["name"], "top")] for d in rows[0]] + [0]) + 10
    row_top, row_bot = [], []
    for r, row in enumerate(rows):
        if r:
            need = max(ext[(d["name"], "bottom")] for d in rows[r - 1]) + max(ext[(d["name"], "top")] for d in row)
            y += max(GAP_Y, need + 50, 2 * LABEL_ZONE + len(gap_links[r]) * CH_PITCH + 40)
        for d in row:
            d["y"] = y
        row_top.append(y)
        row_bot.append(y + max(d["h"] for d in row))
        y = row_bot[-1]
    height = y + max(ext[(d["name"], "bottom")] for d in rows[-1]) + MARGIN

    # ---- 8. slot coordinates
    for (dev, side), items in edge.items():
        d = names[dev]
        span = spans[(dev, side)]
        for (_, li, e), off in items:
            if side in ("top", "bottom"):
                px = d["x"] + d["w"] / 2 - span / 2 + off
                py = d["y"] if side == "top" else d["y"] + d["h"]
            else:
                px = d["x"] if side == "left" else d["x"] + d["w"]
                if d["_kind"] == "switch":
                    py = d["y"] + SIDE_START + d.get("_tpad", 0) + off
                else:
                    py = d["y"] + d["h"] / 2 - span / 2 + off
            links[li]["_p" + e] = (px, py)

    # ---- 9. routes
    for r, gl in gap_links.items():                    # channels in the gap above row r
        lo, hi = (row_bot[r - 1], row_top[r]) if r else (MARGIN + TITLE_H, row_top[0])
        over = sorted([l for l in gl if l["_kind"] == "over"], key=lambda l: -abs(l["_pa"][0] - l["_pb"][0]))
        right, left = [], []
        for l in gl:
            if l["_kind"] != "vert":
                continue
            up, low = (l["_pa"], l["_pb"]) if l["_pa"][1] < l["_pb"][1] else (l["_pb"], l["_pa"])
            (right if low[0] >= up[0] else left).append((up[0], l))
        order = over + [l for _, l in sorted(right, key=lambda t: -t[0])] + [l for _, l in sorted(left, key=lambda t: t[0])]
        mid = (lo + hi) / 2 if r else hi - LABEL_ZONE - 20
        for k, l in enumerate(order):
            l["_cy"] = mid + (k - (len(order) - 1) / 2) * CH_PITCH
    sats = []
    groups = defaultdict(list)
    jog = Counter()
    for l in links:
        if l["_route"] == "core":
            (x1, y1), (x2, y2) = l["_pa"], l["_pb"]
            if l["_kind"] == "horiz":
                if abs(y1 - y2) < 0.5:
                    pts = [(x1, y1), (x2, y2)]
                else:
                    key = tuple(sorted((l["device_a"], l["device_b"])))
                    mx = (x1 + x2) / 2 + (jog[key] % 5 - 2) * 12
                    jog[key] += 1
                    pts = [(x1, y1), (mx, y1), (mx, y2), (x2, y2)]
            elif abs(x1 - x2) < 0.5:
                pts = [(x1, y1), (x2, y2)]
            else:
                pts = [(x1, y1), (x1, l["_cy"]), (x2, l["_cy"]), (x2, y2)]
            l["_pts"] = pts
        elif l["_route"] == "loop":
            ox, oy = OUT[l["_sa"]]
            a, b = l["_pa"], l["_pb"]
            la = (a[0] + ox * loop_leg, a[1] + oy * loop_leg)
            lb = (b[0] + ox * loop_leg, b[1] + oy * loop_leg)
            l["_pts"] = [a, la, lb, b]
            r_ = abs((lb[0] - la[0]) + (lb[1] - la[1])) / 2
            l["_arc"] = (la, lb, r_)
        else:
            far = l["device_" + l["_far"]] if l["_satkind"] == "ixia" else ""
            groups[(l["device_" + l["_host"]], l["_s" + l["_host"]], l["_satkind"], far)].append(l)
    for (host, side, kind, _far), gl in groups.items():
        ox, oy = OUT[side]
        if kind == "traffic":
            for l in gl:
                p = l["_p" + l["_host"]]
                end = (p[0] + ox * stub_t, p[1] + oy * stub_t)
                l["_pts"] = [p, end] if l["_host"] == "a" else [end, p]
                sats.append({"kind": "text", "x": end[0], "y": end[1], "side": side, "link": l,
                             "text": endpoint_text(l), "patch": l.get("patch_" + l["_far"], ""),
                             "id": f"sat{len(sats)}"})
                l["_sat"] = sats[-1]
        else:
            ps = [l["_p" + l["_host"]] for l in gl]
            if oy:
                cx = sum(p[0] for p in ps) / len(ps)
                w = max(IXIA_W, (max(p[0] for p in ps) - min(p[0] for p in ps)) + 90)
                h = IXIA_H
                edge_y = ps[0][1]
                cy = edge_y + oy * (stub_i + h / 2)
            else:
                cy = sum(p[1] for p in ps) / len(ps)
                h = max(IXIA_H, (max(p[1] for p in ps) - min(p[1] for p in ps)) + 60)
                w = IXIA_W
                cx = ps[0][0] + ox * (stub_i + w / 2)
            single = len(gl) == 1
            far = gl[0]["device_" + gl[0]["_far"]]
            sub = (gl[0].get("port_" + gl[0]["_far"]) or far.upper().replace("IXIA", "").strip()) if single else ""
            sat = {"kind": "ixia", "x": cx - w / 2, "y": cy - h / 2, "w": w, "h": h, "side": side,
                   "title": "IXIA", "sub": sub, "ip": names[far].get("emp_ip", ""),
                   "id": f"sat{len(sats)}", "links": gl, "single": single}
            sats.append(sat)
            a, b = w / 2, h / 2
            for l in gl:
                p = l["_p" + l["_host"]]
                if oy:
                    t = min(abs(p[0] - cx) / a, 0.92)
                    end = (p[0], cy - oy * b * (1 - t * t) ** 0.5)
                else:
                    t = min(abs(p[1] - cy) / b, 0.92)
                    end = (cx - ox * a * (1 - t * t) ** 0.5, p[1])
                l["_pts"] = [p, end] if l["_host"] == "a" else [end, p]
                l["_sat"] = sat
    place_labels(names, links, sats)
    ixia_ips = sorted({names[n].get("emp_ip") for n in names if kinds[n] == "ixia" and names[n].get("emp_ip")})
    return {"names": names, "cores": cores, "links": links, "sats": sats, "width": width, "height": height,
            "notes": notes,
            "ixia_ips": ixia_ips, "servers": sorted({n for n in names if kinds[n] == "server" and n not in core})}


def endpoint_text(l):
    far = l["_far"]
    return l.get("port_" + far) or l["device_" + far]


# ---------------------------------------------------------------- drawing helpers
def esc(s):
    return html.escape(str(s), quote=True)


def unit(p, q):
    dx, dy = q[0] - p[0], q[1] - p[1]
    ln = max((dx * dx + dy * dy) ** 0.5, 1e-6)
    return dx / ln, dy / ln


def pill(x, y, text, stroke, fill="#ffffff", color="#1f3a5f", vertical=False, bold=False):
    w = pw(text)
    rot = f' transform="rotate(90 {x:.1f} {y:.1f})"' if vertical else ""
    fw = ' font-weight="bold"' if bold else ' font-weight="600"'
    return (f'<g{rot}><rect x="{x - w / 2:.1f}" y="{y - PH / 2:.1f}" width="{w:.1f}" height="{PH}" rx="{PH / 2}" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="1.4"/>'
            f'<text x="{x:.1f}" y="{y + 4.2:.1f}" font-size="{PFS}"{fw} text-anchor="middle" fill="{color}">{esc(text)}</text></g>')


def show_speed(l):
    return l.get("_show_speed") and l.get("speed")


def pill_rect(x, y, text, vertical):
    w, h = pw(text), PH
    if vertical:
        w, h = h, w
    return (x - w / 2, y - h / 2, x + w / 2, y + h / 2)


def overlaps(a, b, m=3):
    return not (a[2] + m <= b[0] or b[2] + m <= a[0] or a[3] + m <= b[1] or b[3] + m <= a[1])


def point_at(pts, frac):
    segs = list(zip(pts, pts[1:]))
    lens = [abs(a[0] - b[0]) + abs(a[1] - b[1]) for a, b in segs]
    target = sum(lens) * frac
    for (a, b), ln in zip(segs, lens):
        if target <= ln:
            t = target / ln if ln else 0
            return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t), abs(a[0] - b[0]) < abs(a[1] - b[1])
        target -= ln
    a, b = segs[-1]
    return b, abs(a[0] - b[0]) < abs(a[1] - b[1])


def place_labels(names, links, sats):
    """Port pills next to each DUT, Ixia port pills next to the ellipse, speed pills where they fit."""
    taken = [(d["x"], d["y"], d["x"] + d["w"], d["y"] + d["h"]) for d in names.values() if "x" in d]
    taken += [(s["x"], s["y"], s["x"] + s["w"], s["y"] + s["h"]) for s in sats if s["kind"] == "ixia"]
    for l in links:
        l["_labels"] = []
        pts = l["_pts"]
        ends = [("a", pts[0], pts[1]), ("b", pts[-1], pts[-2])]
        for e, p, q in ends:
            port = l.get("port_" + e)
            if not port:
                continue
            if l["_route"] == "sat" and e == l["_far"]:
                sat = l["_sat"]
                if sat["kind"] != "ixia" or sat["single"]:
                    continue                           # shown as the endpoint text / inside the ellipse
            ux, uy = unit(p, q)
            vertical = abs(uy) > 0.5
            dist = pw(port) / 2 + (10 if l["_route"] == "loop" else 12)
            pos = (p[0] + ux * dist, p[1] + uy * dist)
            l["_labels"].append((pos, port, vertical, "port"))
            taken.append(pill_rect(*pos, port, vertical))
    for l in links:
        if not show_speed(l):
            continue
        txt, pts = str(l["speed"]), l["_pts"]
        if l["_route"] == "loop":
            la, lb, r_ = l["_arc"]
            ox, oy = OUT[l["_sa"]]
            apex = ((la[0] + lb[0]) / 2 + ox * r_, (la[1] + lb[1]) / 2 + oy * r_)
            off = PH / 2 + 6 if oy else pw(txt) / 2 + 6
            cands = [((apex[0] + ox * off, apex[1] + oy * off), False)]
        elif l["_route"] == "sat":
            cands = [point_at(pts, f) for f in (.62, .7, .55, .78, .48)]
        else:
            cands = [point_at(pts, f) for f in (.5, .42, .58, .34, .66, .26, .74, .46, .54, .38, .62, .3, .7)]
        for pos, vertical in cands:
            r = pill_rect(*pos, txt, vertical)
            if not any(overlaps(r, t) for t in taken):
                break
        else:
            pos, vertical = cands[0]
            r = pill_rect(*pos, txt, vertical)
        taken.append(r)
        l["_labels"].append((pos, txt, vertical, "speed"))


def label_points(l):
    return l.get("_labels", [])


PATCH_COLOR = "#8a3a12"


def patch_marks(links):
    """Patch-panel labels just inside each DUT box. ("text", (x, y, text, anchor)) / ("bar", (x1, y1, x2, y2))."""
    per_edge = defaultdict(list)
    for l in links:
        for e in l.get("_core_ends", "ab"):
            per_edge[(l["device_" + e], l["_s" + e])].append((l["_p" + e], l.get("patch_" + e, "")))
    marks = []
    for (_, side), items in per_edge.items():
        vertical_edge = side in ("left", "right")
        items.sort(key=lambda t: t[0][1] if vertical_edge else t[0][0])
        runs = []
        for p, patch in items:
            if runs and patch and runs[-1][1] == patch:
                runs[-1][0].append(p)
            else:
                runs.append(([p], patch))
        for pts, patch in runs:
            if not patch:
                continue
            (x1, y1), (x2, y2) = pts[0], pts[-1]
            if vertical_edge:
                inward = 1 if side == "left" else -1
                bx = x1 + inward * 12
                if len(pts) > 1:
                    marks.append(("bar", (bx, y1 - 6, bx, y2 + 6)))
                    marks.append(("text", (bx + inward * 8, (y1 + y2) / 2 + 5, patch, "start" if inward > 0 else "end")))
                else:
                    marks.append(("text", (x1 + inward * 12, y1 + 5, patch, "start" if inward > 0 else "end")))
            else:
                inward = 1 if side == "top" else -1
                if len(pts) > 1:
                    by = y1 + inward * 12
                    marks.append(("bar", (x1 - 6, by, x2 + 6, by)))
                    marks.append(("text", ((x1 + x2) / 2, by + inward * 16 + 5, patch, "middle")))
                else:
                    marks.append(("text", (x1, y1 + inward * 20 + 5, patch, "middle")))
    return marks


# ---------------------------------------------------------------- SVG
def svg_line(d, st, extra=""):
    dash = f' stroke-dasharray="{st["dash"]}"' if st["dash"] else ""
    mk = ""
    if st["arrow"] == "b":
        mk = f' marker-end="url(#arr{st["color"][1:]})"'
    elif st["arrow"] == "a":
        mk = f' marker-start="url(#arr{st["color"][1:]})"'
    w = st["width"] + (1 if st["double"] else 0)
    out = (f'<path d="{d}" fill="none" stroke="{st["color"]}" stroke-width="{w}" '
           f'stroke-linejoin="round"{dash}{mk}>{extra}</path>')
    if st["double"]:
        out += f'<path d="{d}" fill="none" stroke="#ffffff" stroke-width="{max(1, w - 3)}" stroke-linejoin="round"{dash}/>'
    return out


def link_path(l):
    if l["_route"] == "loop":
        a, b = l["_pts"][0], l["_pts"][-1]
        la, lb, r_ = l["_arc"]
        side = l["_sa"]
        # bulge outward: top & right -> sweep 1 when going from the lower-coordinate leg
        first_lower = (la[0] < lb[0]) if side in ("top", "bottom") else (la[1] < lb[1])
        sweep = (1 if side in ("top", "right") else 0) if first_lower else (0 if side in ("top", "right") else 1)
        return (f"M{a[0]:.1f},{a[1]:.1f} L{la[0]:.1f},{la[1]:.1f} "
                f"A{r_:.1f},{r_:.1f} 0 0 {sweep} {lb[0]:.1f},{lb[1]:.1f} L{b[0]:.1f},{b[1]:.1f}")
    return "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in l["_pts"])


def device_text(d):
    sub = " · ".join(v for v in (d.get("model"), d.get("role")) if v)
    return sub, d.get("emp_ip", "")


def badge_geom(d):
    sub, ip = device_text(d)
    tp, bp = d.get("_tpad", 0), d.get("_bpad", 0)
    bw = min(d["w"] - 40, max(160, 12 * len(d["name"]) + 60, 8 * len(ip) + 44))
    bh = 64 if ip else 44
    cy = d["y"] + tp + (d["h"] - tp - bp) / 2 - (10 if sub else 0)
    return d["x"] + (d["w"] - bw) / 2, cy - bh / 2, bw, bh


def svg_device(d):
    fill, stroke, tc, sc = SHAPE[d["_kind"]]
    x, y, w, h = d["x"], d["y"], d["w"], d["h"]
    sub, ip = device_text(d)
    tip = f'<title>{esc(d.get("notes") or d["name"])}</title>'
    if d["_kind"] == "switch":
        bx, by, bw, bh = badge_geom(d)
        out = (f'<g>{tip}<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" rx="12" fill="{fill}" '
               f'stroke="{stroke}" stroke-width="2.2" filter="url(#sh)"/>'
               f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="8" fill="{BADGE}"/>'
               f'<text x="{x + w / 2:.1f}" y="{by + (29 if ip else 29):.1f}" font-size="22" font-weight="bold" '
               f'text-anchor="middle" fill="#ffffff">{esc(d["name"])}</text>')
        if ip:
            out += (f'<text x="{x + w / 2:.1f}" y="{by + 51:.1f}" font-size="15" text-anchor="middle" '
                    f'fill="#ffffff">{esc(ip)}</text>')
        if sub:
            out += (f'<text x="{x + w / 2:.1f}" y="{by + bh + 20:.1f}" font-size="13" text-anchor="middle" '
                    f'fill="{sc}">{esc(sub)}</text>')
        return out + "</g>"
    lines = [(d["name"], 16, "bold", tc)] + [(t, 12, "normal", sc) for t in (sub, ip) if t]
    cy = y + h / 2 - (len(lines) - 1) * 9
    txt = "".join(f'<text x="{x + w / 2:.1f}" y="{cy + i * 18 + 5:.1f}" font-size="{s}" font-weight="{fw}" '
                  f'text-anchor="middle" fill="{c}">{esc(t)}</text>' for i, (t, s, fw, c) in enumerate(lines))
    if d["_kind"] == "ixia":
        shape = (f'<ellipse cx="{x + w / 2:.1f}" cy="{y + h / 2:.1f}" rx="{w / 2:.1f}" ry="{h / 2:.1f}" '
                 f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
    else:
        shape = (f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" rx="8" fill="{fill}" '
                 f'stroke="{stroke}" stroke-width="1.8" filter="url(#sh)"/>')
    return f"<g>{tip}{shape}{txt}</g>"


def svg_sat(s, cfg):
    if s["kind"] == "ixia":
        x, y, w, h = s["x"], s["y"], s["w"], s["h"]
        cx, cy = x + w / 2, y + h / 2
        t = (f'<ellipse cx="{cx:.1f}" cy="{cy:.1f}" rx="{w / 2:.1f}" ry="{h / 2:.1f}" fill="#3aa36a" '
             f'stroke="#23784a" stroke-width="2"/>')
        lines = [("IXIA", 17, "bold", "#ffffff")]
        if s["sub"]:
            lines.append((s["sub"], 14, "normal", "#ffffff"))
        if s.get("ip"):
            lines.append((s["ip"], 11, "normal", "#e3f6ea"))
        sizes = [ln[1] for ln in lines]
        top = cy - (sum(sizes) + 3 * (len(lines) - 1)) / 2
        for text, size, weight, col in lines:
            top += size
            t += (f'<text x="{cx:.1f}" y="{top - 2:.1f}" font-size="{size}" font-weight="{weight}" '
                  f'text-anchor="middle" fill="{col}">{esc(text)}</text>')
            top += 3
        return f"<g>{t}</g>"
    side, x, y = s["side"], s["x"], s["y"]
    anchor, dx, dy = {"right": ("start", 10, 5), "left": ("end", -10, 5),
                      "top": ("middle", 0, -10), "bottom": ("middle", 0, 22)}[side]
    patch = f'<tspan font-size="12" font-weight="normal" fill="#6b7280"> · {esc(s["patch"])}</tspan>' if s["patch"] else ""
    l = s["link"]
    tip = esc(f"{l['device_' + l['_far']]} {l.get('port_' + l['_far']) or ''}")
    return (f'<text x="{x + dx:.1f}" y="{y + dy:.1f}" font-size="16" font-weight="bold" text-anchor="{anchor}" '
            f'fill="{TEXT_BLUE}"><title>{tip}</title>{esc(s["text"])}{patch}</text>')


def legend_rows(scene, cfg, show_planned):
    links = scene["links"]
    styles = [style_for(l, cfg) for l in links]
    rows_ = []
    base = {"width": 2.5, "double": False, "dash": "", "arrow": None}
    dd = [(l, st) for l, st in zip(links, styles) if st["role"] == "dut-dut"]
    if any(not st["speed_class"] for _, st in dd):
        sp = {norm(l.get("speed")) for l, st in dd if not st["speed_class"] and l.get("speed")}
        rows_.append(("DUT ↔ DUT" + (f" ({sp.pop()})" if len(sp) == 1 else ""),
                      dict(base, color=cfg["roles"]["dut-dut"])))
    for lab in sorted({st["speed_class"] for _, st in dd if st["speed_class"]}, key=lambda s: speed_gbps(s) or 0):
        col = next(st["color"] for _, st in dd if st["speed_class"] == lab)
        rows_.append((f"{lab} link", dict(base, color=col)))
    names = {"traffic": "Traffic port → server NIC" + (f" ({', '.join(scene['servers'])})" if scene["servers"] else ""),
             "ixia": "Ixia connection", "loopback": "Loopback cable (same DUT)", "mgmt": "Management",
             "other": "Other connection"}
    for r in ("traffic", "ixia", "loopback", "mgmt", "other"):
        if any(st["role"] == r for st in styles):
            rows_.append((names[r], dict(base, color=cfg["roles"].get(r, "#6b7280"), arrow="b" if r == "traffic" else None)))
    cab = {l["_cable"] for l in links}
    grey = "#4b5563"
    if "dac" in cab:
        w = cfg["cables"]["dac"][0]
        rows_.append(("DAC cable (double line)", dict(base, color=grey, width=w, double=True)))
    if "fiber" in cab and len(cab - {"fiber"}) > 0:
        rows_.append(("Fiber (thicker line)", dict(base, color=grey, width=cfg["cables"]["fiber"][0])))
    if any(l.get("_unk_port") and not l.get("_planned") for l in links):
        rows_.append(("Switch port not recorded", dict(base, color=grey, width=2, dash=DOTTED)))
    if show_planned and any(l.get("_planned") for l in links):
        rows_.append(("Planned (not cabled yet)", dict(base, color=grey, width=2, dash=DASHED)))
    notes = []
    if any(l.get("patch_" + e) for l in links for e in l["_core_ends"]):
        notes += ["P# inside DUT = DUT-side patch port", "Bracketed P# = one patch group"]
    if any(l["_route"] == "sat" and l["_satkind"] == "traffic" and l.get("patch_" + l["_far"]) for l in links):
        notes.append("P# after Eth = server-side patch port")
    if scene["ixia_ips"]:
        notes.append("Ixia chassis: " + ", ".join(scene["ixia_ips"]))
    return rows_, notes


def svg(scene, cfg, title, subtitle, show_planned):
    links, width, height = scene["links"], scene["width"], scene["height"]
    rows_, notes = legend_rows(scene, cfg, show_planned)
    lw = max([CW * 1.1 * len(n) + 100 for n, _ in rows_] + [CW * len(t) + 50 for t in notes] + [160])
    lh = 46 + 29 * len(rows_) + 24 * len(notes) + 10
    H = height + lh + 20
    o = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.0f}" height="{H:.0f}" '
         f'viewBox="0 0 {width:.0f} {H:.0f}" font-family="{FONT}">',
         '<rect width="100%" height="100%" fill="#ffffff"/>']
    colors = {style_for(l, cfg)["color"] for l in links} | set(cfg["roles"].values())
    o.append("<defs>" + "".join(
        f'<marker id="arr{c[1:]}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5.5" markerHeight="5.5" '
        f'orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{c}"/></marker>' for c in colors)
        + '<filter id="sh" x="-5%" y="-5%" width="112%" height="112%"><feDropShadow dx="3" dy="4" '
          'stdDeviation="2" flood-color="#000000" flood-opacity="0.08"/></filter></defs>')
    o.append(f'<text x="{MARGIN}" y="{MARGIN + 14}" font-size="34" font-weight="800" fill="#1f2a44">{esc(title)}</text>')
    if subtitle:
        o.append(f'<text x="{MARGIN}" y="{MARGIN + 44}" font-size="17" fill="#6b7280">{esc(subtitle)}</text>')

    labels = []
    for l in links:
        st = style_for(l, cfg)
        tip = esc(f"{l.get('device_a')} {l.get('port_a') or '?'} ↔ {l.get('device_b')} {l.get('port_b') or '?'}"
                  f" · {st['role']}" + (f" · {l['_cable']}" if l["_cable"] else "")
                  + (f" {l.get('speed')}" if l.get("speed") else "")
                  + (f" · {l.get('notes')}" if l.get("notes") else ""))
        o.append(svg_line(link_path(l), st, f"<title>{tip}</title>"))
        for (x, y), text, vertical, kind in label_points(l):
            if kind == "speed":
                labels.append(pill(x, y, text, st["color"], fill=st["color"], color="#ffffff", vertical=vertical, bold=True))
            else:
                tc = st["color"] if st["role"] in ("traffic", "ixia", "loopback") or st["speed_class"] else "#1f3a5f"
                labels.append(pill(x, y, text, st["color"], color=tc, vertical=vertical))
    o.extend(labels)
    for d in scene["cores"].values():
        o.append(svg_device(d))
    for s in scene["sats"]:
        o.append(svg_sat(s, cfg))
    for kind, m in patch_marks(links):
        if kind == "bar":
            o.append(f'<line x1="{m[0]:.1f}" y1="{m[1]:.1f}" x2="{m[2]:.1f}" y2="{m[3]:.1f}" '
                     f'stroke="{PATCH_COLOR}" stroke-width="2.2"/>')
        else:
            x, y, text, anchor = m
            o.append(f'<text x="{x:.1f}" y="{y:.1f}" font-size="14" font-weight="bold" '
                     f'text-anchor="{anchor}" fill="{PATCH_COLOR}">{esc(text)}</text>')
    if rows_ or notes:
        lx, ly = MARGIN, height + 10
        o.append(f'<rect x="{lx}" y="{ly}" width="{lw:.0f}" height="{lh}" rx="10" fill="#f9fafb" stroke="#d1d5db"/>'
                 f'<text x="{lx + 20}" y="{ly + 30}" font-size="16" font-weight="bold" fill="#1f2937">Legend</text>')
        yy = ly + 58
        for name, st in rows_:
            o.append(svg_line(f"M{lx + 20},{yy - 5} L{lx + 66},{yy - 5}", st) +
                     f'<text x="{lx + 78}" y="{yy}" font-size="14" fill="#1f2937">{esc(name)}</text>')
            yy += 29
        for t in notes:
            o.append(f'<text x="{lx + 20}" y="{yy}" font-size="13" fill="#4b5563">{esc(t)}</text>')
            yy += 24
    o.append("</svg>")
    return "\n".join(o)


# ---------------------------------------------------------------- draw.io
def drawio(scene, cfg, title):
    names, links = scene["names"], scene["links"]
    cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
    ids, geo = {}, {}
    for i, d in enumerate(scene["cores"].values()):
        cid = f"dev{i}"
        ids[d["name"]] = cid
        geo[cid] = (d["x"], d["y"], d["w"], d["h"])
        fill, stroke, tc, _ = SHAPE[d["_kind"]]
        sub, ip = device_text(d)
        is_sw = d["_kind"] == "switch"
        val = "" if is_sw else f"<b>{html.escape(d['name'])}</b>" + "".join(f"<br>{html.escape(t)}" for t in (sub, ip) if t)
        shape = "ellipse;" if d["_kind"] == "ixia" else "rounded=1;arcSize=5;"
        cells.append(f'<mxCell id="{cid}" value="{esc(val)}" vertex="1" parent="1" '
                     f'style="{shape}whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};strokeWidth=2;'
                     f'shadow=1;fontColor={tc};fontSize=12;">'
                     f'<mxGeometry x="{d["x"]:.1f}" y="{d["y"]:.1f}" width="{d["w"]:.1f}" height="{d["h"]:.1f}" as="geometry"/></mxCell>')
        if is_sw:
            bx, by, bw, bh = badge_geom(d)
            bval = f"<b style='font-size:22px'>{html.escape(d['name'])}</b>" + (f"<br>{html.escape(ip)}" if ip else "")
            cells.append(f'<mxCell id="{cid}b" value="{esc(bval)}" vertex="1" connectable="0" parent="{cid}" '
                         f'style="rounded=1;arcSize=12;whiteSpace=wrap;html=1;fillColor={BADGE};strokeColor=none;'
                         f'fontColor=#ffffff;fontSize=15;">'
                         f'<mxGeometry x="{bx - d["x"]:.1f}" y="{by - d["y"]:.1f}" width="{bw:.1f}" height="{bh:.1f}" as="geometry"/></mxCell>')
            if sub:
                cells.append(f'<mxCell id="{cid}m" value="{esc(sub)}" vertex="1" connectable="0" parent="{cid}" '
                             f'style="text;html=1;align=center;fontSize=13;fontColor=#8a5a44;">'
                             f'<mxGeometry x="0" y="{by + bh + 6 - d["y"]:.1f}" width="{d["w"]:.1f}" height="18" as="geometry"/></mxCell>')
    for s in scene["sats"]:
        if s["kind"] == "ixia":
            val = ("<b>IXIA</b>" + (f"<br>{html.escape(s['sub'])}" if s["sub"] else "")
                   + (f"<br><font style='font-size:11px'>{html.escape(s['ip'])}</font>" if s.get("ip") else ""))
            cells.append(f'<mxCell id="{s["id"]}" value="{esc(val)}" vertex="1" parent="1" '
                         f'style="ellipse;whiteSpace=wrap;html=1;fillColor=#3aa36a;strokeColor=#23784a;strokeWidth=2;'
                         f'fontColor=#ffffff;fontSize=15;">'
                         f'<mxGeometry x="{s["x"]:.1f}" y="{s["y"]:.1f}" width="{s["w"]:.1f}" height="{s["h"]:.1f}" as="geometry"/></mxCell>')
            geo[s["id"]] = (s["x"], s["y"], s["w"], s["h"])
        else:
            val = f"<b>{html.escape(s['text'])}</b>" + (f" <font color='#6b7280' style='font-size:12px'>· {html.escape(s['patch'])}</font>" if s["patch"] else "")
            tw = CW * 1.3 * (len(s["text"]) + len(s["patch"]) + 3) + 10
            side = s["side"]
            if side == "right":
                gx, gy, gw, gh, al = s["x"], s["y"] - 12, tw, 24, "left"
            elif side == "left":
                gx, gy, gw, gh, al = s["x"] - tw, s["y"] - 12, tw, 24, "right"
            elif side == "top":
                gx, gy, gw, gh, al = s["x"] - tw / 2, s["y"] - 24, tw, 24, "center"
            else:
                gx, gy, gw, gh, al = s["x"] - tw / 2, s["y"], tw, 24, "center"
            cells.append(f'<mxCell id="{s["id"]}" value="{esc(val)}" vertex="1" parent="1" '
                         f'style="text;html=1;align={al};verticalAlign=middle;fontSize=16;fontColor={TEXT_BLUE};spacing=2;">'
                         f'<mxGeometry x="{gx:.1f}" y="{gy:.1f}" width="{gw:.1f}" height="{gh}" as="geometry"/></mxCell>')
            geo[s["id"]] = (gx, gy, gw, gh)
    for i, l in enumerate(links):
        sty = style_for(l, cfg)
        color = sty["color"]
        src = ids.get(l["device_a"]) if "a" in l["_core_ends"] else l["_sat"]["id"]
        tgt = ids.get(l["device_b"]) if "b" in l["_core_ends"] else l["_sat"]["id"]
        pa, pb = l["_pts"][0], l["_pts"][-1]
        ga, gb = geo[src], geo[tgt]
        dash = f"dashed=1;dashPattern={sty['dash']};" if sty["dash"] else "dashed=0;"
        st = (f"startArrow={'block' if sty['arrow'] == 'a' else 'none'};startFill=1;"
              f"endArrow={'block' if sty['arrow'] == 'b' else 'none'};endFill=1;html=1;rounded=0;edgeStyle=none;"
              f"{'shape=link;width=3;' if sty['double'] else ''}{'curved=1;' if l['_route'] == 'loop' else ''}"
              f"strokeColor={color};strokeWidth={sty['width'] if not sty['double'] else 1.5};{dash}"
              f"exitX={(pa[0] - ga[0]) / ga[2]:.3f};exitY={(pa[1] - ga[1]) / ga[3]:.3f};exitDx=0;exitDy=0;exitPerimeter=0;"
              f"entryX={(pb[0] - gb[0]) / gb[2]:.3f};entryY={(pb[1] - gb[1]) / gb[3]:.3f};entryDx=0;entryDy=0;entryPerimeter=0;")
        eid = f"link{i}"
        mids = l["_pts"][1:-1]
        if l["_route"] == "loop":
            la, lb, r_ = l["_arc"]
            ox, oy = OUT[l["_sa"]]
            apex = ((la[0] + lb[0]) / 2 + ox * r_, (la[1] + lb[1]) / 2 + oy * r_)
            mids = [la, apex, lb]
        g = '<mxGeometry relative="1" as="geometry">'
        if mids:
            g += '<Array as="points">' + "".join(f'<mxPoint x="{x:.1f}" y="{y:.1f}"/>' for x, y in mids) + "</Array>"
        g += "</mxGeometry>"
        cells.append(f'<mxCell id="{eid}" value="" edge="1" parent="1" source="{src}" target="{tgt}" style="{st}">{g}</mxCell>')
        for j, ((x, y), text, vertical, kind) in enumerate(label_points(l)):
            tw = pw(text)
            fillc, fc = (color, "#ffffff") if kind == "speed" else ("#ffffff", "#1f3a5f")
            rot = "rotation=90;" if vertical else ""
            cells.append(f'<mxCell id="{eid}t{j}" value="{esc(text)}" vertex="1" parent="1" '
                         f'style="rounded=1;arcSize=50;html=1;fontSize={PFS};fontStyle=1;fillColor={fillc};strokeColor={color};'
                         f'fontColor={fc};{rot}">'
                         f'<mxGeometry x="{x - tw / 2:.1f}" y="{y - PH / 2:.1f}" width="{tw:.1f}" height="{PH}" as="geometry"/></mxCell>')
    for j, (kind, m) in enumerate(patch_marks(links)):
        if kind == "bar":
            x1, y1, x2, y2 = m
            cells.append(f'<mxCell id="pb{j}" value="" vertex="1" parent="1" style="rounded=0;fillColor={PATCH_COLOR};strokeColor=none;">'
                         f'<mxGeometry x="{min(x1, x2) - 1:.1f}" y="{min(y1, y2) - 1:.1f}" '
                         f'width="{abs(x2 - x1) + 2:.1f}" height="{abs(y2 - y1) + 2:.1f}" as="geometry"/></mxCell>')
        else:
            x, y, text, anchor = m
            tw = 9 * len(text) + 6
            left = x if anchor == "start" else x - tw if anchor == "end" else x - tw / 2
            align = {"start": "left", "end": "right", "middle": "center"}[anchor]
            cells.append(f'<mxCell id="pt{j}" value="{esc(text)}" vertex="1" parent="1" '
                         f'style="text;html=1;fontSize=14;fontStyle=1;fontColor={PATCH_COLOR};align={align};verticalAlign=middle;spacing=0;">'
                         f'<mxGeometry x="{left:.1f}" y="{y - 14:.1f}" width="{tw:.1f}" height="18" as="geometry"/></mxCell>')
    return (f'<mxfile host="topo.py"><diagram name="{esc(title)}"><mxGraphModel grid="1" gridSize="10">'
            f'<root>{"".join(cells)}</root></mxGraphModel></diagram></mxfile>')


# ---------------------------------------------------------------- main
def render(wb_path, title=None, subtitle=None, show_planned=False, out=None):
    """Draw one workbook. Returns True on success."""
    print(f"== {wb_path}")
    devices, links, cfg = load(wb_path)
    if not devices:
        print("ERROR: no devices in the Devices sheet")
        return False
    names, drawn, errors, warnings = validate(devices, links, show_planned, cfg)
    if any(l.get("type") for l in links):
        print("NOTE: the Links 'type' column is ignored now - the connection kind is worked out from the "
              "devices. Run upgrade_workbook.py to convert an older workbook.")
    for w in warnings:
        print("WARNING:", w)
    if errors:
        for e in errors:
            print("ERROR:", e)
        print("Fix the errors above in the workbook and run again.")
        return False

    scene = plan(names, drawn, cfg)
    for n in scene["notes"]:
        print("NOTE:", n)
    title = title or wb_path.stem.replace("_", " ")
    if subtitle is None:
        n_dut = sum(1 for d in scene["cores"].values() if d["_kind"] == "switch")
        subtitle = f"{n_dut} DUTs" + (f" · Ixia chassis {', '.join(scene['ixia_ips'])}" if scene["ixia_ips"] else "")
    base = Path(out) if out else wb_path.with_suffix("")

    svg_text = svg(scene, cfg, title, subtitle, show_planned)
    base.with_suffix(".svg").write_text(svg_text, encoding="utf-8")
    base.with_suffix(".drawio").write_text(drawio(scene, cfg, title), encoding="utf-8")
    made = [base.with_suffix(".svg").name, base.with_suffix(".drawio").name]
    try:
        import cairosvg
        cairosvg.svg2png(bytestring=svg_text.encode(), write_to=str(base.with_suffix(".png")), scale=2)
        cairosvg.svg2pdf(bytestring=svg_text.encode(), write_to=str(base.with_suffix(".pdf")))
        made += [base.with_suffix(".png").name, base.with_suffix(".pdf").name]
    except ImportError:
        pass
    except Exception as e:  # cairo libs missing on some Windows installs
        print("PNG/PDF skipped:", e)
    print(f"{len(names)} devices, {len(drawn)} links -> {', '.join(made)}")
    return True


def main():
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description="Draw lab topologies from Excel links lists")
    ap.add_argument("workbooks", nargs="*", help="workbook(s) to draw (default: every .xlsx in setups/)")
    ap.add_argument("--title", default=None, help="diagram title (default: workbook name)")
    ap.add_argument("--subtitle", default=None, help="line under the title")
    ap.add_argument("--show-planned", action="store_true", help="also draw links with status 'planned'")
    ap.add_argument("--out", default=None, help="output base name (single workbook only)")
    a = ap.parse_args()

    paths = [Path(p) for p in a.workbooks] or sorted(
        p for p in (here / "setups").glob("*.xlsx") if not p.name.startswith("~$") and ".bak" not in p.name)
    if not paths:
        sys.exit("No workbooks given and none found in setups/.\n"
                 "Start from the example:  mkdir setups  then copy examples/example_topology.xlsx into setups/")
    missing = [p for p in paths if not p.exists()]
    if missing:
        sys.exit("Workbook not found: " + ", ".join(map(str, missing)))
    single = len(paths) == 1
    ok = all([render(p, a.title if single else None, a.subtitle if single else None,
                     a.show_planned, a.out if single else None) for p in paths])
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
