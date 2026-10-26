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
            loop_side = top | bottom | left | right  (optional, loopbacks only)
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

MIN_SIZE = {"switch": (210, 96), "ixia": (130, 56), "server": (150, 56), "mgmt": (170, 56)}
SHAPE = {  # fill, stroke, text, sub-text
    "switch": ("#fff6ef", "#d4622a", "#7a2e0e", "#5b4636"),
    "ixia": ("#34a165", "#23784a", "#ffffff", "#e8fff1"),
    "server": ("#eef2f7", "#64748b", "#1e293b", "#475569"),
    "mgmt": ("#f3f4f6", "#9ca3af", "#374151", "#6b7280"),
}
DEFAULT_ROLES = {  # color, description  (role is derived from the two devices, never typed)
    "dut-dut": ("#1f3a5f", "Switch <-> switch"),
    "traffic": ("#2a6fbd", "Switch -> server NIC (arrow points at the server)"),
    "ixia": ("#2e9e5b", "Switch <-> Ixia port"),
    "mgmt": ("#8b95a5", "Anything <-> management network"),
    "loopback": ("#7c3aed", "Port to port on the same device"),
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
LOOP_D = 72           # how far a loopback sticks out of its device


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


def style_for(l, cfg):
    """Everything needed to draw one link: colour from the connection kind, width/line from the cable."""
    width, line, _ = cfg["cables"].get(l["_cable"], cfg["cables"][""])
    dash = DASHED if l.get("_planned") else DOTTED if l.get("_unk_port") else ("10 4 2 4" if line == "dashdot" else "")
    return {"role": l["_role"], "color": cfg["roles"].get(l["_role"], DEFAULT_ROLES["other"][0]),
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
def tier_of(d):
    try:
        return int(float(d.get("tier") or 1))
    except ValueError:
        return 1


def plan(names, links):
    tiers = defaultdict(list)
    for d in names.values():
        tiers[tier_of(d)].append(d)
    rows = [tiers[t] for t in sorted(tiers)]
    for ri, row in enumerate(rows):
        for ci, d in enumerate(row):
            d["_r"], d["_c"] = ri, ci

    # 1. which edge each link end uses
    opp = {"left": "right", "right": "left"}
    gap_links = defaultdict(list)          # gap above row r -> links routed through it
    for l in links:
        if l.get("_loop"):
            continue
        A, B = names[l["device_a"]], names[l["device_b"]]
        if A["_r"] != B["_r"]:
            l["_sa"] = "bottom" if A["_r"] < B["_r"] else "top"
            l["_sb"] = "top" if l["_sa"] == "bottom" else "bottom"
            l["_route"] = "vert"
            gap_links[max(A["_r"], B["_r"])].append(l)
        elif abs(A["_c"] - B["_c"]) == 1:
            l["_sa"] = "right" if B["_c"] > A["_c"] else "left"
            l["_sb"] = opp[l["_sa"]]
            l["_route"] = "horiz"
        else:
            l["_sa"] = l["_sb"] = "top"
            l["_route"] = "over"
            gap_links[A["_r"]].append(l)

    # 2. device sizes from the number of links on each edge
    cnt = Counter()
    for l in links:
        if not l.get("_loop"):
            cnt[(l["device_a"], l["_sa"])] += 1
            cnt[(l["device_b"], l["_sb"])] += 1
    loop_top, loop_bot = set(), set()        # rows that have loops sticking up / down
    for l in links:                          # loopbacks: chosen side, or the least busy one
        if not l.get("_loop"):
            continue
        n = l["device_a"]
        side = (l.get("loop_side") or "").lower() or min(SIDES, key=lambda s: (cnt[(n, s)], SIDES.index(s)))
        l["_sa"] = l["_sb"] = side
        l["_route"] = "loop"
        cnt[(n, side)] += 2
        if side == "top":
            loop_top.add(names[n]["_r"])
        elif side == "bottom":
            loop_bot.add(names[n]["_r"])
    patched = {(l["device_" + e], l["_s" + e]) for l in links for e in "ab" if l.get("patch_" + e)}
    for d in names.values():
        n = d["name"]
        mw, mh = MIN_SIZE[d["_kind"]]
        ntb = max(cnt[(n, "top")], cnt[(n, "bottom")])
        nlr = max(cnt[(n, "left")], cnt[(n, "right")])
        d["w"] = max(mw, ntb * TB_PITCH + 40)
        d["h"] = max(mh, nlr * SIDE_PITCH + 36)
        # extra room for patch-panel labels on the top/bottom edge
        d["_tpad"] = 30 if (n, "top") in patched else 0
        d["_bpad"] = 30 if (n, "bottom") in patched else 0
        d["h"] += d["_tpad"] + d["_bpad"]

    # 3. positions
    gaps = []
    for r in range(len(rows)):
        n = len(gap_links[r])
        if r == 0:
            g = LABEL_ZONE + n * CH_PITCH + 20 if n else 0
            if r in loop_top:
                g = max(g, LOOP_D + 20)
        else:
            g = max(MIN_VGAP, 2 * LABEL_ZONE + n * CH_PITCH + 24)
            g += 40 * ((r in loop_top) + (r - 1 in loop_bot))
        gaps.append(g)
    row_w = [sum(d["w"] for d in row) + HGAP * (len(row) - 1) for row in rows]
    side_loop = any(l.get("_route") == "loop" and l["_sa"] in ("left", "right") for l in links)
    width = max(row_w) + 2 * MARGIN + (2 * (LOOP_D + 40) if side_loop else 0)
    y = MARGIN + TITLE_H
    row_top, row_bot = [], []
    for ri, row in enumerate(rows):
        y += gaps[ri]
        x = (width - row_w[ri]) / 2
        for d in row:
            d["x"], d["y"] = x, y
            x += d["w"] + HGAP
        row_top.append(y)
        row_bot.append(y + max(d["h"] for d in row))
        y = row_bot[-1]
    height = y + MARGIN + (LOOP_D if len(rows) - 1 in loop_bot else 0)

    # 4. slot positions on each edge
    slots = defaultdict(list)
    for li, l in enumerate(links):
        A, B = names[l["device_a"]], names[l["device_b"]]
        slots[(A["name"], l["_sa"])].append((B["x"] + B["w"] / 2, li, "a"))
        slots[(B["name"], l["_sb"])].append((A["x"] + A["w"] / 2, li, "b"))
    for (dev, side), items in slots.items():
        d = names[dev]
        if side in ("left", "right"):
            items.sort(key=lambda t: t[1])            # link order -> parallel links line up
        else:
            items.sort()
        n = len(items)
        for k, (_, li, end) in enumerate(items):
            if side in ("top", "bottom"):
                px = d["x"] + d["w"] / 2 + (k - (n - 1) / 2) * TB_PITCH
                py = d["y"] if side == "top" else d["y"] + d["h"]
            else:
                px = d["x"] if side == "left" else d["x"] + d["w"]
                py = d["y"] + (d["h"] / 2 if n == 1 and d["_kind"] != "switch" else 28 + k * SIDE_PITCH)
            links[li]["_p" + end] = (px, py)

    # 5. routes (orthogonal), channels assigned per gap so parallel runs don't cross
    for r, gl in gap_links.items():
        if r == 0:
            lo, hi = MARGIN + TITLE_H, row_top[0]
        else:
            lo, hi = row_bot[r - 1], row_top[r]
        over = sorted([l for l in gl if l["_route"] == "over"],
                      key=lambda l: -abs(l["_pa"][0] - l["_pb"][0]))
        right, left = [], []
        for l in gl:
            if l["_route"] != "vert":
                continue
            up, low = (l["_pa"], l["_pb"]) if l["_pa"][1] < l["_pb"][1] else (l["_pb"], l["_pa"])
            (right if low[0] >= up[0] else left).append((up[0], l))
        order = over + [l for _, l in sorted(right, key=lambda t: -t[0])] + \
            [l for _, l in sorted(left, key=lambda t: t[0])]
        mid = (lo + hi) / 2 if r else hi - LABEL_ZONE - 10 - (len(order) - 1) * CH_PITCH / 2
        for k, l in enumerate(order):
            l["_cy"] = mid + (k - (len(order) - 1) / 2) * CH_PITCH

    pair_jog = Counter()
    for l in links:
        (x1, y1), (x2, y2) = l["_pa"], l["_pb"]
        if l["_route"] == "loop":
            ox, oy = OUT[l["_sa"]]
            pts = [(x1, y1), (x1 + ox * LOOP_D, y1 + oy * LOOP_D), (x2 + ox * LOOP_D, y2 + oy * LOOP_D), (x2, y2)]
        elif l["_route"] == "horiz":
            if abs(y1 - y2) < 0.5:
                pts = [(x1, y1), (x2, y2)]
            else:
                key = tuple(sorted((l["device_a"], l["device_b"])))
                mx = (x1 + x2) / 2 + (pair_jog[key] % 5 - 2) * 10
                pair_jog[key] += 1
                pts = [(x1, y1), (mx, y1), (mx, y2), (x2, y2)]
        elif abs(x1 - x2) < 0.5:
            pts = [(x1, y1), (x2, y2)]
        else:
            cy = l["_cy"]
            pts = [(x1, y1), (x1, cy), (x2, cy), (x2, y2)]
        l["_pts"] = pts
    return width, height


# ---------------------------------------------------------------- drawing helpers
def esc(s):
    return html.escape(str(s), quote=True)


def unit(p, q):
    dx, dy = q[0] - p[0], q[1] - p[1]
    ln = max((dx * dx + dy * dy) ** 0.5, 1e-6)
    return dx / ln, dy / ln


def pill(x, y, text, stroke, fill="#ffffff", color="#1f2937", vertical=False, bold=False):
    w = 6.7 * len(text) + 14
    h = 18
    rot = f' transform="rotate(90 {x:.1f} {y:.1f})"' if vertical else ""
    fw = ' font-weight="bold"' if bold else ""
    return (f'<g{rot}><rect x="{x - w / 2:.1f}" y="{y - h / 2:.1f}" width="{w:.1f}" height="{h}" rx="9" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="1.2"/>'
            f'<text x="{x:.1f}" y="{y + 4:.1f}" font-size="11"{fw} text-anchor="middle" fill="{color}">{esc(text)}</text></g>')


def show_speed(l):
    return l.get("_show_speed") and l.get("speed")


def label_points(l):
    """(position, text, vertical) for both port labels and the speed label."""
    pts = l["_pts"]
    out = []
    for (p, q), port in (((pts[0], pts[1]), l.get("port_a")), ((pts[-1], pts[-2]), l.get("port_b"))):
        if port:
            ux, uy = unit(p, q)
            vertical = abs(uy) > 0.5
            dist = (6.7 * len(port) + 14) / 2 + 8 if vertical else (6.7 * len(port) + 14) / 2 + 10
            out.append(((p[0] + ux * dist, p[1] + uy * dist), port, vertical, "port"))
    if show_speed(l) and l.get("_route") == "loop":     # speed pill just beyond the tip of the loop
        ox, oy = OUT[l["_sa"]]
        (ax, ay), (bx, by) = pts[1], pts[2]
        txt = str(l["speed"])
        off = 13 if oy else (6.7 * len(txt) + 14) / 2 + 4
        out.append((((ax + bx) / 2 + ox * off, (ay + by) / 2 + oy * off), txt, False, "speed"))
    elif show_speed(l):
        segs = list(zip(pts, pts[1:]))
        a, b = max(segs, key=lambda s: abs(s[0][0] - s[1][0]) + abs(s[0][1] - s[1][1]))
        out.append((((a[0] + b[0]) / 2, (a[1] + b[1]) / 2), str(l["speed"]),
                    abs(a[0] - b[0]) < abs(a[1] - b[1]), "speed"))
    return out


PATCH_COLOR = "#8a3a12"


def patch_marks(links):
    """Patch-panel labels drawn just inside each box.
    Returns [(kind, data)]: ("text", (x, y, text, anchor)) and ("bar", (x1, y1, x2, y2))."""
    per_edge = defaultdict(list)
    for l in links:
        for end, side in (("a", l["_sa"]), ("b", l["_sb"])):
            per_edge[(l["device_" + end], side)].append((l["_p" + end], l.get("patch_" + end, "")))
    marks = []
    for (_, side), items in per_edge.items():
        vertical_edge = side in ("left", "right")
        items.sort(key=lambda t: t[0][1] if vertical_edge else t[0][0])
        runs = []                                  # group neighbours with the same P#
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
                    marks.append(("text", (bx + inward * 8, (y1 + y2) / 2 + 4, patch,
                                           "start" if inward > 0 else "end")))
                else:
                    marks.append(("text", (x1 + inward * 10, y1 + 4, patch,
                                           "start" if inward > 0 else "end")))
            else:
                inward = 1 if side == "top" else -1
                if len(pts) > 1:
                    by = y1 + inward * 12
                    marks.append(("bar", (x1 - 6, by, x2 + 6, by)))
                    marks.append(("text", ((x1 + x2) / 2, by + inward * 14 + 4, patch, "middle")))
                else:
                    marks.append(("text", (x1, y1 + inward * 16 + 4, patch, "middle")))
    return marks


# ---------------------------------------------------------------- SVG
def svg_line(d, st, extra=""):
    """A path (or legend line) in a given style; DAC = hollow double line."""
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


def svg(names, links, cfg, width, height, title, subtitle, show_planned):
    styles = [style_for(l, cfg) for l in links]
    roles_used = list(dict.fromkeys(st["role"] for st in styles))
    cables_used = list(dict.fromkeys(l["_cable"] for l in links))
    legend_h = 76 if links else 0
    H = height + legend_h
    o = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.0f}" height="{H:.0f}" '
         f'viewBox="0 0 {width:.0f} {H:.0f}" font-family="{FONT}">',
         '<rect width="100%" height="100%" fill="#ffffff"/>']
    colors = set(cfg["roles"].values())
    o.append("<defs>" + "".join(
        f'<marker id="arr{c[1:]}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" '
        f'orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{c}"/></marker>' for c in colors) + "</defs>")
    o.append(f'<text x="{MARGIN}" y="{MARGIN + 6}" font-size="26" font-weight="bold" fill="#1f2a44">{esc(title)}</text>')
    o.append(f'<text x="{MARGIN}" y="{MARGIN + 32}" font-size="13" fill="#6b7280">{esc(subtitle)}</text>')

    labels = []
    for l, st in zip(links, styles):
        d = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in l["_pts"])
        tip = esc(f"{l.get('device_a')} {l.get('port_a') or '?'} ↔ {l.get('device_b')} {l.get('port_b') or '?'}"
                  f" · {st['role']}" + (f" · {l['_cable']}" if l["_cable"] else "")
                  + (f" {l.get('speed')}" if l.get("speed") else "")
                  + (f" · {l.get('notes')}" if l.get("notes") else ""))
        o.append(svg_line(d, st, f"<title>{tip}</title>"))
        color = st["color"]
        for (x, y), text, vertical, kind in label_points(l):
            if kind == "speed":
                labels.append(pill(x, y, text, color, fill=color, color="#ffffff", vertical=vertical, bold=True))
            else:
                labels.append(pill(x, y, text, color, vertical=vertical))
    o.extend(labels)

    for d in names.values():
        fill, stroke, tc, sc = SHAPE[d["_kind"]]
        x, y, w, h = d["x"], d["y"], d["w"], d["h"]
        rx = h / 2 if d["_kind"] == "ixia" else 8
        lines = [(d["name"], 16, "bold", tc)]
        sub = " · ".join(v for v in (d.get("model"), d.get("role")) if v)
        if sub:
            lines.append((sub, 12, "normal", sc))
        if d.get("emp_ip"):
            lines.append((("EMP " if d["_kind"] == "switch" else "") + d["emp_ip"], 12, "normal", sc))
        tp, bp = d.get("_tpad", 0), d.get("_bpad", 0)
        cy = y + tp + (h - tp - bp) / 2 - (len(lines) - 1) * 9
        txt = "".join(f'<text x="{x + w / 2:.1f}" y="{cy + i * 18 + 5:.1f}" font-size="{s}" font-weight="{fw}" '
                      f'text-anchor="middle" fill="{c}">{esc(t)}</text>' for i, (t, s, fw, c) in enumerate(lines))
        o.append(f'<g><title>{esc(d.get("notes") or d["name"])}</title>'
                 f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" rx="{rx:.1f}" '
                 f'fill="{fill}" stroke="{stroke}" stroke-width="1.8"/>{txt}</g>')

    for kind, m in patch_marks(links):
        if kind == "bar":
            o.append(f'<line x1="{m[0]:.1f}" y1="{m[1]:.1f}" x2="{m[2]:.1f}" y2="{m[3]:.1f}" '
                     f'stroke="{PATCH_COLOR}" stroke-width="2"/>')
        else:
            x, y, text, anchor = m
            o.append(f'<text x="{x:.1f}" y="{y:.1f}" font-size="12" font-weight="bold" '
                     f'text-anchor="{anchor}" fill="{PATCH_COLOR}">{esc(text)}</text>')

    if links:
        ly = height + 10
        o.append(f'<text x="{MARGIN}" y="{ly + 4}" font-size="12" font-weight="bold" fill="#374151">Connection</text>')
        lx = MARGIN + 84
        for r in roles_used:
            st = {"color": cfg["roles"].get(r, DEFAULT_ROLES["other"][0]), "width": 2.5, "double": False,
                  "dash": "", "arrow": "b" if r == "traffic" else None}
            o.append(svg_line(f"M{lx},{ly} L{lx + 36},{ly}", st) +
                     f'<text x="{lx + 44}" y="{ly + 4}" font-size="12" fill="#374151">{esc(r)}</text>')
            lx += 64 + 7 * len(r)
        ly += 24
        o.append(f'<text x="{MARGIN}" y="{ly + 4}" font-size="12" font-weight="bold" fill="#374151">Cable</text>')
        lx = MARGIN + 84
        items = []
        for c in cables_used:
            w, line, _ = cfg["cables"].get(c, cfg["cables"][""])
            items.append((c or "not recorded", {"color": "#4b5563", "width": w, "double": line == "double",
                                                "dash": "10 4 2 4" if line == "dashdot" else "", "arrow": None}))
        if any(l.get("_unk_port") and not l.get("_planned") for l in links):
            items.append(("switch port not recorded", {"color": "#4b5563", "width": 2, "double": False, "dash": DOTTED, "arrow": None}))
        if show_planned:
            items.append(("planned", {"color": "#4b5563", "width": 2, "double": False, "dash": DASHED, "arrow": None}))
        for name, st in items:
            o.append(svg_line(f"M{lx},{ly} L{lx + 36},{ly}", st) +
                     f'<text x="{lx + 44}" y="{ly + 4}" font-size="12" fill="#374151">{esc(name)}</text>')
            lx += 64 + 7 * len(name)
        shown = [c for c, v in cfg["cables"].items() if c and v[2]]
        o.append(f'<text x="{MARGIN}" y="{ly + 26}" font-size="11" fill="#6b7280">'
                 f'Speed shown on {"/".join(shown) or "no"} cables · P# inside a box = patch-panel port at that end, '
                 f'bracket = one patch group · hover a line for full details</text>')
    o.append("</svg>")
    return "\n".join(o)


# ---------------------------------------------------------------- draw.io
def drawio(names, links, cfg, title):
    cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
    ids = {}
    for i, d in enumerate(names.values()):
        cid = f"dev{i}"
        ids[d["name"]] = cid
        fill, stroke, tc, _ = SHAPE[d["_kind"]]
        sub = " · ".join(v for v in (d.get("model"), d.get("role")) if v)
        val = f"<b>{html.escape(d['name'])}</b>"
        if sub:
            val += f"<br>{html.escape(sub)}"
        if d.get("emp_ip"):
            val += f"<br>{'EMP ' if d['_kind'] == 'switch' else ''}{html.escape(d['emp_ip'])}"
        arc = "arcSize=50;absoluteArcSize=0;" if d["_kind"] == "ixia" else "arcSize=8;"
        cells.append(f'<mxCell id="{cid}" value="{esc(val)}" vertex="1" parent="1" '
                     f'style="rounded=1;whiteSpace=wrap;html=1;{arc}fillColor={fill};strokeColor={stroke};'
                     f'fontColor={tc};fontSize=12;">'
                     f'<mxGeometry x="{d["x"]:.1f}" y="{d["y"]:.1f}" width="{d["w"]:.1f}" height="{d["h"]:.1f}" as="geometry"/></mxCell>')
    for i, l in enumerate(links):
        sty = style_for(l, cfg)
        color = sty["color"]
        A, B = names[l["device_a"]], names[l["device_b"]]
        (ax, ay), (bx, by) = l["_pa"], l["_pb"]
        dash = f"dashed=1;dashPattern={sty['dash'].replace(' ', ' ')};" if sty["dash"] else "dashed=0;"
        st = (f"startArrow={'block' if sty['arrow'] == 'a' else 'none'};startFill=1;"
              f"endArrow={'block' if sty['arrow'] == 'b' else 'none'};endFill=1;html=1;rounded=0;edgeStyle=none;"
              f"{'shape=link;width=3;' if sty['double'] else ''}"
              f"strokeColor={color};strokeWidth={sty['width'] if not sty['double'] else 1.5};{dash}"
              f"exitX={(ax - A['x']) / A['w']:.3f};exitY={(ay - A['y']) / A['h']:.3f};exitDx=0;exitDy=0;"
              f"entryX={(bx - B['x']) / B['w']:.3f};entryY={(by - B['y']) / B['h']:.3f};entryDx=0;entryDy=0;")
        eid = f"link{i}"
        mids = l["_pts"][1:-1]
        geo = '<mxGeometry relative="1" as="geometry">'
        if mids:
            geo += '<Array as="points">' + "".join(f'<mxPoint x="{x:.1f}" y="{y:.1f}"/>' for x, y in mids) + "</Array>"
        geo += "</mxGeometry>"
        cells.append(f'<mxCell id="{eid}" value="" edge="1" parent="1" source="{ids[l["device_a"]]}" '
                     f'target="{ids[l["device_b"]]}" style="{st}">{geo}</mxCell>')
        # labels as free text boxes at the same spots as in the SVG
        for j, ((x, y), text, vertical, kind) in enumerate(label_points(l)):
            tw, th = 6.7 * len(text) + 14, 18
            fillc, fc = (color, "#ffffff") if kind == "speed" else ("#ffffff", "#1f2937")
            rot = "rotation=90;" if vertical else ""
            cells.append(f'<mxCell id="{eid}t{j}" value="{esc(text)}" vertex="1" parent="1" '
                         f'style="rounded=1;arcSize=50;html=1;fontSize=11;fillColor={fillc};strokeColor={color};'
                         f'fontColor={fc};{rot}{"fontStyle=1;" if kind == "speed" else ""}">'
                         f'<mxGeometry x="{x - tw / 2:.1f}" y="{y - th / 2:.1f}" width="{tw:.1f}" height="{th}" as="geometry"/></mxCell>')
    for j, (kind, m) in enumerate(patch_marks(links)):
        if kind == "bar":
            x1, y1, x2, y2 = m
            cells.append(f'<mxCell id="pb{j}" value="" vertex="1" parent="1" '
                         f'style="rounded=0;fillColor={PATCH_COLOR};strokeColor=none;">'
                         f'<mxGeometry x="{min(x1, x2) - 1:.1f}" y="{min(y1, y2) - 1:.1f}" '
                         f'width="{abs(x2 - x1) + 2:.1f}" height="{abs(y2 - y1) + 2:.1f}" as="geometry"/></mxCell>')
        else:
            x, y, text, anchor = m
            tw = 8 * len(text) + 6
            left = x if anchor == "start" else x - tw if anchor == "end" else x - tw / 2
            align = {"start": "left", "end": "right", "middle": "center"}[anchor]
            cells.append(f'<mxCell id="pt{j}" value="{esc(text)}" vertex="1" parent="1" '
                         f'style="text;html=1;fontSize=12;fontStyle=1;fontColor={PATCH_COLOR};align={align};'
                         f'verticalAlign=middle;spacing=0;">'
                         f'<mxGeometry x="{left:.1f}" y="{y - 12:.1f}" width="{tw:.1f}" height="16" as="geometry"/></mxCell>')
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

    width, height = plan(names, drawn)
    title = title or wb_path.stem.replace("_", " ")
    subtitle = subtitle or f"{len(names)} devices · {len(drawn)} links · generated {dt.date.today().isoformat()}"
    base = Path(out) if out else wb_path.with_suffix("")

    svg_text = svg(names, drawn, cfg, width, height, title, subtitle, show_planned)
    base.with_suffix(".svg").write_text(svg_text, encoding="utf-8")
    base.with_suffix(".drawio").write_text(drawio(names, drawn, cfg, title), encoding="utf-8")
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
        sys.exit("No workbooks given and none found in setups/")
    missing = [p for p in paths if not p.exists()]
    if missing:
        sys.exit("Workbook not found: " + ", ".join(map(str, missing)))
    single = len(paths) == 1
    ok = all([render(p, a.title if single else None, a.subtitle if single else None,
                     a.show_planned, a.out if single else None) for p in paths])
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
