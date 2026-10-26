# draw_topo

Lab topology diagrams generated from an Excel list of devices and cables.
Edit the workbook, run one command, and get an up-to-date diagram (`.svg`, plus a `.drawio` file for manual tweaks).

## Layout

```
draw_topo/
├── topo.py               # draws the diagrams
├── upgrade_workbook.py   # converts an older-format workbook to the current columns
├── requirements.txt
└── setups/               # one workbook per lab setup, diagrams are written next to it
    ├── topology.xlsx  → topology.svg, topology.drawio
    └── jupiter_setup1.xlsx  → jupiter_setup1.svg, jupiter_setup1.drawio
```

## Setup (once)

```
python -m venv venv
venv\Scripts\activate          # Windows   (Linux/macOS: source venv/bin/activate)
pip install -r requirements.txt
```

## Use

```
python topo.py                                   # redraw every workbook in setups/
python topo.py setups/jupiter_setup1.xlsx --title "Jupiter AOSX - Setup 1"
python topo.py setups/topology.xlsx --show-planned   # also draw planned links (dashed)
```

New setup: copy an existing workbook in `setups/`, rename it, and edit it.

## Workbook

| Sheet | What goes in it |
|---|---|
| **Devices** | `name`, `kind` (switch / ixia / server / mgmt), `model`, `role`, `emp_ip`, `tier` (row in the diagram, 1 = top; left-to-right = sheet order), `notes` |
| **Links** | One row per cable: `device_a`, `port_a`, `patch_a`, `device_b`, `port_b`, `patch_b`, `cable` (fiber / eth / dac), `speed`, `status` (active / planned / removed), `date_changed`, `notes`, `loop_side` |
| **Roles** | Colour per connection kind (edit colours only) |
| **Cables** | Line width / style per cable type, and whether the speed is printed on the line |

The connection kind is never typed in. It is worked out from the two devices:
switch–switch = **dut-dut**, switch–server = **traffic** (arrow to the server), switch–Ixia = **ixia**,
anything–mgmt = **mgmt**, same device on both ends = **loopback**.

How a line is drawn:
colour = connection kind · thickness / double line = cable type · speed label on fiber and DAC ·
dotted = a switch port is not recorded · dashed = planned ·
`P#` inside a box = patch-panel port, a bracket = several neighbouring links through the same patch port.

Add a connection: add a row. Remove one: set `status` to `removed` (keeps history) or delete the row.

The script warns when a port is used twice and stops when a device name is unknown or a loopback uses the same port on both ends.

## Putting it on Git

```
git init
git add .
git commit -m "Lab topology diagrams"
git remote add origin <your repo URL>
git push -u origin main
```

`venv/`, Excel lock files (`~$*.xlsx`) and backups (`*.bak.xlsx`) are ignored by `.gitignore`.
The generated `.svg` / `.drawio` files are committed on purpose so the diagrams can be viewed in the repo;
re-run `python topo.py` before committing a workbook change.
