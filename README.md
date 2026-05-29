# DC Single Line Diagram Generator

> **A176 LAB – Think different project**

Automatically generate DC Single Line Diagram (SLD) DXF files for utility-scale solar PV plants from an Excel cable list.  
The tool stamps a template inverter block for every inverter defined in the cable list, wires in all string/MPPT labels, and produces one A3 paper-space layout per transformer — ready to open in AutoCAD, BricsCAD, or any DXF-compatible viewer.

---

## Features

- **GUI application** (`sld_gui_v4.py`) — four-tab interface, no command-line knowledge required
- **Auto-layout** — inverters arranged in rows (one per transformer) and columns automatically
- **Fully parameterised labels** — inverter title, cabin header, string labels, and panel info all driven by user inputs
- **Auto-calculate DC power** — set DC power to 0 and the tool computes `strings × panels × Wp ÷ 1000` per inverter
- **A176 LAB logo** embedded in every paper-space layout (top-right corner) and shown in the GUI header
- **Background generation** — UI stays responsive; progress shown in the live log
- **Input validation** — clear error messages before any file is touched

---

## Requirements

| Package | Purpose |
|---------|---------|
| `ezdxf` | Read/write DXF files |
| `openpyxl` | Read Excel cable lists |
| `Pillow` | Logo display in GUI and image pixel-size detection *(optional but recommended)* |

```bash
pip install ezdxf openpyxl Pillow
```

Python **3.10+** required (walrus operator `:=` used internally).

---

## Quick Start

```bash
cd "SLD Diagram/code"
python sld_gui_v4.py
```

Press **F5** or click **▶ Generate SLD** on the Generate tab.

---

## GUI Overview

### Tab 1 — Files

| Field | Description |
|-------|-------------|
| Excel Cable List | Project cable schedule — must contain sheet `2E802-3` (see [Excel Format](#excel-format)) |
| Output DXF | Destination file — **auto-filled** to the same folder and base-name as the Excel file when you browse |

### Tab 2 — Equipment

#### Solar Panel
| Field | Default | Notes |
|-------|---------|-------|
| Panel Model | *(blank)* | Shown in each string label, e.g. `JA Solar JAM72S20-460` |
| Panels per String | 20 | Appended to string labels, e.g. `String 1.2.3 - 20× JA Solar 460Wp` |

#### Inverter Specs
| Field | Default | Notes |
|-------|---------|-------|
| Inverter Model | *(blank)* | Shown in inverter title, e.g. `Huawei SUN2000-330KTL` |
| Inverter Max MPPTs | Auto | Maximum number of MPPT channels (detects from Excel if set to `Auto`) |
| DC Power per Inverter | 350 KWp | Set to **0** to auto-calculate from panel data |
| AC Power | 320 KWac | Shown in inverter title |
| Temperature Rating | 40 °C | Shown as `@40°C` in title |
| Show Cable Info | False (unchecked) | If checked, appends cable lengths (L+, L-) and section to string labels |

### Tab 3 — Workspace Parameters

#### Array Grid Layout Steps
| Field | Default | Notes |
|-------|---------|-------|
| Horizontal Grid Col Step | 15000 units | Spacing between stamped inverter columns |
| Vertical Grid Row Step | 10200 units | Spacing between stamped transformer rows |

#### Visual Node Elements
| Field | Default | Notes |
|-------|---------|-------|
| Module Circle Radius | 24.59 units | Radius of terminal switch/disconnector circles |
| String Label Text Height | 60.44 units | Font height of generated string label texts |

#### Heavy Cable Run Custom Styling
| Field | Default | Notes |
|-------|---------|-------|
| Target Heavy Section | 1x10 mm² | Section to format differently (e.g. 1x10) |
| Heavy Run Linetype | TRATTEGGIATA | AutoCAD linetype to apply to matching heavy runs |
| Heavy Run Color (ACI) | 40 | AutoCAD Color Index (ACI) used to draw heavy runs |
| Heavy Run Layer Name | TRATTEGGIATA | Layer name where heavy runs will be placed |

### Tab 4 — Generate

- **▶ Generate SLD (F5)** — starts generation in a background thread
- Live log shows each step: Excel read → internal template load → entity extraction → section stamping → paper-space creation → save
- Success/error dialogs on completion

---

## Excel Format

Sheet name: **`2E802-3`**

| Column | Content | Example |
|--------|---------|---------|
| 1 | Inverter ID (`T.I` format) | `1.2` → Transformer 1, Inverter 2 |
| 3 | String name | `1.2.5` |
| 4 | MPPT number | `3` |

Rows without an inverter ID in column 1 are treated as continuation rows of the current inverter.

---

## Output

- One **DXF model-space** containing all inverter sections for all transformers
- One **A3 landscape paper-space layout** per transformer (`Tx1`, `Tx2`, …), auto-scaled to fit the full inverter row
- **A176 LAB logo** placed in the top-right corner of each layout (28 × 28 mm)
- Labels generated per inverter section:

  | Label | Example |
  |-------|---------|
  | Inverter title | `INVERTER 2.4 - Huawei SUN2000-330KTL - P= 350 KWp - P= 320 KWac @40°C` |
  | Cabin header | `CABIN 2` *(+ transformer power if supplied)* |
  | Cabin label | `Cabin Tx.2 / Inverter 2.4` |
  | String label | `String 2.4.7 - 20× JA Solar JAM72S20-460 460Wp` |
  | Unused slot | `reserve` |

---

## Project Structure

```
SLD Diagram/
├── code/
│   ├── sld_gui_v4.py        # GUI application (latest entry point)
│   ├── extract_template.py  # Script to extract template DXF geometry to JSON
│   ├── template_data.json   # Internal template geometry data loaded by GUI
│   ├── generate_sld.py      # Original CLI script (standalone, no GUI)
│   └── logoA176LAB.jpg      # A176 LAB logo
├── YANEL/
│   ├── 26S001_2E103 - DC Single Line Diagram.dxf   # Template DXF (stamp source)
│   └── Yanel - Lista Cavi - Cavi LV-DC.xlsx
├── Example/
│   ├── generate_sld.py      # Example project version
│   └── Eaton Socon - Lista Cavi - Cavi LV-DC.xlsx
├── 2025.017 - PV WYMONDLEY/
│   └── Priory Farm - Lista Cavi - Cavi LV-DC.xlsx
├── .gitignore
└── README.md
```

> **Large binary files** (template DXF, generated DXFs, Excel cable lists) are excluded from version control via `.gitignore`.  
> For the template DXF, use [Git LFS](https://git-lfs.com): `git lfs track "*.dxf"`.

---

## Auto-Calculate DC Power

Set **DC Power per Inverter** to `0` and fill in:
- Panel Power (Wp)
- Panels per String

The tool counts the total strings assigned to each inverter from the Excel file and computes:

```
DC power (KWp) = total_strings × panels_per_string × panel_power_Wp / 1000
```

This produces a unique per-inverter DC value if inverters have different string counts.

---

## Template DXF & Extract Requirements

The template geometry is loaded internally from `code/template_data.json`. If you need to update it from a new master DXF template:
1. Update `DXF_PATH` in `code/extract_template.py` to point to your new DXF template.
2. Run `python extract_template.py` to regenerate `template_data.json`.

The master DXF template must contain **Inverter 1.1** in the Y-band `159 400 – 168 000` drawing units with:

- `MTEXT` matching `INVERTER 1.1 … P= …` → used as the inverter title template
- `MTEXT` matching `CABIN \d+` → used as the cabin header template
- `MTEXT` matching `Cabin Tx.\d+ … Inverter` → used as the cabin label template
- Port labels in `N-M` format → mapped to MPPT/port positions
- String labels `String \d+\.\d+\.\d+` → stamped with actual string names

All other geometry (lines, polylines, arcs, circles, inserts) is copied verbatim.

---

## Credits

Developed by **Muhammad Abbasi** — Data Scientist and Automation Engineer at *A176 LAB*  