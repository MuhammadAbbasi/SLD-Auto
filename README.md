# DC Single Line Diagram Generator

**Automatically generate DC Single Line Diagram (SLD) DXF files** for utility-scale solar PV plants from an Excel cable schedule.

The tool stamps a template inverter block for every inverter in the cable schedule, wires in all string/MPPT labels, auto-sizes each inverter section to match its MPPT count, and produces one A3 paper-space layout per transformer - ready to open in AutoCAD, BricsCAD, or any DXF viewer.

---

## Two Ways to Run

| Mode | File | Best for |
|---|---|---|
| Windows EXE | `A176LAB_SLD_Generator.exe` | End users, no Python needed |
| Python GUI | `code/smart_sld_generator.py` | Developers, customisation |
| Web App (Docker) | `webapp/` | Team use, browser-based UI |

---

## Standalone Windows EXE

Download `A176LAB_SLD_Generator.exe` and double-click. No installation required.

The GUI opens directly. Fill in the fields and press **F5** or click **Execute Generation**.

---

## Run from Source (Python GUI)

### Requirements

```
Python 3.10+
```

```bash
pip install ezdxf openpyxl customtkinter Pillow
```

| Package | Purpose |
|---|---|
| `ezdxf` | Read/write DXF files |
| `openpyxl` | Read Excel cable schedules |
| `customtkinter` | Modern Material-style GUI |
| `Pillow` | Logo image processing (optional) |

### Launch

```bash
cd "SLD Diagram/code"
python smart_sld_generator.py
```

Press **F5** or click **Execute Generation**.

### Template path (optional)

Set the environment variable so the template field pre-fills on launch:

```bash
# Windows
set SLD_TEMPLATE_PATH=C:\path\to\your\template.dxf

# PowerShell
$env:SLD_TEMPLATE_PATH = "C:\path\to\your\template.dxf"
```

---

## Web App (Docker)

See [webapp/README.md](webapp/README.md) for Docker deployment instructions.

---

## GUI Overview

### Section 1 - Files

| Field | Description |
|---|---|
| Excel Cable Schedule | Project cable schedule (must contain sheet `2E802-3`) |
| Template DXF | Base DXF containing one stamped inverter block for Inverter 1.1 |
| Output DXF | Destination file - auto-filled from the Excel filename |

### Section 2 - PV System Configuration

| Field | Default | Notes |
|---|---|---|
| Panel Model | *(blank)* | Shown in string labels, e.g. `JA Solar JAM72D42-625/LB` |
| Panels per String | 28 | Appended to string labels |
| Inverter Model | *(blank)* | Shown in inverter title; selecting from the list auto-fills DC/AC power |
| DC Power (KWp) | 350 | Set to `0` to auto-calculate from `strings x panels x Wp / 1000` |
| AC Power (KWac) | 320 | Shown in inverter title |
| Temperature Rating | 40 C | Shown as `@40C` in the title |
| Transformer Power | *(blank)* | Shown under the cabin header |

### Section 3 - Display Options

| Option | Default | Notes |
|---|---|---|
| Show cable lengths | off | Appends `L+`, `L-`, section to string labels |
| Hide string details | on | Hides panel count/model from labels |
| Show annotation circle | on | Large red circle from the template |

### Section 4 - Advanced Settings

| Field | Default |
|---|---|
| Column spacing | 14323 |
| Row spacing | 12036 |
| Circle radius | 24.59 |
| Text size | 60.44 |
| Heavy section | 1x10 |
| Heavy linetype | TRATTEGGIATA |
| Heavy color (ACI) | 40 |
| Heavy layer | TRATTEGGIATA |

---

## Excel Format

The workbook must contain a sheet named **`2E802-3`** (case-insensitive).

The tool auto-detects headers by scanning rows 26-32 for a row containing `String Name`. Expected columns:

| Header | Content | Example |
|---|---|---|
| `Inverter` | Inverter ID in `T.I` format | `1.2` = Transformer 1, Inverter 2 |
| `String Name` | String identifier | `1.2.5` |
| `MPPT` | MPPT channel number | `3` |
| `Section` | Cable cross-section | `1x6` |
| `Module Type` | Panel power in Wp | `625` |
| `Posizione Stringa` | Tracker/table position | `Table 3` |

An optional second sheet named **`Inverter To String`** provides cable routing lengths:

| Column | Content |
|---|---|
| 1 | String name |
| 2 | L+ length (m) |
| 3 | L- length (m) |
| 5 | Table/position reference |

---

## Output

- One DXF model-space with all inverter sections grouped by transformer
- One A3 landscape paper-space layout per transformer (`Tx1`, `Tx2`, ...), auto-scaled
- Labels per inverter:

| Label | Example |
|---|---|
| Inverter title | `INVERTER 2.4 - Sungrow SG350HX - P= 350 KWp - P= 320 KWac @40C` |
| Cabin header | `CABIN 2` |
| String label | `String 2.4.7 - 28x JA Solar JAM72D42-625/LB 625Wp` |
| Unused slot | `reserve` |

---

## Template DXF Requirements

The template DXF must contain a stamped **Inverter 1.1** block in the Y-band `159400 - 168000` drawing units (auto-detected from the `INVERTER 1.1 ... P= ...` MTEXT anchor) with:

- `MTEXT` matching `INVERTER 1.1 ... P= ...` - inverter title placeholder
- `MTEXT` matching `CABIN \d+` - cabin header placeholder
- `MTEXT` matching `Cabin Tx.\d+ ... Inverter` - cabin label placeholder
- Port labels in `N-M` format - maps to MPPT/port positions
- String labels `String \d+.\d+.\d+` - stamped with actual string names

All other geometry (lines, polylines, arcs, circles, inserts) is copied verbatim for each inverter.

To re-extract template geometry from a new master DXF:

```bash
cd code
python extract_template.py
```

Update `DXF_PATH` inside `extract_template.py` to point to your template before running.

---

## Building the EXE

```bash
cd code
pip install pyinstaller
pyinstaller A176LAB_SLD_Generator.spec
```

The compiled executable is written to `code/dist/`.

---

## Project Structure

```
SLD Diagram/
+-- code/
|   +-- smart_sld_generator.py   # Main engine - CLI + GUI (production)
|   +-- sld_gui_v7.py            # Standalone tkinter GUI (alternative)
|   +-- extract_template.py      # One-time utility: DXF geometry -> JSON
|   +-- old/                     # Archived previous versions (gitignored)
+-- webapp/                      # Docker web application
|   +-- app.py                   # Flask server
|   +-- sld_core.py              # Engine wrapper
|   +-- templates/index.html     # Browser UI
|   +-- code/                    # Copy of smart_sld_generator.py
|   +-- template/                # Mount your template.dxf here
|   +-- Dockerfile
|   +-- docker-compose.yml
|   +-- requirements.txt
|   +-- README.md
+-- .gitignore
+-- README.md
```

---

## Credits

Developed by **Muhammad Abbasi** - Data Scientist and Automation Engineer at *A176 LAB*
