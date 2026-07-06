# SLD Generator - Workflow Documentation

Auto-generates DC Single Line Diagrams (DXF) from an Excel cable schedule by stamping
and procedurally rebuilding an inverter block template for each inverter in each cabin.

---

## Table of Contents

1. [High-Level Flow](#1-high-level-flow)
2. [Excel Parsing](#2-excel-parsing)
3. [Template Loading and Geometry Extraction](#3-template-loading-and-geometry-extraction)
4. [Geometry Classification and Frame Splitting](#4-geometry-classification-and-frame-splitting)
5. [X-Axis Bus Positions](#5-x-axis-bus-positions)
6. [Per-Inverter Drawing Loop](#6-per-inverter-drawing-loop)
   - [Frame Stretch / Shrink](#frame-stretch--shrink)
   - [Section A - Static Frame](#section-a---static-frame)
   - [Section B - String Rows and MPPT Combiners](#section-b---string-rows-and-mppt-combiners)
   - [Section C - DC Switch Brackets](#section-c---dc-switch-brackets)
   - [Section D - Headers](#section-d---headers)
7. [MPPT Layout and Reserved Slots](#7-mppt-layout-and-reserved-slots)
8. [String Combinations and Cable Drawing](#8-string-combinations-and-cable-drawing)
9. [DC Switch Grouping](#9-dc-switch-grouping)
10. [Cabin Stacking and Column Layout](#10-cabin-stacking-and-column-layout)
11. [Paper Space Viewports](#11-paper-space-viewports)
12. [Data Flow Diagram](#12-data-flow-diagram)
13. [Geometry Coordinate Map](#13-geometry-coordinate-map)

---

## 1. High-Level Flow

```
Excel (.xlsx)
    |
    |-- Sheet "2E802-3"     --> string schedule (Cabin / Inverter / MPPT / section / Wp)
    |-- Sheet "Inverter To String" --> routing cable lengths
    |
    v
[ Parse Excel ]
    |
    v
[ Load template DXF ]    <-- single GELA inverter column slice
    |
    v
[ Extract template geometry ]
    |
    |-- Classify MTEXT entities (title, cc_side, ac_side, string_label, ...)
    |-- Measure row_pitch from port-label Y gaps
    |-- Detect circles (terminal points), cable_x_end, BUS_X, OUTER_X
    |-- Split into: frame (static) + repeating (per-string-row)
    |
    v
[ Clear model space + all paper layouts ]
    |
    v
FOR EACH CABIN (T):
  |  Pre-calculate Y-offset for this cabin (stacked below previous cabin)
  |
  FOR EACH INVERTER (I) in cabin:
    |
    |-- Section A: place stretched/shrunk static frame
    |-- Section B: procedurally draw string rows + MPPT combiner brackets
    |-- Section C: procedurally draw DC switch brackets
    |-- Section D: place inverter title / cabin header / cabin label
    |
    v
  Create A3 Paper Space viewport for this cabin
    |
    v
Save output DXF
```

---

## 2. Excel Parsing

Two sheets are read from the Excel workbook.

**Sheet `2E802-3`** (master cable schedule):

The header row is auto-detected by scanning rows 26-32 for a cell containing "string name".
Columns are then mapped dynamically (inverter, string name, MPPT number, cable section, panel Wp).

Data is grouped into:

```
excel[(T, I)][mppt_number] = [
    {name, wp, l_plus, l_minus, section},
    ...  # one entry per string
]
```

Where T = cabin (transformer) number, I = inverter number within the cabin.

**Sheet `Inverter To String`** (routing lengths):

```
its_data[string_id] = {l_plus, l_minus, table_num}
```

These cable lengths are merged into the string data and shown in the label
when `show_cable_info` is enabled.

---

## 3. Template Loading and Geometry Extraction

The reference DXF (GELA 9-column template) is loaded by `ezdxf`.

A Y-band window is determined by locating the "INVERTER 1.1" MTEXT anchor. All entities
within `[anchor_y - 10000, anchor_y + 5000]` are extracted.

An X-crop is then applied to isolate a single inverter column:

```
xmin         = leftmost x coordinate of any entity in the band
xcut         = xmin + COL_STEP          (geometry cutoff)
_mtext_xcut  = xmin + COL_STEP * 1.20  (MTEXT cutoff, 20% wider to keep right-aligned labels)
```

This prevents the 9-column template's texts from bleeding into the output.

Each entity is converted to a plain dict (`_extract_entity`), classified, and stored in `tmpl`.

Before extraction, block definitions are scanned to read the original panel count (e.g.
"28 PV modules") and then update it to match `panels_per_string`. This must happen before
the count label is re-read, since the update changes the text in-place.

---

## 4. Geometry Classification and Frame Splitting

### MTEXT Classification (`_classify_mtext`)

| Class | Detection keyword/pattern | Role |
|-------|---------------------------|------|
| `title` | matches `INVERTER \d+\.\d+` | inverter title row (replaced per inverter) |
| `cabin_header` | starts with "CABIN" | cabin label above the box |
| `cabin_label` | matches "Cabin Tx." pattern | right-side cabin/inverter reference |
| `string_label` | matches `String \d+\.\d+\.\d+` | per-string row labels (repeating) |
| `port_label` | matches `\d+-\d+` (e.g. "8-1") | MPPT-port connection labels (repeating) |
| `panel_count` | contains "PV module" | panel count text inside the module strip |
| `fixed` | everything else | static labels (CC side, AC side, cable specs, etc.) |

### Repeating Entity Filter (`_is_repeating`)

The `frame` list contains only the static scaffolding. Entities removed as "repeating"
(i.e. to be redrawn procedurally in Section B) are:

- All MTEXT with class `string_label` or `port_label`
- All CIRCLE entities with radius < 50 and x > 20800 (terminal connection points)
- LWPOLYLINE entities with any vertex in x-range `[20200, 20520]` (combiner bus zone)
- Short LWPOLYLINE decorations just right of the terminal area (string cables)
- INSERT block references in the terminal area (small symbols/arrows)

---

## 5. X-Axis Bus Positions

All x-coordinates in the drawing are absolute values from the template geometry.

```
Template inverter box (schematic cross-section, left to right):
---------------------------------------------------------------------------
|  inverter box  |  DC switch  |  MPPT bus  |  terminal circles  | strings |
---------------------------------------------------------------------------
 BOX_BUS_X       OUTER_X       BUS_X         circ_x
  ~19731          ~20243        ~20471         ~20945
```

| Constant | Value | Purpose |
|----------|-------|---------|
| `BOX_BUS_X` | ~19731 | Left edge of combiner region; horizontal cable from DC switch enters here |
| `OUTER_X` | ~20243 | Right edge of DC switch bracket |
| `BUS_X` | ~20471 | MPPT combiner vertical bus |
| `circ_x` | ~20945 | X position of all terminal circles (connection points) |
| `cable_x_end` | circ_x + ~1225 | Right end of the horizontal string cable |

These values are auto-detected from template geometry, not hard-coded.

---

## 6. Per-Inverter Drawing Loop

For each cabin T and inverter I (indexed left to right):

```
dxo = inverter_index * col_spacing     # horizontal offset
dyo = cabin_y_offset[T]                # vertical offset (negative = lower)
```

### Frame Stretch / Shrink

The template has a fixed number of rows (`template_rows = max_tmpl_m * 2` for a
16-MPPT / 2-strings/MPPT template = 32 rows).

Each real inverter may need a different count: `total_rows = num_mppts * eff_k`.

```
TOP_Y       = first_circle_y            # Y of the topmost terminal circle
bottom_y    = TOP_Y - (total_rows - 1) * row_pitch
tmpl_bottom_y = Y of the lowest terminal circle in the template

frame_shift = bottom_y - tmpl_bottom_y
   > 0: box is SHORTER than template (fewer rows) - bottom rises
   < 0: box is TALLER than template (more rows) - bottom drops
   = 0: exact match

split_y     = max(bottom_y, tmpl_bottom_y) - row_pitch * 0.5
EXTRA       = -frame_shift
```

Every static entity below `split_y` is shifted by `EXTRA` (i.e. by `-frame_shift`),
so the box bottom and lower annotations always sit just below the last drawn row.

```
Tall inverter (more rows than template):
                               template bottom
                                     |
   [ row 1 ]                         v
   [ row 2 ]     split_y ------>  --------
   [ row 3 ]     (below this      [  box  ]    <-- template bottom entities
   [ row 4 ]      shift by EXTRA)  [bottom ]       shifted down by |EXTRA|
   [ row 5 ]
       ^-- new bottom


Short inverter (fewer rows than template):
   [ row 1 ]
   [ row 2 ]     split_y ------>  --------
                                  [  box  ]    <-- template bottom entities
                                  [bottom ]       shifted up by EXTRA (positive)
```

### Section A - Static Frame

All entities in `frame` are placed with `_place_entity_stretched`. Special handling:

| Entity | Action |
|--------|--------|
| Large red CIRCLE (annotation marker) | Skipped if `show_annot_circle = False` |
| MTEXT class `title/cabin_header/cabin_label` | Skipped here (placed in Section D) |
| MTEXT class `panel_count` | Text updated with `panels_per_string` and `panel_model` |
| MTEXT containing "mmq" | Replaced with actual cable section used by this inverter |
| MTEXT containing inverter model name | Replaced with configured `inverter_model` |
| MTEXT with "AC output power" or "Nominal AC voltage" | Rebuilt with `ac_power_kwac`, `ac_power_30c`, `max_ac_current` |
| MTEXT with "CC side" or "MPPT range" | Rebuilt with MPPT count, `max_vdc`, current specs, `mppt_voltage_range`; y-shifted by `frame_shift` |
| All other entities | Placed with stretch transformation |

### Section B - String Rows and MPPT Combiners

This is the core procedural section. A global row counter `gr` advances by 1 for each
drawn slot.

```
FOR m = 1 to num_mppts:
  grp_ys = []
  slots = build_slot_list(m)   # [(slot_p, is_active, lookup_p), ...]

  FOR (slot_p, active, lookup_p) in slots:
    yc = TOP_Y - gr * row_pitch        # Y position of this row

    Draw terminal CIRCLE at (circ_x + dxo, yc + dyo)
      - color 7 (white) if active, 8 (grey) if reserve

    Draw LEFT cable:  BUS_X --> circle left edge  [horizontal LWPOLYLINE]
      - color/linetype from wire style (heavy section = dashed)

    Draw RIGHT cable: circle right edge --> cable_x_end  [horizontal LWPOLYLINE]
      - uses _panel_cable_right_x for the first row of MPPT 1 (panel-detail strip)

    If active: place row decoration INSERT symbols (small arrows etc.)

    Place PORT LABEL  "{m}-{slot_p}"  above circle at (PORT_X, yc + lbl_dy)

    Place STRING LABEL  "String T.I.m.slot_p - Nx Model"  above circle
      - first row of MPPT 1: placed higher at _panel_lbl_y_offset (above panel strip)
      - all other rows: placed at lbl_dy above circle

    grp_ys.append(yc)
    gr += 1

  _draw_combiner(dxo, dyo, grp_ys, "MPP{m}"):
    yb0, yb1 = first and last Y in grp_ys
    yc_mid   = (yb0 + yb1) / 2
    Draw vertical   BUS_X from yb0 to yb1
    Draw horizontal OUTER_X --> BUS_X at yc_mid   (MPPT center)
    Place MPP label at yc_mid
```

### Section C - DC Switch Brackets

DC switches group `mppts_per_switch` consecutive MPPTs under one bracket.

```
n_sw = ceil(num_mppts / mppts_per_switch)

FOR s = 0 to n_sw-1:
  first_m = s * mppts_per_switch + 1
  last_m  = min((s+1) * mppts_per_switch, num_mppts)

  y_first_mid = center-Y of MPPT first_m   (TOP_Y - ((first_m-1)*eff_k + (eff_k-1)/2) * row_pitch)
  y_last_mid  = center-Y of MPPT last_m

  IF mppts_per_switch > 1:
    Draw vertical at OUTER_X from y_first_mid to y_last_mid

  Draw horizontal BOX_BUS_X --> OUTER_X at y_sw_mid = (y_first_mid + y_last_mid) / 2

  Place "DC SWITCH {s+1}" label at y_sw_mid
```

### Section D - Headers

```
Place INVERTER title MTEXT:   "INVERTER T.I - Model - P= X,XX KWp - P= X KWac @40°C"
Place CABIN header MTEXT:     "CABIN T\nTransformer Power"
Place CABIN label MTEXT:      "Cabin Tx.T\nInverter T.I"
```

---

## 7. MPPT Layout and Reserved Slots

A JSON dict `mppt_layout` allows fine control over which string slots are active vs reserved:

```json
{
  "1": [1, 2, 3, 4, "R"],
  "2": [1, 2, "R", 3, "R"]
}
```

- Integer value: active slot, drawn in color 7 with full string label
- `"R"`: reserved slot, drawn in color 8 with "reserve" label

When `mppt_layout` is not provided, the generator auto-detects it by scanning
the template's string labels for the word "reserve". If none found and
`strings_per_mppt >= 3`, it builds a default layout with all active slots.

`eff_k` = the effective slots per MPPT row = max(strings_per_mppt, actual_max_strings,
layout_max_row_length). This is the number of rows drawn per MPPT.

---

## 8. String Combinations and Cable Drawing

### String Label Format

```
"String T.I.m.p - Nx ModelName"        (with panel model known)
"String T.I.m.p - NP XxxWp"            (no model, just Wp)
"String T.I.m.p - Nx ModelName (L+=12.500m, L-=11.200m, 1x6)"  (with cable info)
```

Where:
- T = cabin number
- I = inverter number
- m = MPPT number
- p = port/slot number within MPPT

### Cable Wire Style

| Condition | Linetype | Color | Layer |
|-----------|----------|-------|-------|
| Active, section matches `heavy_section` (e.g. "1x10") | TRATTEGGIATA (dashed) | 40 | TRATTEGGIATA |
| Active, other section | Continuous | 40 | 0 |
| Reserve / inactive | Continuous | 8 (grey) | 0 |

### Panel-Detail Strip (First Row of MPPT 1)

The first string slot (MPPT 1, slot 1) gets special treatment:
- Its string label is placed at `_panel_lbl_y_offset` above the circle (higher than normal),
  sitting above the panel module strip diagram
- Its right-side cable extends to `_panel_cable_right_x` (the right edge of the panel strip),
  which is wider than `cable_x_end`

---

## 9. DC Switch Grouping

```
Example: 6 MPPTs, mppts_per_switch = 2

Switch 1: MPPTs 1-2
Switch 2: MPPTs 3-4
Switch 3: MPPTs 5-6

Geometry (schematic, Y increases upward):

  MPPT 1 center -----o-------
                     |       |
  MPPT 2 center -----o       |---  "DC SWITCH 1"
                     |       |
  [OUTER_X vertical] [BOX_BUS_X horizontal]

  MPPT 3 center -----o-------
                     |       |
  MPPT 4 center -----o       |---  "DC SWITCH 2"
  ...
```

When `mppts_per_switch = 1` (each MPPT has its own switch), the OUTER_X vertical
is omitted and only the horizontal feed to BOX_BUS_X is drawn.

---

## 10. Cabin Stacking and Column Layout

### Column Layout (within a cabin)

Inverters within a cabin are placed left to right with spacing `col_spacing`:

```
Inverter 1.1     Inverter 1.2     Inverter 1.3
   dxo=0          dxo=col_spacing  dxo=2*col_spacing
```

### Cabin Stacking (vertical)

Cabins are stacked downward (decreasing Y). Each cabin's Y-offset is computed
before drawing begins:

```
cabin_y_offset[T1] = 0
cabin_y_offset[T2] = -(row_spacing + max_stretch_in_cabin_T1)
cabin_y_offset[T3] = cabin_y_offset[T2] - (row_spacing + max_stretch_in_cabin_T2)
...
```

`max_stretch` for a cabin = the height extension of its tallest inverter compared
to the template. This guarantees no overlap between cabins even when inverters
have different numbers of MPPTs.

---

## 11. Paper Space Viewports

One A3 (420 x 297 mm) paper space layout is created per cabin:

```
Layout name:  "Tx{T}"  (e.g. "Tx1", "Tx2")

Viewport center (model space):
  cx = x-center of template column + (n_inv - 1) * col_spacing / 2
  cy = cabin_y_offset[T] + (tmpl_y_max + min_bottom_y) / 2

View height:  max(cabin_height * 1.05, row_width / (420/297) * 1.05)
  (ensures all inverters fit horizontally and vertically)
```

---

## 12. Data Flow Diagram

```
Excel
 |
 |  excel[(T,I)][mppt] = [string_data, ...]
 v
generate(cfg)
 |
 +--[1] parse_excel --> excel dict
 |
 +--[2] load_template DXF
 |       |
 |       +--> extract column slice entities (tmpl)
 |       +--> classify MTEXT (tmpl_texts)
 |       +--> measure row_pitch, detect circ_x, cable_x_end, BUS_X, OUTER_X
 |       +--> build frame (tmpl - repeating entities)
 |
 +--[3] clear model space + paper layouts
 |
 +--[4] FOR EACH CABIN T:
 |       |
 |       +-- pre-calc cabin_y_offset[T]
 |       |
 |       FOR EACH INVERTER I:
 |         |
 |         +-- compute frame_shift, split_y, EXTRA
 |         |
 |         +-- Section A: place frame with stretch transform
 |         |     (static box, AC spec, CC spec, cable section label, model label)
 |         |
 |         +-- Section B: FOR EACH MPPT m:
 |         |     FOR EACH SLOT p:
 |         |       draw circle + left cable + right cable
 |         |       draw port label + string label
 |         |     _draw_combiner (BUS_X vertical + OUTER_X horizontal)
 |         |
 |         +-- Section C: FOR EACH SWITCH group:
 |         |     draw OUTER_X vertical + BOX_BUS_X horizontal
 |         |     draw DC SWITCH label
 |         |
 |         +-- Section D: inverter title + cabin header + cabin label
 |
 +--[5] create A3 paper space viewports
 |
 +--[6] save DXF
```

---

## 13. Geometry Coordinate Map

This diagram shows the horizontal structure of one inverter column in model-space
coordinates (all values relative to the template's xmin ~ 16852):

```
  x = 16852       19731        20243      20471    20945   22170
  |                |            |          |        |        |
  | INVERTER BOX   | BOX_BUS_X  | OUTER_X  | BUS_X  | circ_x | cable_x_end
  |                |            |          |        |        |
  |                |            |          |        |        |
  |   AC/CC spec   |<--- DC switch horiz -->|        |        |
  |   box outline  |            |<--MPPT horiz -->   |        |
  |                |            |          |        o---------> string cable
  |                |            |  [MPPT   o        o (circle)
  |                |            |  vertical|        o
  |                |            |  bus]    o---------> string cable
  |                |            |          |        |
  |                |            [DC switch |        |
  |                |            vertical]  |        |
  |                |                       |        |
  |________________|_______________________|________|
```

```
Y positions (absolute, from template):

  first_circle_y  ~ 163529     <- topmost terminal (MPPT 1, slot 1)
                                   also = TOP_Y
  ...
  (each row: row_pitch ~ 108.84 units apart)
  ...
  tmpl_bottom_y   ~ 160222     <- bottom terminal (last circle)
  box_bottom      ~ 160147     <- bottom edge of inverter box outline
  CC text y       ~ 160726     <- DC-side spec block (CC side)
```

Row Y formula:
```
y_of_row(mppt_m, slot_p) = TOP_Y - ( (m-1)*eff_k + (slot_index) ) * row_pitch
```

MPPT center Y formula:
```
y_of_mppt_center(m) = TOP_Y - ( (m-1)*eff_k + (eff_k-1)/2 ) * row_pitch
```
