# -*- coding: utf-8 -*-
"""
extract_template.py
===================
One-time utility: reads the source template DXF, extracts the Inverter-1.1
template band (Y 159 400 - 168 000) plus the two referenced block definitions,
and writes everything to template_data.json alongside this script.

Run once whenever the master DXF changes:
    python extract_template.py
"""

import ezdxf
import json
import os

DXF_PATH       = r'C:\Users\user\Desktop\SLD Diagram\Example\ETN-EPC-SUN-ELE-DRW-003 - LV Single Line Diagram.dxf'
OUTPUT_PATH    = os.path.join(os.path.dirname(__file__), 'template_data.json')
TEMPLATE_Y_MIN = 159_000
TEMPLATE_Y_MAX = 168_000
COL_STEP       = 11_737.3



# ─────────────────────────────────────────────────────────────────────────────

def entity_y(e):
    try:
        t = e.dxftype()
        if t in ('TEXT', 'MTEXT', 'INSERT'):  return e.dxf.insert.y
        if t in ('ARC', 'CIRCLE', 'ELLIPSE'): return e.dxf.center.y
        if t == 'LINE':                        return e.dxf.start.y
        if t == 'LWPOLYLINE':
            pts = list(e.get_points())
            return pts[0][1] if pts else None
        if t == 'POLYLINE':
            vs = list(e.vertices)
            return vs[0].dxf.location.y if vs else None
    except Exception:
        pass
    return None


def common_attrs(e):
    d = {}
    for a in ('layer', 'color', 'linetype', 'lineweight', 'ltscale'):
        try:
            if e.dxf.hasattr(a):
                d[a] = getattr(e.dxf, a)
        except Exception:
            pass
    return d


def extract_entity(e):
    t = e.dxftype()
    d = {'type': t}
    d.update(common_attrs(e))
    try:
        if t == 'LWPOLYLINE':
            d['pts']    = [list(p) for p in e.get_points()]
            d['closed'] = e.closed
            if e.dxf.hasattr('const_width'):
                d['const_width'] = e.dxf.const_width

        elif t == 'MTEXT':
            pos = e.dxf.insert
            d['x'] = pos.x
            d['y'] = pos.y
            d['text'] = e.text
            for a in ('char_height', 'width', 'attachment_point', 'flow_direction',
                      'line_spacing_style', 'line_spacing_factor', 'style'):
                try:
                    if e.dxf.hasattr(a):
                        d[a] = getattr(e.dxf, a)
                except Exception:
                    pass

        elif t == 'ARC':
            c = e.dxf.center
            d.update({'cx': c.x, 'cy': c.y, 'cz': c.z,
                      'radius': e.dxf.radius,
                      'start_angle': e.dxf.start_angle,
                      'end_angle': e.dxf.end_angle})

        elif t == 'CIRCLE':
            c = e.dxf.center
            d.update({'cx': c.x, 'cy': c.y, 'cz': c.z, 'radius': e.dxf.radius})

        elif t == 'LINE':
            s, en = e.dxf.start, e.dxf.end
            d.update({'sx': s.x, 'sy': s.y, 'sz': s.z,
                      'ex': en.x, 'ey': en.y, 'ez': en.z})

        elif t == 'ELLIPSE':
            c, ma = e.dxf.center, e.dxf.major_axis
            d.update({'cx': c.x, 'cy': c.y, 'cz': c.z,
                      'major_axis': [ma.x, ma.y, ma.z],
                      'ratio': e.dxf.ratio,
                      'start_param': e.dxf.start_param,
                      'end_param': e.dxf.end_param})

        elif t == 'INSERT':
            ins = e.dxf.insert
            d.update({'name': e.dxf.name,
                      'ix': ins.x, 'iy': ins.y,
                      'iz': ins.z if hasattr(ins, 'z') else 0.0})
            for a in ('xscale', 'yscale', 'rotation'):
                try:
                    if e.dxf.hasattr(a):
                        d[a] = getattr(e.dxf, a)
                except Exception:
                    pass

        elif t == 'POLYLINE':
            d['pts3d'] = [[v.dxf.location.x, v.dxf.location.y, v.dxf.location.z]
                          for v in e.vertices]

        else:
            return None

    except Exception as ex:
        print(f"  [warn] extract {t}: {ex}")
        return None

    return d


def extract_block(doc, name):
    try:
        blk = doc.blocks[name]
        entities = []
        for e in blk:
            d = extract_entity(e)
            if d is not None:
                entities.append(d)
        return entities
    except Exception as ex:
        print(f"  [warn] block '{name}': {ex}")
        return []


# ─────────────────────────────────────────────────────────────────────────────

def main():
    if not os.path.isfile(DXF_PATH):
        print(f"ERROR: DXF file not found at:")
        print(f"  {DXF_PATH}")
        print(f"\nMake sure the file exists and the path is correct.")
        return False
    
    print(f"Loading {DXF_PATH} …")
    try:
        doc = ezdxf.readfile(DXF_PATH)
    except Exception as ex:
        print(f"ERROR: Failed to open DXF file:")
        print(f"  {ex}")
        print(f"\nThe file may be corrupted or in an unsupported DXF version.")
        return False
    
    msp = doc.modelspace()

    # Collect template-band entities
    tmpl_raw = [e for e in msp
                if (y := entity_y(e)) is not None
                and TEMPLATE_Y_MIN <= y <= TEMPLATE_Y_MAX
                and (not e.dxf.hasattr('layer') or e.dxf.layer.lower() != 'defpoints')]
    print(f"Template-band entities : {len(tmpl_raw)}")

    # Extract to dicts
    entities = []
    skipped  = 0
    for e in tmpl_raw:
        d = extract_entity(e)
        if d is not None:
            entities.append(d)
        else:
            skipped += 1
    print(f"Extracted {len(entities)} entities  ({skipped} skipped/unsupported)")

    # Find referenced block names
    block_names = sorted(set(d['name'] for d in entities if d['type'] == 'INSERT'))
    print(f"Referenced blocks: {block_names}")

    # Extract block definitions
    blocks = {}
    for name in block_names:
        blk_entities = extract_block(doc, name)
        blocks[name] = blk_entities
        print(f"  Block '{name}': {len(blk_entities)} entities")

    # Extract custom linetype definitions used by template entities
    STANDARD_LT = {'bylayer', 'byblock', 'continuous', 'continua'}
    used_lt = set()
    for e in entities + [be for bl in blocks.values() for be in bl]:
        lt = e.get('linetype', '')
        if lt and lt.lower() not in STANDARD_LT:
            used_lt.add(lt)

    linetypes = {}
    for lt_name in sorted(used_lt):
        try:
            lt = doc.linetypes.get(lt_name)
            raw = lt.simplified_line_pattern()
            # Convert (dash, gap, ...) to ezdxf pattern [+dash, -gap, ...]
            pattern = []
            for i, v in enumerate(raw):
                if abs(v) < 1e-10:
                    pattern.append(0.0)
                elif i % 2 == 0:
                    pattern.append(abs(v))
                else:
                    pattern.append(-abs(v))
            linetypes[lt_name] = {
                'pattern':     pattern,
                'description': lt.dxf.get('description', ''),
            }
            print(f"  Linetype '{lt_name}': {pattern}")
        except Exception as ex:
            print(f"  [warn] linetype '{lt_name}': {ex}")

    # Compute template geometry bounds for reference
    xs = []
    for d in entities:
        t = d['type']
        try:
            if t == 'LWPOLYLINE': xs += [p[0] for p in d['pts']]
            elif t == 'MTEXT':    xs.append(d['x'])
            elif t in ('ARC', 'CIRCLE', 'ELLIPSE'): xs.append(d['cx'])
            elif t == 'LINE':     xs += [d['sx'], d['ex']]
            elif t == 'INSERT':   xs.append(d['ix'])
            elif t == 'POLYLINE': xs += [p[0] for p in d['pts3d']]
        except Exception:
            pass
    xmin = min(xs) if xs else 0

    payload = {
        'meta': {
            'source':          os.path.basename(DXF_PATH),
            'template_y_min':  TEMPLATE_Y_MIN,
            'template_y_max':  TEMPLATE_Y_MAX,
            'col_step':        COL_STEP,
            'row_step':        10_200,
            'xmin':            xmin,
        },
        'entities':  entities,
        'blocks':    blocks,
        'linetypes': linetypes,
    }

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(payload, f, separators=(',', ':'))

    size_kb = os.path.getsize(OUTPUT_PATH) / 1024
    print(f"\nWritten: {OUTPUT_PATH}  ({size_kb:.1f} KB)")
    print("Done.")
    return True


if __name__ == '__main__':
    main()
