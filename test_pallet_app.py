import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import datetime, tempfile, os, math

# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY        = "1E4080"
NAVY_MID    = "2B5FAD"
NAVY_LIGHT  = "4A7FC1"
NAVY_HEADER = "1A3A6C"
LOGO_BG     = "C8DEFF"   # light periwinkle — logo-friendly, readable
ORANGE      = "E8722A"
ORANGE_DARK = "C45A1A"
ORANGE_SOFT = "FDF0E8"   # very soft peach for alternating rows
CREAM       = "FFFBF5"
WHITE       = "FFFFFF"
LIGHT_BLUE  = "EBF3FC"
TABLE_ALT   = "F4F8FD"
SILVER      = "BDC8D8"
DARK_TEXT   = "1A2744"
MID_TEXT    = "4A5678"
TEAL        = "1A6B7C"   # accent for accessories section
TEAL_LIGHT  = "E0F4F7"
TEAL_DARK   = "154F5C"

# ── BORDER HELPERS ────────────────────────────────────────────────────────────
def side(color, style="thin"):
    return Side(style=style, color=color)

def b_thin(c=SILVER):
    s = side(c)
    return Border(left=s, right=s, top=s, bottom=s)

def b_med(c=NAVY_MID):
    s = side(c, "medium")
    return Border(left=s, right=s, top=s, bottom=s)

# ── CELL WRITER ───────────────────────────────────────────────────────────────
def W(ws, row, col, value=None, bold=False, sz=10, color=DARK_TEXT,
      bg=None, ha="left", va="center", wrap=False, fmt=None,
      bdr=None, italic=False, ind=0):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.font = Font(name="Calibri", size=sz, bold=bold, color=color, italic=italic)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap, indent=ind)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        c.number_format = fmt
    if bdr:
        c.border = bdr
    return c

def fill(ws, row, c1, c2, bg, bdr=None):
    for col in range(c1, c2 + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=bg)
        if bdr:
            ws.cell(row=row, column=col).border = bdr

def mg(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def set_row_h(ws, r, h):
    ws.row_dimensions[r].height = h

# ── PHYSICS — FORMULAS ───────────────────────────────────────────────────────
def upright_sheet_size(width, depth):
    return width + depth + width + 15

def pipe_beam_sheet_size(height, width):
    return 2 * (height + width)

def roll_form_beam_sheet_size(height, width):
    return ((4 * width) + (2 * height)) + 20

def deep_bar_size(rack_depth):
    return rack_depth - 65

def cross_bar_length(deep_bar, gap, upright_depth):
    a = deep_bar - 50
    b = gap
    return ((a ** 2 + b ** 2) ** 0.5) + 50

def weight(length, width, thickness):
    return (length * width * thickness * 7.85) / 1_000_000

def calc_components(rack):
    uw, ud   = rack["uw"], rack["ud"]
    ul, ut   = rack["ul"], rack["ut"]
    bh, bw   = rack["bh"], rack["bw"]
    bl, bth  = rack["bl"], rack["bth"]
    depth    = rack["depth"]
    dth      = rack["dth"]
    gap      = rack["gap"]
    cth      = rack["cth"]
    levels   = rack["levels"]
    method   = rack["method"]
    main_qty = rack["main_qty"]
    addon_qty= rack["addon_qty"]

    # --- Upright ---
    uwid  = upright_sheet_size(uw, ud)
    u_wt  = weight(ul, uwid, ut)
    u_main_qty  = 4 * main_qty
    u_addon_qty = 2 * addon_qty
    u_main  = u_wt * u_main_qty
    u_addon = u_wt * u_addon_qty

    # --- Beam ---
    if rack["bt"] == "Pipe Beam":
        bwid = pipe_beam_sheet_size(bh, bw)
    else:
        bwid = roll_form_beam_sheet_size(bh, bw)
    b_wt = weight(bl, bwid, bth) + 1.15

    beam_per_rack = 2 * levels
    b_main_qty  = beam_per_rack * main_qty
    b_addon_qty = beam_per_rack * addon_qty
    b_main  = b_wt * b_main_qty
    b_addon = b_wt * b_addon_qty

    # --- Deep Bar ---
    dlen  = deep_bar_size(depth)
    d_wt  = weight(dlen, 92, dth)
    d_main_qty  = 4 * main_qty
    d_addon_qty = 2 * addon_qty
    d_main  = d_wt * d_main_qty
    d_addon = d_wt * d_addon_qty

    # --- Cross Brace ---
    eff_h     = ul - method
    num_cross = int(eff_h // gap)
    clen      = cross_bar_length(dlen, gap, ud)
    c_wt      = weight(clen, 92, cth)
    c_per_main  = num_cross * 2
    c_per_addon = num_cross
    c_main_qty  = c_per_main  * main_qty
    c_addon_qty = c_per_addon * addon_qty
    c_main  = c_wt * c_main_qty
    c_addon = c_wt * c_addon_qty

    total_main  = round(u_main  + b_main  + d_main  + c_main,  3)
    total_addon = round(u_addon + b_addon + d_addon + c_addon, 3)

    area_m2        = (bl / 1000) * (depth / 1000)
    udl_kg_m2      = 500
    load_per_level = round(udl_kg_m2 * area_m2, 1)

    return {
        "uwid": round(uwid, 1), "ul": ul, "ut": ut,
        "u_wt": round(u_wt, 4),
        "u_main_qty": u_main_qty, "u_addon_qty": u_addon_qty,
        "u_main": round(u_main, 3), "u_addon": round(u_addon, 3),
        "bwid": round(bwid, 1), "bl": bl, "bth": bth,
        "b_wt": round(b_wt, 4),
        "beam_per_rack": beam_per_rack,
        "b_main_qty": b_main_qty, "b_addon_qty": b_addon_qty,
        "b_main": round(b_main, 3), "b_addon": round(b_addon, 3),
        "dlen": round(dlen, 1), "dth": dth,
        "d_wt": round(d_wt, 4),
        "d_main_qty": d_main_qty, "d_addon_qty": d_addon_qty,
        "d_main": round(d_main, 3), "d_addon": round(d_addon, 3),
        "clen": round(clen, 1), "cth": cth, "num_cross": num_cross,
        "c_per_main": c_per_main, "c_per_addon": c_per_addon,
        "c_wt": round(c_wt, 4),
        "c_main_qty": c_main_qty, "c_addon_qty": c_addon_qty,
        "c_main": round(c_main, 3), "c_addon": round(c_addon, 3),
        "total_main": total_main, "total_addon": total_addon,
        "area_m2": round(area_m2, 3),
        "udl_kg_m2": udl_kg_m2,
        "load_per_level": load_per_level,
    }


def calc_accessories(acc_data, rack_data):
    """Calculate accessory weights and totals. Returns list of dicts."""
    items = []

    total_uprights = sum(4 * r["main_qty"] + 2 * r["addon_qty"] for r in rack_data)

    # Column Guard
    cg_qty = acc_data.get("cg_qty", 0)
    cg_wt  = 3.75
    items.append({"name": "Column Guard", "spec": "Standard",
                  "qty": cg_qty, "wt_each": cg_wt,
                  "total_wt": round(cg_qty * cg_wt, 2)})

    # Row Connector
    rc_qty = acc_data.get("rc_qty", 0)
    rc_wt  = 1.0
    items.append({"name": "Row Connector", "spec": "Standard",
                  "qty": rc_qty, "wt_each": rc_wt,
                  "total_wt": round(rc_qty * rc_wt, 2)})

    # Row Guards (multiple types)
    for rg in acc_data.get("row_guards", []):
        h, l, qty = rg["h"], rg["l"], rg["qty"]
        wt = (((240 * h * 2 * 7.85) + (240 * l * 2 * 7.85)) * 2) / 1_000_000
        items.append({"name": "Row Guard", "spec": f"{h}×{l}",
                      "qty": qty, "wt_each": round(wt, 4),
                      "total_wt": round(wt * qty, 2)})

    # Tie Beams (multiple types)
    for tb in acc_data.get("tie_beams", []):
        qty, w, d, l, t = tb["qty"], tb["w"], tb["d"], tb["l"], tb["t"]
        wt = weight(l, upright_sheet_size(w, d), t)
        items.append({"name": "Tie Beam", "spec": f"{w}×{d}",
                      "qty": qty, "wt_each": round(wt, 4),
                      "total_wt": round(wt * qty, 2)})

    # Back Pallet Stoppers (multiple types)
    for bps in acc_data.get("bps_list", []):
        qty, l = bps["qty"], bps["l"]
        wt = ((160 * 1.6 * l * 7.85) / 1_000_000) + 0.6
        items.append({"name": "Back Pallet Stopper", "spec": f"L={l}",
                      "qty": qty, "wt_each": round(wt, 4),
                      "total_wt": round(wt * qty, 2)})

    # Base Plate (auto from uprights)
    items.append({"name": "Base Plate", "spec": "Auto",
                  "qty": total_uprights, "wt_each": "Included",
                  "total_wt": "-"})

    return items


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — COMMERCIAL OFFER
# ═══════════════════════════════════════════════════════════════════════════════
def build_quotation_sheet(ws, client, product, offer_no, date_obj,
                          project_name, rack_data, rate_per_kg,
                          acc_data=None, logo_path=None):
    ws.sheet_view.showGridLines = False

    col_w = {1:1.5, 2:5, 3:6, 4:28, 5:14, 6:9, 7:16, 8:17, 9:1.5}
    for col, w in col_w.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    R = 1

    # Top accent bar
    set_row_h(ws, R, 5); fill(ws, R, 1, 9, ORANGE); R += 1

    # ── LOGO + COMPANY NAME ROW — light background so logo shows clearly ──
    set_row_h(ws, R, 68); fill(ws, R, 1, 9, LOGO_BG)
    if logo_path and os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 200; img.height = 62
            img.anchor = f"B{R}"
            ws.add_image(img)
        except Exception:
            pass
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES"
    c.font = Font(name="Calibri", size=22, bold=True, color=NAVY_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=LOGO_BG)
    R += 1

    set_row_h(ws, R, 16); fill(ws, R, 1, 9, NAVY_MID)
    mg(ws, R, 2, R, 8)
    W(ws, R, 2, "DSS Dolphin Storage Solutions  ·  Modular Mezzanine & Racking Systems",
      sz=9, color="C8DEFF", bg=NAVY_MID, ha="center")
    R += 1

    set_row_h(ws, R, 14); fill(ws, R, 1, 9, NAVY_LIGHT)
    mg(ws, R, 2, R, 8)
    W(ws, R, 2, "86/3/1 Road No 7, Mundka Industrial Area, New Delhi – 110041   |   GST: 07AAMFB6403G1ZM   |   +91 9625589161 / 9811096149",
      sz=8, color=WHITE, bg=NAVY_LIGHT, ha="center")
    R += 1

    set_row_h(ws, R, 13); fill(ws, R, 1, 9, NAVY_LIGHT)
    mg(ws, R, 2, R, 8)
    W(ws, R, 2, "brijindustries09@rediffmail.com   |   www.brijindustries.in",
      sz=8, color="D8EEFF", bg=NAVY_LIGHT, ha="center", italic=True)
    R += 1

    set_row_h(ws, R, 5); fill(ws, R, 1, 9, ORANGE); R += 1
    set_row_h(ws, R, 6); R += 1

    # Offer title
    set_row_h(ws, R, 24)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "COMMERCIAL OFFER"
    c.font = Font(name="Calibri", size=13, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    c.border = b_med(NAVY_HEADER)
    R += 1

    def detail_pair(lbl_l, val_l, lbl_r, val_r, bg=LIGHT_BLUE):
        nonlocal R
        set_row_h(ws, R, 20)
        fill(ws, R, 2, 8, bg, b_thin())
        mg(ws, R, 2, R, 3)
        W(ws, R, 2, lbl_l, bold=True, sz=9, color=NAVY_MID, bg=bg, ha="right", bdr=b_thin())
        mg(ws, R, 4, R, 5)
        W(ws, R, 4, val_l, sz=9, color=DARK_TEXT, bg=bg, ind=1, bdr=b_thin())
        W(ws, R, 6, lbl_r, bold=True, sz=9, color=NAVY_MID, bg=bg, ha="right", bdr=b_thin())
        mg(ws, R, 7, R, 8)
        W(ws, R, 7, val_r, sz=9, color=DARK_TEXT, bg=bg, ind=1, bdr=b_thin())
        R += 1

    detail_pair("To :",      client,       "Date :",      date_obj.strftime("%d %B %Y"), LIGHT_BLUE)
    detail_pair("Product :", product,      "Offer No. :", offer_no,                      WHITE)
    detail_pair("Project :", project_name, "",            "",                             LIGHT_BLUE)

    set_row_h(ws, R, 6); R += 1

    # Technical Details section
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "TECHNICAL DETAILS"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_MID)
    c.border = b_med(NAVY_MID)
    R += 1

    set_row_h(ws, R, 22)
    tech_hdrs = [(2,"MODULE","center"),(3,"UPRIGHT\nHEIGHT (mm)","center"),
                 (4,"BEAM\nLENGTH (mm)","center"),(5,"RACK\nDEPTH (mm)","center"),
                 (6,"LEVELS","center"),(7,"LOAD / LEVEL\n(kg)","center"),
                 (8,"PALLETS\n/ LEVEL","center")]
    for col, txt, al in tech_hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        c.border = b_thin()
    R += 1

    for idx, rack in enumerate(rack_data):
        comp = calc_components(rack)
        bg = WHITE if idx % 2 == 0 else TABLE_ALT
        set_row_h(ws, R, 18)
        fill(ws, R, 2, 8, bg, b_thin())
        W(ws, R, 2, f"Module {rack['module']}", bold=True, sz=9, color=NAVY_MID, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 3, rack["ul"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 4, rack["bl"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 5, rack["depth"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 6, rack["levels"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 7, comp["load_per_level"], sz=9, bg=bg, ha="center", bdr=b_thin(), fmt="#,##0")
        pallets = max(1, int(rack["bl"] / 1200))
        W(ws, R, 8, pallets, sz=9, bg=bg, ha="center", bdr=b_thin())
        R += 1

    set_row_h(ws, R, 6); R += 1

    # Scope of Supply
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "SCOPE OF SUPPLY"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=ORANGE)
    c.border = b_med(ORANGE_DARK)
    R += 1

    set_row_h(ws, R, 26); fill(ws, R, 2, 8, NAVY_HEADER)
    scope_hdrs = [(2,"SR.","center"),(3,"","center"),(4,"DESCRIPTION","center"),
                  (5,"TYPE","center"),(6,"QTY","center"),
                  (7,"UNIT PRICE (₹)","center"),(8,"AMOUNT (₹)","center")]
    for col, txt, al in scope_hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        c.alignment = Alignment(horizontal=al, vertical="center")
        c.border = b_thin()
    R += 1

    total_basic = 0.0
    sr = 1

    for rack in rack_data:
        comp = calc_components(rack)
        main_price  = comp["total_main"]  * rate_per_kg
        addon_price = comp["total_addon"] * rate_per_kg
        main_total  = main_price * rack["main_qty"]
        total_basic += main_total

        bg = WHITE if sr % 2 == 1 else TABLE_ALT
        set_row_h(ws, R, 20); fill(ws, R, 2, 8, bg, b_thin())
        W(ws, R, 2, sr, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
        mg(ws, R, 3, R, 4)
        W(ws, R, 3, f"  MODULE {rack['module']}", bold=True, sz=9, color=DARK_TEXT, bg=bg, bdr=b_thin())
        W(ws, R, 5, "Main Rack", sz=9, color=NAVY_MID, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 6, rack["main_qty"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 7, round(main_price, 2), sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        W(ws, R, 8, round(main_total, 2), bold=True, sz=9, color=NAVY, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        R += 1; sr += 1

        if rack["addon_qty"] > 0:
            addon_total  = addon_price * rack["addon_qty"]
            total_basic += addon_total
            bg = WHITE if sr % 2 == 1 else TABLE_ALT
            set_row_h(ws, R, 20); fill(ws, R, 2, 8, bg, b_thin())
            W(ws, R, 2, sr, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
            mg(ws, R, 3, R, 4)
            W(ws, R, 3, f"  MODULE {rack['module']}", sz=9, color=MID_TEXT, bg=bg, bdr=b_thin())
            W(ws, R, 5, "Add-on Rack", sz=9, color=ORANGE, bg=bg, ha="center", bdr=b_thin())
            W(ws, R, 6, rack["addon_qty"], sz=9, bg=bg, ha="center", bdr=b_thin())
            W(ws, R, 7, round(addon_price, 2), sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
            W(ws, R, 8, round(addon_total, 2), bold=True, sz=9, color=NAVY, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
            R += 1; sr += 1

    # ── ACCESSORIES ROWS ─────────────────────────────────────────────────────
    if acc_data:
        acc_items = calc_accessories(acc_data, rack_data)
        has_priced = any(isinstance(a["total_wt"], (int, float)) for a in acc_items)
        if has_priced:
            # Accessories sub-header
            set_row_h(ws, R, 20); fill(ws, R, 2, 8, TEAL)
            mg(ws, R, 2, R, 8)
            c = ws.cell(row=R, column=2)
            c.value = "  ACCESSORIES"
            c.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.fill = PatternFill("solid", fgColor=TEAL)
            c.border = b_med(TEAL_DARK)
            R += 1

            for acc in acc_items:
                bg = TEAL_LIGHT if sr % 2 == 1 else WHITE
                set_row_h(ws, R, 18); fill(ws, R, 2, 8, bg, b_thin())
                W(ws, R, 2, sr, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
                mg(ws, R, 3, R, 4)
                W(ws, R, 3, f"  {acc['name']}", sz=9, color=TEAL_DARK, bg=bg, bdr=b_thin())
                W(ws, R, 5, acc["spec"], sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
                W(ws, R, 6, acc["qty"], sz=9, bg=bg, ha="center", bdr=b_thin())

                wt_each = acc["wt_each"]
                if isinstance(wt_each, (int, float)):
                    W(ws, R, 7, round(wt_each, 4), sz=9, bg=bg, ha="right", fmt="#,##0.000", bdr=b_thin())
                else:
                    W(ws, R, 7, str(wt_each), sz=9, bg=bg, ha="center", bdr=b_thin())

                tot = acc["total_wt"]
                if isinstance(tot, (int, float)):
                    W(ws, R, 8, f"—  {tot:.2f} kg", sz=9, color=TEAL_DARK, bg=bg, ha="right", bdr=b_thin())
                else:
                    W(ws, R, 8, str(tot), sz=9, color=TEAL_DARK, bg=bg, ha="center", bdr=b_thin())
                R += 1; sr += 1

    # Subtotal
    set_row_h(ws, R, 22); fill(ws, R, 2, 8, LIGHT_BLUE)
    mg(ws, R, 2, R, 7)
    c = ws.cell(row=R, column=2)
    c.value = "SUBTOTAL (BASIC AMOUNT)"
    c.font = Font(name="Calibri", size=10, bold=True, color=NAVY_HEADER)
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    c.border = b_med(NAVY_MID)
    W(ws, R, 8, round(total_basic, 2), bold=True, sz=10, color=NAVY_HEADER,
      bg=LIGHT_BLUE, ha="right", fmt="#,##0.00", bdr=b_med(NAVY_MID))
    R += 1
    set_row_h(ws, R, 6); R += 1

    # Pricing summary
    gst   = round(total_basic * 0.18, 2)
    grand = round(total_basic + gst, 2)

    def price_row(lbl, val, bg=WHITE, bold=False, vc=DARK_TEXT, is_text=False):
        nonlocal R
        set_row_h(ws, R, 20); fill(ws, R, 5, 8, bg)
        mg(ws, R, 5, R, 7)
        c_l = ws.cell(row=R, column=5)
        c_l.value = lbl
        c_l.font = Font(name="Calibri", size=10, bold=bold, color=vc)
        c_l.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        c_l.fill = PatternFill("solid", fgColor=bg)
        c_l.border = b_thin()
        c_v = ws.cell(row=R, column=8)
        c_v.value = val
        c_v.font = Font(name="Calibri", size=10, bold=bold, color=vc, italic=is_text)
        c_v.alignment = Alignment(horizontal="center" if is_text else "right", vertical="center", indent=1)
        c_v.fill = PatternFill("solid", fgColor=bg)
        c_v.border = b_thin()
        if not is_text:
            c_v.number_format = "#,##0.00"
        R += 1

    price_row("Basic Amount (₹)",    round(total_basic, 2), WHITE, True,  NAVY_HEADER)
    price_row("Freight Charges",     "Inclusive",           ORANGE_SOFT, False, ORANGE, True)
    price_row("Erection Charges",    "Inclusive",           ORANGE_SOFT, False, ORANGE, True)
    price_row("GST @ 18% (₹)",       gst,                   WHITE, False, MID_TEXT)

    # Grand Total
    set_row_h(ws, R, 28); fill(ws, R, 2, 8, ORANGE)
    mg(ws, R, 2, R, 7)
    c = ws.cell(row=R, column=2)
    c.value = "GRAND TOTAL  (Inclusive of GST @ 18%)"
    c.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=ORANGE)
    c.border = b_med(ORANGE_DARK)
    c_gv = ws.cell(row=R, column=8)
    c_gv.value = grand
    c_gv.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c_gv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    c_gv.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    c_gv.border = b_med(NAVY_HEADER)
    c_gv.number_format = '₹ #,##0.00'
    R += 1
    set_row_h(ws, R, 8); R += 1

    # Terms & Bank
    set_row_h(ws, R, 20); fill(ws, R, 2, 8, NAVY_MID)
    mg(ws, R, 2, R, 5)
    c = ws.cell(row=R, column=2)
    c.value = "TERMS & CONDITIONS"
    c.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_MID)
    mg(ws, R, 6, R, 8)
    c2 = ws.cell(row=R, column=6)
    c2.value = "BANK DETAILS"
    c2.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill("solid", fgColor=NAVY_MID)
    R += 1

    terms = [
        "Payment: 50% advance, balance against delivery",
        "Delivery: 10–12 weeks from advance receipt",
        "Warranty: 12 months from commissioning date",
        "GST @ 18% applicable as per Government norms",
        "Prices valid for 4–5 days from date of offer",
    ]
    bank = [
        "Account Name : BRIJ INDUSTRIES",
        "Bank           : ICICI Bank Ltd.",
        "Account No.  : 135805001108",
        "IFSC Code     : ICIC0001358",
        "Branch         : Mundka, New Delhi",
    ]
    for t, b in zip(terms, bank):
        set_row_h(ws, R, 17)
        mg(ws, R, 2, R, 5)
        c = ws.cell(row=R, column=2)
        c.value = f"  • {t}"
        c.font = Font(name="Calibri", size=8, color=MID_TEXT)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        c.border = b_thin()
        mg(ws, R, 6, R, 8)
        c2 = ws.cell(row=R, column=6)
        c2.value = f"  {b}"
        c2.font = Font(name="Calibri", size=8, color=MID_TEXT)
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.fill = PatternFill("solid", fgColor=ORANGE_SOFT)
        c2.border = b_thin()
        R += 1

    set_row_h(ws, R, 6); R += 1

    # Signature
    set_row_h(ws, R, 42)
    mg(ws, R, 2, R, 5)
    c = ws.cell(row=R, column=2)
    c.value = "Customer Signature & Stamp"
    c.font = Font(name="Calibri", size=8, color=SILVER, italic=True)
    c.alignment = Alignment(horizontal="center", vertical="bottom")
    c.border = Border(top=Side(style="medium", color=NAVY_MID))
    mg(ws, R, 6, R, 8)
    c2 = ws.cell(row=R, column=6)
    c2.value = "For BRIJ INDUSTRIES"
    c2.font = Font(name="Calibri", size=8, color=SILVER, italic=True)
    c2.alignment = Alignment(horizontal="center", vertical="bottom")
    c2.border = Border(top=Side(style="medium", color=NAVY_MID))
    R += 1

    set_row_h(ws, R, 5); fill(ws, R, 1, 9, ORANGE); R += 1
    set_row_h(ws, R, 14); fill(ws, R, 1, 9, NAVY_HEADER)
    mg(ws, R, 2, R, 8)
    W(ws, R, 2, "Thank you for considering DSS Dolphin Storage Solutions. We look forward to serving you.",
      sz=8, color="8AADDD", bg=NAVY_HEADER, ha="center", italic=True)

    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.print_area = f"A1:{get_column_letter(9)}{R}"

    return total_basic, gst, grand


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 2 — BILL OF MATERIALS
# ═══════════════════════════════════════════════════════════════════════════════
def build_bom_sheet(ws, client, offer_no, date_obj, rack_data, acc_data=None):
    ws.sheet_view.showGridLines = False

    col_w = {1:1.5, 2:4, 3:22, 4:13, 5:11, 6:11,
             7:12, 8:11, 9:11, 10:14, 11:14, 12:15, 13:1.5}
    for col, w in col_w.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    R = 1
    set_row_h(ws, R, 5); fill(ws, R, 1, 13, ORANGE); R += 1

    # Header — same light blue logo-friendly background
    set_row_h(ws, R, 52); fill(ws, R, 1, 13, LOGO_BG)
    mg(ws, R, 2, R, 12)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES  —  BILL OF MATERIALS"
    c.font = Font(name="Calibri", size=17, bold=True, color=NAVY_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=LOGO_BG)
    R += 1

    set_row_h(ws, R, 14); fill(ws, R, 1, 13, NAVY_LIGHT)
    mg(ws, R, 2, R, 12)
    W(ws, R, 2,
      f"DSS Dolphin Storage Solutions  |  Offer No: {offer_no}  |  Customer: {client}  |  Date: {date_obj.strftime('%d %B %Y')}",
      sz=9, color=WHITE, bg=NAVY_LIGHT, ha="center")
    R += 1

    set_row_h(ws, R, 5); fill(ws, R, 1, 13, ORANGE); R += 1
    set_row_h(ws, R, 6); R += 1

    grand_main_wt  = 0.0
    grand_addon_wt = 0.0

    for rack in rack_data:
        comp = calc_components(rack)

        set_row_h(ws, R, 22)
        mg(ws, R, 2, R, 12)
        c = ws.cell(row=R, column=2)
        c.value = (f"  MODULE {rack['module']}   |   Main Racks: {rack['main_qty']}"
                   f"   |   Add-on Racks: {rack['addon_qty']}   |   Levels: {rack['levels']}"
                   f"   |   Cross Method: {rack['method']} mm   |   Gap: {rack['gap']} mm")
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.fill = PatternFill("solid", fgColor=NAVY_MID)
        c.border = b_med(NAVY_MID)
        R += 1

        set_row_h(ws, R, 34); fill(ws, R, 2, 12, NAVY_HEADER)
        BOM_H = [
            (2,  "SR.",                  "center"),
            (3,  "COMPONENT",            "left"),
            (4,  "SECTION\nPROFILE",     "center"),
            (5,  "LENGTH\n(mm)",         "center"),
            (6,  "THICK\n(mm)",          "center"),
            (7,  "WT / PCS\n(kg)",       "center"),
            (8,  "QTY /\nMAIN",          "center"),
            (9,  "QTY /\nADD-ON",        "center"),
            (10, "MAIN\nTOTAL (kg)",     "center"),
            (11, "ADD-ON\nTOTAL (kg)",   "center"),
            (12, "LOAD /\nLEVEL (kg)",   "center"),
        ]
        for col, txt, al in BOM_H:
            c = ws.cell(row=R, column=col)
            c.value = txt
            c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
            c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
            c.border = b_thin()
        R += 1

        def comp_row(sr_n, comp_name, section, length, thick, wt_each,
                     qty_main, qty_addon, load_level_val, bg):
            nonlocal R
            set_row_h(ws, R, 19)
            fill(ws, R, 2, 12, bg, b_thin())
            main_tot  = round(wt_each * qty_main,  3)
            addon_tot = round(wt_each * qty_addon, 3)
            row_vals = [
                (2,  sr_n,             "center", None),
                (3,  f"  {comp_name}", "left",   None),
                (4,  section,          "center", None),
                (5,  length,           "center", None),
                (6,  thick,            "center", None),
                (7,  wt_each,          "center", "#,##0.000"),
                (8,  qty_main,         "center", None),
                (9,  qty_addon,        "center", None),
                (10, main_tot,         "right",  "#,##0.000"),
                (11, addon_tot,        "right",  "#,##0.000"),
                (12, load_level_val,   "center", "#,##0"),
            ]
            for col, val, al, nf in row_vals:
                c = ws.cell(row=R, column=col)
                c.value = val
                c.font = Font(name="Calibri", size=8.5, color=DARK_TEXT)
                c.alignment = Alignment(horizontal=al, vertical="center")
                c.fill = PatternFill("solid", fgColor=bg)
                c.border = b_thin()
                if nf:
                    c.number_format = nf
            R += 1

        if rack["bt"] == "Pipe Beam":
            beam_sec = f"{rack['bh']}×{rack['bw']} Pipe"
        else:
            beam_sec = f"{rack['bh']}×{rack['bw']} Roll"

        comp_row(1, "Upright / Column",
                 f"{rack['uw']}×{rack['ud']} Box",
                 rack["ul"], rack["ut"], comp["u_wt"],
                 comp["u_main_qty"], comp["u_addon_qty"],
                 "—", WHITE)

        comp_row(2, f"Beam  [2 × {rack['levels']} levels = {comp['beam_per_rack']} per rack]",
                 beam_sec,
                 rack["bl"], rack["bth"], comp["b_wt"],
                 comp["b_main_qty"], comp["b_addon_qty"],
                 comp["load_per_level"], TABLE_ALT)

        comp_row(3, "Deep Bar (Shelf Support)",
                 "92 mm flat",
                 int(comp["dlen"]), rack["dth"], comp["d_wt"],
                 comp["d_main_qty"], comp["d_addon_qty"],
                 "—", WHITE)

        comp_row(4, f"Cross Brace  [{comp['num_cross']} crosses × 2 per main, × 1 per addon]",
                 "92 mm flat",
                 int(comp["clen"]), rack["cth"], comp["c_wt"],
                 comp["c_main_qty"], comp["c_addon_qty"],
                 "—", TABLE_ALT)

        total_m_wt = round(comp["total_main"]  * rack["main_qty"],  2)
        total_a_wt = round(comp["total_addon"] * rack["addon_qty"], 2)
        grand_main_wt  += total_m_wt
        grand_addon_wt += total_a_wt

        def mod_summary(lbl, v_main, v_addon, bg, bold=False, vc=DARK_TEXT, nf="#,##0.00"):
            nonlocal R
            set_row_h(ws, R, 20); fill(ws, R, 2, 12, bg)
            mg(ws, R, 2, R, 9)
            c = ws.cell(row=R, column=2)
            c.value = lbl
            c.font = Font(name="Calibri", size=9, bold=bold, color=vc)
            c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            c.fill = PatternFill("solid", fgColor=bg)
            c.border = b_med(NAVY_MID) if bold else b_thin()
            for col, val in [(10, v_main), (11, v_addon)]:
                cv = ws.cell(row=R, column=col)
                cv.value = val
                cv.font = Font(name="Calibri", size=9, bold=bold, color=vc)
                cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
                cv.fill = PatternFill("solid", fgColor=bg)
                cv.border = b_med(NAVY_MID) if bold else b_thin()
                cv.number_format = nf
            R += 1

        mod_summary("Wt — Single Rack (kg)",
                    comp["total_main"], comp["total_addon"], LIGHT_BLUE)
        mod_summary(f"Total Wt — All Racks  [Main ×{rack['main_qty']}  |  Add-on ×{rack['addon_qty']}]",
                    total_m_wt, total_a_wt,
                    LIGHT_BLUE, bold=True, vc=NAVY_HEADER)

        set_row_h(ws, R, 10); R += 1

    # ── ACCESSORIES SECTION ───────────────────────────────────────────────────
    if acc_data:
        acc_items = calc_accessories(acc_data, rack_data)
        set_row_h(ws, R, 5); fill(ws, R, 1, 13, TEAL); R += 1
        set_row_h(ws, R, 22)
        mg(ws, R, 2, R, 12)
        c = ws.cell(row=R, column=2)
        c.value = "ACCESSORIES"
        c.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.border = b_med(TEAL_DARK)
        R += 1

        # Accessories header
        set_row_h(ws, R, 30); fill(ws, R, 2, 12, TEAL_DARK)
        ACC_H = [
            (2, "SR.", "center"),
            (3, "ITEM", "left"),
            (4, "SPECIFICATION", "center"),
            (5, "QTY", "center"),
            (6, "WT / PCS (kg)", "center"),
            (10, "TOTAL WT (kg)", "center"),
        ]
        for col in range(2, 13):
            c = ws.cell(row=R, column=col)
            c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=TEAL_DARK)
            c.border = b_thin()
        for col, txt, al in ACC_H:
            c = ws.cell(row=R, column=col)
            c.value = txt
            c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        R += 1

        for idx, acc in enumerate(acc_items):
            bg = TEAL_LIGHT if idx % 2 == 0 else WHITE
            set_row_h(ws, R, 19); fill(ws, R, 2, 12, bg, b_thin())
            W(ws, R, 2, idx + 1, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
            mg(ws, R, 3, R, 3)
            W(ws, R, 3, f"  {acc['name']}", sz=9, color=TEAL_DARK, bg=bg, bold=True, bdr=b_thin())
            mg(ws, R, 4, R, 4)
            W(ws, R, 4, acc["spec"], sz=9, color=DARK_TEXT, bg=bg, ha="center", bdr=b_thin())
            W(ws, R, 5, acc["qty"], sz=9, bg=bg, ha="center", bdr=b_thin())

            wt_each = acc["wt_each"]
            if isinstance(wt_each, (int, float)):
                W(ws, R, 6, round(wt_each, 4), sz=9, bg=bg, ha="right",
                  fmt="#,##0.000", bdr=b_thin())
            else:
                W(ws, R, 6, str(wt_each), sz=9, bg=bg, ha="center", bdr=b_thin())

            tot = acc["total_wt"]
            if isinstance(tot, (int, float)):
                W(ws, R, 10, round(tot, 2), sz=9, color=TEAL_DARK, bg=bg,
                  ha="right", bold=True, fmt="#,##0.00", bdr=b_thin())
            else:
                W(ws, R, 10, str(tot), sz=9, color=TEAL_DARK, bg=bg,
                  ha="center", bdr=b_thin())
            R += 1

        set_row_h(ws, R, 8); R += 1

    # Grand tonnage summary
    set_row_h(ws, R, 6); fill(ws, R, 1, 13, ORANGE); R += 1
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 12)
    c = ws.cell(row=R, column=2)
    c.value = "TOTAL TONNAGE SUMMARY"
    c.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    c.border = b_med(NAVY_HEADER)
    R += 1

    set_row_h(ws, R, 24); fill(ws, R, 2, 12, NAVY)
    ton_hdrs = [
        (2,  "MODULE",                     "left"),
        (6,  "ALL MAIN RACKS\nWT (kg)",    "center"),
        (8,  "ALL ADD-ON RACKS\nWT (kg)",  "center"),
        (10, "COMBINED\nWT (kg)",          "center"),
        (11, "COMBINED\nWT (MT)",          "center"),
        (12, "LOAD /\nLEVEL (kg)",         "center"),
    ]
    for col, txt, al in ton_hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        c.border = b_thin()
    R += 1

    for idx, rack in enumerate(rack_data):
        comp = calc_components(rack)
        m_wt = round(comp["total_main"]  * rack["main_qty"],  2)
        a_wt = round(comp["total_addon"] * rack["addon_qty"], 2)
        comb = round(m_wt + a_wt, 2)
        bg   = WHITE if idx % 2 == 0 else TABLE_ALT

        set_row_h(ws, R, 19); fill(ws, R, 2, 12, bg, b_thin())
        mg(ws, R, 2, R, 5)
        W(ws, R, 2, f"  MODULE {rack['module']}  (Main ×{rack['main_qty']} | Add-on ×{rack['addon_qty']})",
          bold=True, sz=9, color=NAVY_MID, bg=bg, bdr=b_thin())
        W(ws, R, 6,  m_wt,              sz=9, bg=bg, ha="right", fmt="#,##0.00",  bdr=b_thin())
        W(ws, R, 8,  a_wt,              sz=9, bg=bg, ha="right", fmt="#,##0.00",  bdr=b_thin())
        W(ws, R, 10, comb,              sz=9, bg=bg, ha="right", fmt="#,##0.00",  bdr=b_thin())
        W(ws, R, 11, round(comb/1000,3),sz=9, bg=bg, ha="right", fmt="#,##0.000", bdr=b_thin())
        W(ws, R, 12, comp["load_per_level"], sz=9, bg=bg, ha="center", fmt="#,##0", bdr=b_thin())
        R += 1

    grand_comb = round(grand_main_wt + grand_addon_wt, 2)
    set_row_h(ws, R, 22); fill(ws, R, 2, 12, ORANGE)
    mg(ws, R, 2, R, 5)
    c = ws.cell(row=R, column=2)
    c.value = "  GRAND TOTAL TONNAGE"
    c.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.fill = PatternFill("solid", fgColor=ORANGE)
    c.border = b_med(ORANGE_DARK)
    for col, val, nf in [
        (6,  round(grand_main_wt, 2),  "#,##0.00"),
        (8,  round(grand_addon_wt, 2), "#,##0.00"),
        (10, grand_comb,               "#,##0.00"),
        (11, round(grand_comb/1000,3), "#,##0.000"),
    ]:
        cv = ws.cell(row=R, column=col)
        cv.value = val
        cv.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        cv.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cv.border = b_med(NAVY_HEADER)
        cv.number_format = nf
    R += 1

    set_row_h(ws, R, 5); fill(ws, R, 1, 13, ORANGE); R += 1
    set_row_h(ws, R, 14); fill(ws, R, 1, 13, NAVY_HEADER)
    mg(ws, R, 2, R, 12)
    W(ws, R, 2, "BRIJ INDUSTRIES  |  DSS Dolphin Storage Solutions  |  www.brijindustries.in",
      sz=8, color="8AADDD", bg=NAVY_HEADER, ha="center", italic=True)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.print_area = f"A1:{get_column_letter(13)}{R}"


# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER BUILD
# ═══════════════════════════════════════════════════════════════════════════════
def build_excel(client, product, offer_no, date_obj, project_name,
                rack_data, rate_per_kg, acc_data=None,
                out_path="quotation.xlsx", logo_path=None):
    wb = Workbook()
    ws_q = wb.active
    ws_q.title = "Commercial Offer"
    ws_b = wb.create_sheet("Bill of Materials")

    total_basic, gst, grand = build_quotation_sheet(
        ws_q, client, product, offer_no, date_obj,
        project_name, rack_data, rate_per_kg, acc_data, logo_path
    )
    build_bom_sheet(ws_b, client, offer_no, date_obj, rack_data, acc_data)

    wb.save(out_path)
    return total_basic, gst, grand


# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="DSS Quotation Generator", layout="wide", page_icon="🐬")

st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #1A3A6C 0%, #2B5FAD 55%, #E8722A 100%);
    padding: 26px 36px; border-radius: 14px; color: white;
    text-align: center; margin-bottom: 24px;
    box-shadow: 0 8px 28px rgba(0,0,0,0.22);
}
.main-header h1 { margin:0; font-size:1.9rem; letter-spacing:2px; }
.main-header p  { margin:5px 0 0; font-size:0.92rem; opacity:0.82; }
.stButton > button {
    background: linear-gradient(135deg, #E8722A, #C45A1A);
    color: white; font-weight: bold; border: none;
    border-radius: 10px; padding: 13px 30px;
    font-size: 1.05rem; transition: all 0.2s ease;
}
.stButton > button:hover { transform:translateY(-2px); box-shadow:0 6px 22px rgba(232,114,42,.45); }
.acc-section { background: #E0F4F7; border-left: 4px solid #1A6B7C;
               padding: 10px 14px; border-radius: 6px; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🐬  QUOTATION GENERATOR</h1>
    <p>BRIJ INDUSTRIES — DSS Dolphin Storage Solutions  |  Modular Mezzanine & Racking Systems</p>
</div>
""", unsafe_allow_html=True)

with st.expander("🖼️ Upload Company Logo (optional)", expanded=False):
    logo_file = st.file_uploader("Upload logo PNG/JPG", type=["png","jpg","jpeg"])

logo_path = None
if logo_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
        tf.write(logo_file.read())
        logo_path = tf.name
    st.success(f"✅ Logo uploaded: {logo_file.name}")

# Customer Details
st.subheader("📋 Customer & Offer Details")
c1, c2, c3 = st.columns(3)
with c1:
    client       = st.text_input("Customer Name (M/S)",  "STYLE BAZAAR")
    product      = st.text_input("Product",              "MODULAR MEZZANINE FLOOR")
with c2:
    offer_no     = st.text_input("Offer No",             "DSS-IV/25-26/0712")
    project_name = st.text_input("Project Name",         "MODULE MEZZANINE FLOOR")
with c3:
    date         = st.date_input("Date", datetime.date.today())
    rate_per_kg  = st.number_input("Rate per KG (₹)", value=85.00, min_value=0.0, format="%.2f")

st.divider()

# Rack Configurations
st.subheader("🏗️ Rack Configurations")
rack_types = st.number_input("Number of Rack Types", min_value=1, max_value=10, value=1)

rack_data = []
for i in range(int(rack_types)):
    with st.expander(f"Module {chr(65+i)} — Configuration", expanded=(i == 0)):
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown("**Quantities**")
            main_qty  = st.number_input("Main Rack Qty",   key=f"mq{i}", value=10, min_value=0)
            addon_qty = st.number_input("Add-on Rack Qty", key=f"aq{i}", value=5,  min_value=0)
            levels    = st.number_input("No. of Levels",   key=f"lv{i}", value=3,  min_value=1)
        with c2:
            st.markdown("**Upright Details**")
            upright_section = st.selectbox("Section Type",
                ["Box Section", "Omega Section 70", "Omega Section 90"],
                key=f"us{i}", index=1)  # default to Omega Section 70
            if upright_section == "Box Section":
                uw = st.number_input("Width (mm)",  key=f"uw{i}", value=80)
                ud = st.number_input("Depth (mm)",  key=f"ud{i}", value=60)
            elif upright_section == "Omega Section 70":
                uw = 70; ud = 45
                st.info("70 × 45 mm Omega")
            else:
                uw = 90; ud = 45
                st.info("90 × 45 mm Omega")
            ul = st.number_input("Height (mm)",    key=f"ul{i}", value=3000)
            ut = st.number_input("Thickness (mm)", key=f"ut{i}", value=1.6, format="%.1f")
        with c3:
            st.markdown("**Beam Details**")
            bt  = st.selectbox("Beam Type", ["Pipe Beam","Roll Formed Beam"], key=f"bt{i}")
            bh  = st.number_input("Height (mm)",    key=f"bh{i}", value=100)
            bw  = st.number_input("Width (mm)",     key=f"bw{i}", value=50)
            bl  = st.number_input("Length (mm)",    key=f"bl{i}", value=2000)
            bth = st.number_input("Thickness (mm)", key=f"bth{i}", value=1.6, format="%.1f")
        with c4:
            st.markdown("**Rack / Deep Bar**")
            depth = st.number_input("Rack Depth (mm)",     key=f"dp{i}", value=800)
            dth   = st.number_input("Deep Bar Thick (mm)", key=f"dt{i}", value=1.6, format="%.1f")
        with c5:
            st.markdown("**Cross Brace**")
            method = st.selectbox("Method (mm)", [200, 500], key=f"mt{i}")
            gap    = st.selectbox("Gap (mm)",    [600, 900], key=f"gp{i}")
            cth    = st.number_input("Cross Thick (mm)", key=f"ct{i}", value=1.6, format="%.1f")

        rack_data.append({
            "module": chr(65+i),
            "main_qty": main_qty, "addon_qty": addon_qty, "levels": levels,
            "uw": uw, "ud": ud, "ul": ul, "ut": ut,
            "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
            "depth": depth, "dth": dth,
            "method": method, "gap": gap, "cth": cth,
        })

st.divider()

# ── ACCESSORIES SECTION ───────────────────────────────────────────────────────
st.subheader("🔩 Accessories")
st.markdown('<div class="acc-section">Enter quantities for accessories. Leave 0 / empty to skip.</div>',
            unsafe_allow_html=True)

with st.expander("Configure Accessories", expanded=False):
    a1, a2 = st.columns(2)
    with a1:
        st.markdown("**Standard Items**")
        cg_qty = st.number_input("Column Guard Qty",  min_value=0, value=0, key="cg")
        rc_qty = st.number_input("Row Connector Qty", min_value=0, value=0, key="rc")

    with a2:
        st.markdown("**Row Guards**")
        rg_types_n = st.number_input("Row Guard Types", min_value=0, max_value=5, value=0, key="rgt")

    row_guards = []
    for j in range(int(rg_types_n)):
        with st.expander(f"Row Guard Type {j+1}"):
            rg_c1, rg_c2, rg_c3 = st.columns(3)
            rg_h   = rg_c1.number_input("Height (mm)", key=f"rgh{j}", value=400.0, format="%.1f")
            rg_l   = rg_c2.number_input("Length (mm)", key=f"rgl{j}", value=2000.0, format="%.1f")
            rg_qty = rg_c3.number_input("Qty",         key=f"rgq{j}", value=1, min_value=0)
            row_guards.append({"h": rg_h, "l": rg_l, "qty": rg_qty})

    st.markdown("**Tie Beams**")
    tb_types_n = st.number_input("Tie Beam Types", min_value=0, max_value=5, value=0, key="tbt")
    tie_beams = []
    for j in range(int(tb_types_n)):
        with st.expander(f"Tie Beam Type {j+1}"):
            tc1, tc2, tc3, tc4, tc5 = st.columns(5)
            tb_qty = tc1.number_input("Qty",       key=f"tbq{j}", value=1, min_value=0)
            tb_w   = tc2.number_input("Width",     key=f"tbw{j}", value=80.0, format="%.1f")
            tb_d   = tc3.number_input("Depth",     key=f"tbd{j}", value=60.0, format="%.1f")
            tb_l   = tc4.number_input("Length",    key=f"tbl{j}", value=2000.0, format="%.1f")
            tb_t   = tc5.number_input("Thickness", key=f"tbt2{j}", value=1.6, format="%.1f")
            tie_beams.append({"qty": tb_qty, "w": tb_w, "d": tb_d, "l": tb_l, "t": tb_t})

    st.markdown("**Back Pallet Stoppers**")
    bps_types_n = st.number_input("BPS Types", min_value=0, max_value=5, value=0, key="bpst")
    bps_list = []
    for j in range(int(bps_types_n)):
        with st.expander(f"BPS Type {j+1}"):
            bc1, bc2 = st.columns(2)
            bps_qty = bc1.number_input("Qty",    key=f"bpsq{j}", value=1, min_value=0)
            bps_l   = bc2.number_input("Length (mm)", key=f"bpsl{j}", value=2000.0, format="%.1f")
            bps_list.append({"qty": bps_qty, "l": bps_l})

acc_data = {
    "cg_qty": cg_qty,
    "rc_qty": rc_qty,
    "row_guards": row_guards,
    "tie_beams": tie_beams,
    "bps_list": bps_list,
}

st.divider()

# Live preview metrics
if rack_data:
    comp0    = calc_components(rack_data[0])
    all_wt   = sum(
        calc_components(r)["total_main"]  * r["main_qty"] +
        calc_components(r)["total_addon"] * r["addon_qty"]
        for r in rack_data
    )
    prev_tot   = all_wt * rate_per_kg
    prev_gst   = prev_tot * 0.18
    prev_grand = prev_tot + prev_gst

    p1,p2,p3,p4,p5,p6 = st.columns(6)
    p1.metric("Main Rack Wt (Mod A)",   f"{comp0['total_main']:.3f} kg")
    p2.metric("Add-on Rack Wt (Mod A)", f"{comp0['total_addon']:.3f} kg")
    p3.metric("Total Tonnage",          f"{all_wt/1000:.3f} MT")
    p4.metric("Basic Amount",           f"₹{prev_tot:,.0f}")
    p5.metric("GST (18%)",              f"₹{prev_gst:,.0f}")
    p6.metric("Grand Total",            f"₹{prev_grand:,.0f}")

st.divider()

if st.button("🐬  GENERATE QUOTATION + BOM", type="primary", use_container_width=True):
    safe  = client.replace(" ","_").replace("/","-")
    fname = f"{safe}_Offer_{offer_no.replace('/','-')}.xlsx"
    out   = os.path.join(tempfile.gettempdir(), fname)

    try:
        basic, gst, grand = build_excel(
            client, product, offer_no, date, project_name,
            rack_data, rate_per_kg, acc_data=acc_data,
            out_path=out, logo_path=logo_path
        )
        st.success("✅ Workbook ready — **Commercial Offer** + **Bill of Materials** sheets.")
        with open(out, "rb") as f:
            st.download_button(
                "⬇️  DOWNLOAD EXCEL (Quotation + BOM)",
                data=f, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        r1, r2, r3 = st.columns(3)
        r1.metric("Basic Amount", f"₹{basic:,.2f}")
        r2.metric("GST (18%)",    f"₹{gst:,.2f}")
        r3.metric("Grand Total",  f"₹{grand:,.2f}", delta="Incl. GST")
    except Exception as e:
        st.error(f"Error generating file: {e}")
        raise
