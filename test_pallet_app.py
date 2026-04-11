import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import datetime, io, base64, tempfile, os, math
 
DSS_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAIAAAB7GkOtAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwAAADsABataJCQAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMZNuFAAAAA=="
 
# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"   # lighter navy (was 0D2137)
NAVY_MID    = "2B5280"   # lighter mid navy (was 1B3A5C)
NAVY_LIGHT  = "3D6FA8"   # accent
ORANGE      = "E8722A"
ORANGE_SOFT = "F4A261"
CREAM       = "FDF6EE"
WHITE       = "FFFFFF"
LIGHT_BLUE  = "EBF3FC"
SILVER      = "CBD5E0"
DARK_TEXT   = "1A1A2E"
MID_TEXT    = "4A5568"
TABLE_ALT   = "F0F7FF"
 
# ── HELPERS ──────────────────────────────────────────────────────────────────
def _side(color, style="thin"):
    return Side(style=style, color=color)
 
def thin(c=SILVER):
    s = _side(c)
    return Border(left=s, right=s, top=s, bottom=s)
 
def medium_border(c=NAVY_MID):
    s = _side(c, "medium")
    return Border(left=s, right=s, top=s, bottom=s)
 
def thick_border(c=NAVY):
    s = _side(c, "thick")
    return Border(left=s, right=s, top=s, bottom=s)
 
def _cell(ws, row, col, value=None, bold=False, size=10, color=DARK_TEXT,
          bg=None, ha="left", va="center", wrap=False, num_fmt=None,
          border=None, italic=False, indent=0):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.font = Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap, indent=indent)
    if bg:
        c.fill = PatternFill("solid", start_color=bg, end_color=bg)
    if num_fmt:
        c.number_format = num_fmt
    if border:
        c.border = border
    return c
 
def fill_row(ws, row, c1, c2, bg):
    for col in range(c1, c2+1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", start_color=bg, end_color=bg)
 
def border_row(ws, row, c1, c2, border):
    for col in range(c1, c2+1):
        ws.cell(row=row, column=col).border = border
 
def mg(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
 
# ── PHYSICS ──────────────────────────────────────────────────────────────────
def weight_kg(length_mm, width_mm, thickness_mm):
    return (length_mm * width_mm * thickness_mm * 7.85) / 1_000_000
 
def upright_sheet_size(w, d):      return w + d + w + 15
def pipe_beam_perimeter(h, w):     return 2 * (h + w)
def roll_beam_perimeter(h, w):     return (4*h + 2*w) + 20
def deep_bar_length(depth):        return depth - 65
def cross_brace_length(d, gap, ud): return math.sqrt((d - 50)**2 + gap**2) + 50
 
def calc_components(rack):
    u_sheet = upright_sheet_size(rack["uw"], rack["ud"])
    u_wt    = weight_kg(rack["ul"], u_sheet, rack["ut"])
 
    if rack["bt"] == "Pipe Beam":
        b_perim = pipe_beam_perimeter(rack["bh"], rack["bw"])
    else:
        b_perim = roll_beam_perimeter(rack["bh"], rack["bw"])
    b_wt = weight_kg(rack["bl"], b_perim, rack["bth"])
 
    d_len = deep_bar_length(rack["depth"])
    d_wt  = weight_kg(d_len, 92, rack["dth"])
 
    c_len = cross_brace_length(d_len, rack["gap"], rack["ud"])
    c_wt  = weight_kg(c_len, 92, rack["cth"])
 
    return {
        "upright_sheet_w":   round(u_sheet, 1),
        "upright_len":       rack["ul"],
        "upright_wt_each":   round(u_wt, 3),
        "upright_qty_main":  4, "upright_qty_addon": 2,
        "beam_perim":        round(b_perim, 1),
        "beam_len":          rack["bl"],
        "beam_wt_each":      round(b_wt, 3),
        "beam_qty_per_level": 2,
        "levels":            rack["levels"],
        "deep_bar_len":      round(d_len, 1),
        "deep_bar_wt_each":  round(d_wt, 3),
        "deep_bar_qty_main": 4, "deep_bar_qty_addon": 2,
        "cross_len":         round(c_len, 1),
        "cross_wt_each":     round(c_wt, 3),
        "cross_qty_main":    2, "cross_qty_addon": 1,
        "main_total_wt":     round(u_wt*4 + b_wt*2*rack["levels"] + d_wt*4 + c_wt*2, 2),
        "addon_total_wt":    round(u_wt*2 + b_wt*1*rack["levels"] + d_wt*2 + c_wt*1, 2),
    }
 
# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — COMMERCIAL OFFER
# ═══════════════════════════════════════════════════════════════════════════════
def build_quotation_sheet(ws, client, product, offer_no, date_obj,
                          project_name, rack_data, rate_per_kg, logo_b64):
    ws.sheet_view.showGridLines = False
 
    # column widths  A=1..I=9
    for col, w in {1:2, 2:4, 3:8, 4:30, 5:14, 6:9, 7:16, 8:16, 9:2}.items():
        ws.column_dimensions[get_column_letter(col)].width = w
 
    R = 1
 
    # ── Try embed logo ────────────────────────────────────────────────────────
    logo_tmp = None
    try:
        raw = base64.b64decode(logo_b64)
        if len(raw) > 500:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
                tf.write(raw)
                logo_tmp = tf.name
    except Exception:
        pass
 
    # ── HEADER ────────────────────────────────────────────────────────────────
    # Top accent line
    ws.row_dimensions[R].height = 6
    fill_row(ws, R, 1, 9, ORANGE); R += 1
 
    # Logo + company block
    ws.row_dimensions[R].height = 70
    fill_row(ws, R, 1, 9, NAVY)
 
    if logo_tmp:
        try:
            img = XLImage(logo_tmp)
            img.width, img.height = 175, 62
            img.anchor = f"B{R}"
            ws.add_image(img)
        except Exception:
            pass
 
    # BRIJ INDUSTRIES — centered across full width
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES"
    c.font = Font(name="Calibri", size=20, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    R += 1
 
    # Tagline strip
    ws.row_dimensions[R].height = 16
    fill_row(ws, R, 1, 9, NAVY_MID)
    mg(ws, R, 2, R, 8)
    _cell(ws, R, 2, "DSS Dolphin Storage Solutions  |  Modular Mezzanine & Racking Systems",
          size=9, color="BDD5F0", bg=NAVY_MID, ha="center")
    R += 1
 
    # Address / contact
    ws.row_dimensions[R].height = 14
    fill_row(ws, R, 1, 9, NAVY_LIGHT)
    mg(ws, R, 2, R, 8)
    _cell(ws, R, 2, "86/3/1 Road No 7, Mundka Industrial Area, New Delhi – 110041   |   GST: 07AAMFB6403G1ZM   |   +91 9625589161 / 9811096149",
          size=8, color=WHITE, bg=NAVY_LIGHT, ha="center")
    R += 1
 
    ws.row_dimensions[R].height = 13
    fill_row(ws, R, 1, 9, NAVY_LIGHT)
    mg(ws, R, 2, R, 8)
    _cell(ws, R, 2, "brijindustries09@rediffmail.com   |   www.brijindustries.in",
          size=8, color="D0E8FF", bg=NAVY_LIGHT, ha="center", italic=True)
    R += 1
 
    # Bottom accent
    ws.row_dimensions[R].height = 5
    fill_row(ws, R, 1, 9, ORANGE); R += 1
    ws.row_dimensions[R].height = 8; R += 1
 
    # ── OFFER TITLE ───────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 24
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "COMMERCIAL OFFER"
    c.font = Font(name="Calibri", size=13, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    c.border = medium_border(NAVY)
    R += 1
 
    # ── DETAIL ROWS ───────────────────────────────────────────────────────────
    def detail_row(lbl_l, val_l, lbl_r, val_r, bg=LIGHT_BLUE):
        nonlocal R
        ws.row_dimensions[R].height = 20
        fill_row(ws, R, 2, 8, bg)
        border_row(ws, R, 2, 8, thin())
        mg(ws, R, 2, R, 3)
        _cell(ws, R, 2, lbl_l, bold=True, size=9, color=NAVY_MID, bg=bg, ha="right")
        mg(ws, R, 4, R, 5)
        _cell(ws, R, 4, val_l, size=9, color=DARK_TEXT, bg=bg, indent=1)
        mg(ws, R, 6, R, 6)
        _cell(ws, R, 6, lbl_r, bold=True, size=9, color=NAVY_MID, bg=bg, ha="right")
        mg(ws, R, 7, R, 8)
        _cell(ws, R, 7, val_r, size=9, color=DARK_TEXT, bg=bg, indent=1)
        R += 1
 
    detail_row("To :",      client,       "Date :",     date_obj.strftime("%d %B %Y"), bg=LIGHT_BLUE)
    detail_row("Product :", product,      "Offer No. :", offer_no,                     bg=WHITE)
    detail_row("Project :", project_name, "",            "",                            bg=LIGHT_BLUE)
 
    ws.row_dimensions[R].height = 10; R += 1
 
    # ── SCOPE TABLE ───────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 22
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "SCOPE OF SUPPLY"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=ORANGE, end_color=ORANGE)
    R += 1
 
    # Headers
    ws.row_dimensions[R].height = 26
    hdrs = [(2,"SR.","center"),(3,"","center"),(4,"DESCRIPTION","center"),
            (5,"TYPE","center"),(6,"QTY","center"),(7,"UNIT PRICE (₹)","center"),(8,"AMOUNT (₹)","center")]
    for col, txt, al in hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
        c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
        c.alignment = Alignment(horizontal=al, vertical="center")
        c.border = thin(SILVER)
    R += 1
 
    total_basic = 0
    sr = 1
 
    for rack in rack_data:
        comp = calc_components(rack)
        main_wt    = comp["main_total_wt"]
        addon_wt   = comp["addon_total_wt"]
        main_price = main_wt * rate_per_kg
        addon_price= addon_wt * rate_per_kg
        main_total = main_price * rack["main_qty"]
        total_basic += main_total
 
        bg = WHITE if sr % 2 == 1 else TABLE_ALT
        ws.row_dimensions[R].height = 20
        fill_row(ws, R, 2, 8, bg)
        border_row(ws, R, 2, 8, thin())
        _cell(ws, R, 2, sr, size=9, color=MID_TEXT, bg=bg, ha="center")
        mg(ws, R, 3, R, 4)
        _cell(ws, R, 3, f"  {rack['module']} – {rack['name']}",
              bold=True, size=9, color=DARK_TEXT, bg=bg)
        _cell(ws, R, 5, "Main Rack", size=9, color=NAVY_MID, bg=bg, ha="center")
        _cell(ws, R, 6, rack["main_qty"], size=9, bg=bg, ha="center")
        _cell(ws, R, 7, round(main_price, 2), size=9, bg=bg, ha="right", num_fmt="#,##0.00")
        _cell(ws, R, 8, round(main_total, 2), bold=True, size=9,
              color=NAVY, bg=bg, ha="right", num_fmt="#,##0.00")
        R += 1; sr += 1
 
        if rack["addon_qty"] > 0:
            addon_total  = addon_price * rack["addon_qty"]
            total_basic += addon_total
            bg = WHITE if sr % 2 == 1 else TABLE_ALT
            ws.row_dimensions[R].height = 20
            fill_row(ws, R, 2, 8, bg)
            border_row(ws, R, 2, 8, thin())
            _cell(ws, R, 2, sr, size=9, color=MID_TEXT, bg=bg, ha="center")
            mg(ws, R, 3, R, 4)
            _cell(ws, R, 3, f"  {rack['module']} – {rack['name']}",
                  size=9, color=MID_TEXT, bg=bg)
            _cell(ws, R, 5, "Add-on Rack", size=9, color=ORANGE, bg=bg, ha="center")
            _cell(ws, R, 6, rack["addon_qty"], size=9, bg=bg, ha="center")
            _cell(ws, R, 7, round(addon_price, 2), size=9, bg=bg, ha="right", num_fmt="#,##0.00")
            _cell(ws, R, 8, round(addon_total, 2), bold=True, size=9,
                  color=NAVY, bg=bg, ha="right", num_fmt="#,##0.00")
            R += 1; sr += 1
 
    # Subtotal
    ws.row_dimensions[R].height = 22
    mg(ws, R, 2, R, 7)
    c = ws.cell(row=R, column=2)
    c.value = "SUBTOTAL (BASIC AMOUNT)"
    c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    c.fill = PatternFill("solid", start_color=LIGHT_BLUE, end_color=LIGHT_BLUE)
    c.border = medium_border(NAVY_MID)
    _cell(ws, R, 8, round(total_basic, 2), bold=True, size=10, color=NAVY,
          bg=LIGHT_BLUE, ha="right", num_fmt="#,##0.00", border=medium_border(NAVY_MID))
    R += 1
 
    ws.row_dimensions[R].height = 10; R += 1
 
    # ── PRICING SUMMARY ───────────────────────────────────────────────────────
    gst_amount  = round(total_basic * 0.18, 2)
    grand_total = round(total_basic + gst_amount, 2)
 
    def sum_row(lbl, val, bg=WHITE, bold=False, vc=DARK_TEXT, is_text=False):
        nonlocal R
        ws.row_dimensions[R].height = 20
        mg(ws, R, 5, R, 7)
        c_l = ws.cell(row=R, column=5)
        c_l.value = lbl
        c_l.font = Font(name="Calibri", size=10, bold=bold, color=MID_TEXT)
        c_l.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        c_l.fill = PatternFill("solid", start_color=bg, end_color=bg)
        c_l.border = thin()
        c_v = ws.cell(row=R, column=8)
        c_v.value = val
        c_v.font = Font(name="Calibri", size=10, bold=bold, color=vc, italic=is_text)
        c_v.alignment = Alignment(horizontal="center" if is_text else "right",
                                   vertical="center", indent=1)
        c_v.fill = PatternFill("solid", start_color=bg, end_color=bg)
        c_v.border = thin()
        if not is_text:
            c_v.number_format = "#,##0.00"
        R += 1
 
    sum_row("Basic Amount",     round(total_basic, 2), bg=WHITE, bold=True, vc=NAVY)
    sum_row("Freight Charges",  "Inclusive", bg=CREAM, is_text=True, vc=ORANGE)
    sum_row("Erection Charges", "Inclusive", bg=CREAM, is_text=True, vc=ORANGE)
    sum_row("GST @ 18%",        gst_amount, bg=WHITE, vc=MID_TEXT)
 
    # Grand Total
    ws.row_dimensions[R].height = 28
    mg(ws, R, 2, R, 7)
    c_gl = ws.cell(row=R, column=2)
    c_gl.value = "GRAND TOTAL  (Inclusive of GST @ 18%)"
    c_gl.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c_gl.alignment = Alignment(horizontal="center", vertical="center")
    c_gl.fill = PatternFill("solid", start_color=ORANGE, end_color=ORANGE)
    c_gl.border = medium_border(ORANGE)
    c_gv = ws.cell(row=R, column=8)
    c_gv.value = grand_total
    c_gv.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c_gv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    c_gv.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    c_gv.border = medium_border(NAVY)
    c_gv.number_format = '₹ #,##0.00'
    R += 1
 
    ws.row_dimensions[R].height = 12; R += 1
 
    # ── TERMS + BANK ──────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 20
    mg(ws, R, 2, R, 5)
    c = ws.cell(row=R, column=2)
    c.value = "TERMS & CONDITIONS"
    c.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=NAVY_MID, end_color=NAVY_MID)
    mg(ws, R, 6, R, 8)
    c2 = ws.cell(row=R, column=6)
    c2.value = "BANK DETAILS"
    c2.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill("solid", start_color=NAVY_MID, end_color=NAVY_MID)
    R += 1
 
    terms = [
        "Payment: 50% advance, balance against delivery",
        "Delivery: 10–12 weeks from advance receipt",
        "Warranty: 12 months from commissioning date",
        "GST @ 18% applicable as per Government norms",
        "Prices valid for 30 days from date of offer",
    ]
    bank = [
        "Account Name: BRIJ INDUSTRIES",
        "Bank: HDFC Bank Ltd.",
        "Account No.: 50100512345678",
        "IFSC Code: HDFC0001234",
        "Branch: Mundka, New Delhi",
    ]
    for t, b in zip(terms, bank):
        ws.row_dimensions[R].height = 17
        mg(ws, R, 2, R, 5)
        c = ws.cell(row=R, column=2)
        c.value = f"  • {t}"
        c.font = Font(name="Calibri", size=8, color=MID_TEXT)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.fill = PatternFill("solid", start_color=LIGHT_BLUE, end_color=LIGHT_BLUE)
        c.border = thin()
        mg(ws, R, 6, R, 8)
        c2 = ws.cell(row=R, column=6)
        c2.value = f"  {b}"
        c2.font = Font(name="Calibri", size=8, color=MID_TEXT)
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.fill = PatternFill("solid", start_color=CREAM, end_color=CREAM)
        c2.border = thin()
        R += 1
 
    ws.row_dimensions[R].height = 8; R += 1
 
    # Signature
    ws.row_dimensions[R].height = 40
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
 
    # Footer
    ws.row_dimensions[R].height = 5
    fill_row(ws, R, 1, 9, ORANGE); R += 1
    ws.row_dimensions[R].height = 14
    fill_row(ws, R, 1, 9, NAVY)
    mg(ws, R, 2, R, 8)
    _cell(ws, R, 2, "Thank you for considering DSS Dolphin Storage Solutions. We look forward to serving you.",
          size=8, color="7799BB", bg=NAVY, ha="center", italic=True)
 
    # Print setup
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.print_area = f"A1:{get_column_letter(9)}{R}"
 
    if logo_tmp and os.path.exists(logo_tmp):
        try: os.unlink(logo_tmp)
        except: pass
 
    return total_basic, gst_amount, grand_total
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 2 — BILL OF MATERIALS
# ═══════════════════════════════════════════════════════════════════════════════
def build_bom_sheet(ws, client, offer_no, date_obj, rack_data, rate_per_kg):
    ws.sheet_view.showGridLines = False
 
    # col widths
    for col, w in {1:2, 2:4, 3:22, 4:10, 5:10, 6:10, 7:10,
                   8:10, 9:10, 10:10, 11:10, 12:10, 13:2}.items():
        ws.column_dimensions[get_column_letter(col)].width = w
 
    R = 1
 
    # Top accent
    ws.row_dimensions[R].height = 6
    fill_row(ws, R, 1, 13, ORANGE); R += 1
 
    # Header
    ws.row_dimensions[R].height = 50
    fill_row(ws, R, 1, 13, NAVY)
    mg(ws, R, 2, R, 12)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES — BILL OF MATERIALS"
    c.font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    R += 1
 
    ws.row_dimensions[R].height = 14
    fill_row(ws, R, 1, 13, NAVY_LIGHT)
    mg(ws, R, 2, R, 12)
    _cell(ws, R, 2, f"DSS Dolphin Storage Solutions  |  Offer No: {offer_no}  |  Customer: {client}  |  Date: {date_obj.strftime('%d %B %Y')}",
          size=9, color=WHITE, bg=NAVY_LIGHT, ha="center")
    R += 1
 
    ws.row_dimensions[R].height = 5
    fill_row(ws, R, 1, 13, ORANGE); R += 1
    ws.row_dimensions[R].height = 8; R += 1
 
    # ── For each rack type ────────────────────────────────────────────────────
    for rack in rack_data:
        comp = calc_components(rack)
 
        # Rack title
        ws.row_dimensions[R].height = 22
        mg(ws, R, 2, R, 12)
        c = ws.cell(row=R, column=2)
        c.value = f"  MODULE {rack['module']}  —  {rack['name']}   |   Main Racks: {rack['main_qty']}  |  Add-on Racks: {rack['addon_qty']}"
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.fill = PatternFill("solid", start_color=NAVY_MID, end_color=NAVY_MID)
        R += 1
 
        # Column headers
        ws.row_dimensions[R].height = 30
        BOM_HEADERS = [
            (2,  "SR.",             "center"),
            (3,  "COMPONENT",       "left"),
            (4,  "SECTION / PROFILE","center"),
            (5,  "LENGTH (mm)",     "center"),
            (6,  "THICKNESS (mm)",  "center"),
            (7,  "WT / PCS (kg)",   "center"),
            (8,  "QTY / MAIN",      "center"),
            (9,  "QTY / ADD-ON",    "center"),
            (10, "MAIN TOTAL (kg)", "center"),
            (11, "ADDON TOTAL (kg)","center"),
            (12, "RATE/kg (₹)",     "center"),
        ]
        for col, txt, al in BOM_HEADERS:
            c = ws.cell(row=R, column=col)
            c.value = txt
            c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
            c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
            c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
            c.border = thin(SILVER)
        R += 1
 
        # Component rows
        def comp_row(sr, name, section, length, thick, wt_each,
                     qty_main, qty_addon, bg=WHITE):
            nonlocal R
            ws.row_dimensions[R].height = 18
            main_tot  = round(wt_each * qty_main, 3)
            addon_tot = round(wt_each * qty_addon, 3)
            vals = [
                (2, sr,        "center"),
                (3, f"  {name}", "left"),
                (4, section,   "center"),
                (5, length,    "center"),
                (6, thick,     "center"),
                (7, wt_each,   "center"),
                (8, qty_main,  "center"),
                (9, qty_addon, "center"),
                (10, main_tot, "right"),
                (11, addon_tot,"right"),
                (12, rate_per_kg,"right"),
            ]
            fill_row(ws, R, 2, 12, bg)
            border_row(ws, R, 2, 12, thin())
            for col, val, al in vals:
                c = ws.cell(row=R, column=col)
                c.value = val
                c.font = Font(name="Calibri", size=8, color=DARK_TEXT)
                c.alignment = Alignment(horizontal=al, vertical="center")
                c.fill = PatternFill("solid", start_color=bg, end_color=bg)
                c.border = thin()
                if col in (7, 10, 11, 12) and isinstance(val, float):
                    c.number_format = "#,##0.000"
                if col == 12:
                    c.number_format = "#,##0.00"
            R += 1
 
        levels = rack["levels"]
        upright_section = f"{rack['uw']}×{rack['ud']} Box"
        if rack["bt"] == "Pipe Beam":
            beam_section = f"{rack['bh']}×{rack['bw']} Pipe"
        else:
            beam_section = f"{rack['bh']}×{rack['bw']} Roll"
 
        comp_row(1, "Upright / Column", upright_section,
                 rack["ul"], rack["ut"], comp["upright_wt_each"],
                 4, 2, bg=WHITE)
        comp_row(2, f"Beam ({rack['bt']})  [×{levels} levels]", beam_section,
                 rack["bl"], rack["bth"],
                 round(comp["beam_wt_each"] * levels, 3),
                 2, 1, bg=TABLE_ALT)
        comp_row(3, "Deep Bar (Shelf Support)", "92 mm flat",
                 int(comp["deep_bar_len"]), rack["dth"],
                 comp["deep_bar_wt_each"],
                 4, 2, bg=WHITE)
        comp_row(4, "Cross Brace (Diagonal)", "92 mm flat",
                 int(comp["cross_len"]), rack["cth"],
                 comp["cross_wt_each"],
                 2, 1, bg=TABLE_ALT)
 
        # Summary rows for this module
        ws.row_dimensions[R].height = 8; R += 1
 
        def summary_row(label, val_main, val_addon, bg, bold=False, vc=DARK_TEXT):
            nonlocal R
            ws.row_dimensions[R].height = 20
            fill_row(ws, R, 2, 12, bg)
            mg(ws, R, 2, R, 9)
            c = ws.cell(row=R, column=2)
            c.value = label
            c.font = Font(name="Calibri", size=9, bold=bold, color=vc)
            c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            c.fill = PatternFill("solid", start_color=bg, end_color=bg)
            c.border = medium_border(NAVY_MID) if bold else thin()
            for col, val in [(10, val_main), (11, val_addon)]:
                cv = ws.cell(row=R, column=col)
                cv.value = val
                cv.font = Font(name="Calibri", size=9, bold=bold, color=vc)
                cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
                cv.fill = PatternFill("solid", start_color=bg, end_color=bg)
                cv.border = medium_border(NAVY_MID) if bold else thin()
                cv.number_format = "#,##0.00"
            R += 1
 
        summary_row("Weight — Single Rack (kg)",
                    comp["main_total_wt"], comp["addon_total_wt"],
                    LIGHT_BLUE)
        main_price  = round(comp["main_total_wt"] * rate_per_kg, 2)
        addon_price = round(comp["addon_total_wt"] * rate_per_kg, 2)
        summary_row("Unit Price — Single Rack (₹)",
                    main_price, addon_price, CREAM)
        summary_row(f"TOTAL WEIGHT — All Racks (kg)  [Main×{rack['main_qty']}  |  Add-on×{rack['addon_qty']}]",
                    round(comp["main_total_wt"] * rack["main_qty"], 2),
                    round(comp["addon_total_wt"] * rack["addon_qty"], 2),
                    LIGHT_BLUE, bold=True, vc=NAVY)
        total_main_val  = round(main_price  * rack["main_qty"], 2)
        total_addon_val = round(addon_price * rack["addon_qty"], 2)
        summary_row(f"TOTAL VALUE (₹)  [Main×{rack['main_qty']}  |  Add-on×{rack['addon_qty']}]",
                    total_main_val, total_addon_val,
                    NAVY, bold=True, vc=WHITE)
 
        ws.row_dimensions[R].height = 14; R += 1
 
    # ── Grand summary ─────────────────────────────────────────────────────────
    ws.row_dimensions[R].height = 6
    fill_row(ws, R, 1, 13, ORANGE); R += 1
    ws.row_dimensions[R].height = 22
    mg(ws, R, 2, R, 12)
    c = ws.cell(row=R, column=2)
    c.value = "CONSOLIDATED PRICING SUMMARY"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    R += 1
 
    total_basic = sum(
        calc_components(r)["main_total_wt"] * rate_per_kg * r["main_qty"] +
        calc_components(r)["addon_total_wt"] * rate_per_kg * r["addon_qty"]
        for r in rack_data
    )
    gst     = round(total_basic * 0.18, 2)
    grand   = round(total_basic + gst, 2)
 
    for lbl, val, bg, bold, vc in [
        ("Basic Amount (₹)",       round(total_basic, 2), WHITE, False, DARK_TEXT),
        ("GST @ 18% (₹)",          gst,                   CREAM, False, MID_TEXT),
        ("GRAND TOTAL incl. GST (₹)", grand,              NAVY,  True,  WHITE),
    ]:
        ws.row_dimensions[R].height = 22
        mg(ws, R, 2, R, 9)
        c = ws.cell(row=R, column=2)
        c.value = lbl
        c.font = Font(name="Calibri", size=10, bold=bold, color=vc if bg!=NAVY else WHITE)
        c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        c.fill = PatternFill("solid", start_color=bg, end_color=bg)
        c.border = medium_border(NAVY_MID) if bold else thin()
        mg(ws, R, 10, R, 12)
        cv = ws.cell(row=R, column=10)
        cv.value = val
        cv.font = Font(name="Calibri", size=11 if bold else 10, bold=bold, color=WHITE if bg==NAVY else NAVY)
        cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        cv.fill = PatternFill("solid", start_color=bg if bg!=NAVY else ORANGE, end_color=bg if bg!=NAVY else ORANGE)
        cv.border = medium_border(NAVY_MID) if bold else thin()
        cv.number_format = "₹ #,##0.00" if bold else "#,##0.00"
        R += 1
 
    ws.row_dimensions[R].height = 5
    fill_row(ws, R, 1, 13, ORANGE); R += 1
    ws.row_dimensions[R].height = 14
    fill_row(ws, R, 1, 13, NAVY)
    mg(ws, R, 2, R, 12)
    _cell(ws, R, 2, "BRIJ INDUSTRIES  |  DSS Dolphin Storage Solutions  |  www.brijindustries.in",
          size=8, color="7799BB", bg=NAVY, ha="center", italic=True)
 
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER BUILD
# ═══════════════════════════════════════════════════════════════════════════════
def build_excel(client, product, offer_no, date_obj, project_name,
                rack_data, rate_per_kg, out_path="quotation.xlsx"):
    wb = Workbook()
 
    ws_q = wb.active
    ws_q.title = "Commercial Offer"
    ws_b = wb.create_sheet("Bill of Materials")
 
    total_basic, gst, grand = build_quotation_sheet(
        ws_q, client, product, offer_no, date_obj,
        project_name, rack_data, rate_per_kg, DSS_LOGO_B64
    )
    build_bom_sheet(ws_b, client, offer_no, date_obj, rack_data, rate_per_kg)
 
    wb.save(out_path)
    return total_basic, gst, grand
 
 
# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="DSS Quotation Generator", layout="wide", page_icon="🐬")
 
st.markdown("""
<style>
body { font-family: Calibri, Arial, sans-serif; }
.main-header {
    background: linear-gradient(135deg, #1A3A5C 0%, #2B5280 55%, #E8722A 100%);
    padding: 26px 36px; border-radius: 14px; color: white;
    text-align: center; margin-bottom: 24px;
    box-shadow: 0 8px 28px rgba(0,0,0,0.22);
}
.main-header h1 { margin:0; font-size:1.9rem; letter-spacing:2px; }
.main-header p  { margin:5px 0 0; font-size:0.92rem; opacity:0.82; }
.stButton > button {
    background: linear-gradient(135deg, #E8722A, #c45a1a);
    color: white; font-weight: bold; border: none;
    border-radius: 10px; padding: 13px 30px;
    font-size: 1.05rem; transition: all 0.2s ease;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 22px rgba(232,114,42,0.45);
}
</style>
""", unsafe_allow_html=True)
 
st.markdown("""
<div class="main-header">
    <h1>🐬  QUOTATION GENERATOR</h1>
    <p>BRIJ INDUSTRIES — DSS Dolphin Storage Solutions  |  Modular Mezzanine & Racking Systems</p>
</div>
""", unsafe_allow_html=True)
 
# ── Customer Details ──────────────────────────────────────────────────────────
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
 
# ── Rack Configurations ───────────────────────────────────────────────────────
st.subheader("🏗️ Rack Configurations")
rack_types = st.number_input("Number of Rack Types", min_value=1, max_value=10, value=1)
 
rack_data = []
for i in range(int(rack_types)):
    with st.expander(f"Rack Type {chr(65+i)} — Configuration", expanded=(i == 0)):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            name      = st.text_input("Rack Name",          key=f"n{i}",  value=f"MODULE {chr(65+i)}")
            main_qty  = st.number_input("Main Rack Qty",    key=f"mq{i}", value=10, min_value=0)
            addon_qty = st.number_input("Add-on Rack Qty",  key=f"aq{i}", value=5,  min_value=0)
            levels    = st.number_input("No. of Levels",    key=f"lv{i}", value=3,  min_value=1)
        with c2:
            uw  = st.number_input("Upright Width (mm)",     key=f"uw{i}", value=80)
            ud  = st.number_input("Upright Depth (mm)",     key=f"ud{i}", value=60)
            ul  = st.number_input("Upright Length (mm)",    key=f"ul{i}", value=3000)
            ut  = st.number_input("Upright Thickness (mm)", key=f"ut{i}", value=1.6, format="%.1f")
        with c3:
            bt  = st.selectbox("Beam Type", ["Pipe Beam","Roll Formed Beam"], key=f"bt{i}")
            bh  = st.number_input("Beam Height (mm)",       key=f"bh{i}", value=100)
            bw  = st.number_input("Beam Width (mm)",        key=f"bw{i}", value=50)
            bl  = st.number_input("Beam Length (mm)",       key=f"bl{i}", value=2000)
            bth = st.number_input("Beam Thickness (mm)",    key=f"bth{i}", value=1.6, format="%.1f")
        with c4:
            depth = st.number_input("Rack Depth (mm)",         key=f"dp{i}", value=800)
            dth   = st.number_input("Deep Bar Thickness (mm)", key=f"dt{i}", value=1.6, format="%.1f")
            gap   = st.number_input("Cross Brace Gap (mm)",    key=f"gp{i}", value=500)
            cth   = st.number_input("Cross Brace Thickness",   key=f"ct{i}", value=1.6, format="%.1f")
 
        rack_data.append({
            "module": chr(65+i), "name": name,
            "main_qty": main_qty, "addon_qty": addon_qty, "levels": levels,
            "uw": uw, "ud": ud, "ul": ul, "ut": ut,
            "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
            "depth": depth, "dth": dth, "gap": gap, "cth": cth,
        })
 
st.divider()
 
# ── Live preview ──────────────────────────────────────────────────────────────
if rack_data:
    comp0     = calc_components(rack_data[0])
    prev_tot  = sum(
        calc_components(r)["main_total_wt"] * rate_per_kg * r["main_qty"] +
        calc_components(r)["addon_total_wt"] * rate_per_kg * r["addon_qty"]
        for r in rack_data
    )
    prev_gst  = prev_tot * 0.18
    prev_grand= prev_tot + prev_gst
 
    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("Main Rack Weight",  f"{comp0['main_total_wt']:.2f} kg")
    p2.metric("Add-on Rack Weight",f"{comp0['addon_total_wt']:.2f} kg")
    p3.metric("Basic Amount",      f"₹{prev_tot:,.0f}")
    p4.metric("GST (18%)",         f"₹{prev_gst:,.0f}")
    p5.metric("Grand Total",       f"₹{prev_grand:,.0f}")
 
st.divider()
 
# ── Generate ──────────────────────────────────────────────────────────────────
if st.button("🐬  GENERATE QUOTATION + BOM", type="primary", use_container_width=True):
    safe = client.replace(" ", "_").replace("/", "-")
    fname = f"{safe}_Offer_{offer_no.replace('/', '-')}.xlsx"
    out   = os.path.join(tempfile.gettempdir(), fname)
 
    try:
        basic, gst, grand = build_excel(
            client, product, offer_no, date,
            project_name, rack_data, rate_per_kg, out_path=out
        )
        st.success("✅ Workbook generated — contains **Commercial Offer** + **Bill of Materials** sheets.")
        with open(out, "rb") as f:
            st.download_button(
                "⬇️  DOWNLOAD EXCEL (Quotation + BOM)",
                data=f,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        r1, r2, r3 = st.columns(3)
        r1.metric("Basic Amount", f"₹{basic:,.2f}")
        r2.metric("GST (18%)",    f"₹{gst:,.2f}")
        r3.metric("Grand Total",  f"₹{grand:,.2f}", delta="Incl. GST")
    except Exception as e:
        st.error(f"Error: {e}")
        raise
