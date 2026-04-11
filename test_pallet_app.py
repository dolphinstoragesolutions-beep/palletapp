import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import datetime, base64, tempfile, os, math, requests

# ── REAL DSS LOGO (fetched from web or embedded placeholder) ─────────────────
# We embed the original DSS logo from the first version of the code
DSS_LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAIAAAB7GkOtAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwAAADsABataJCQAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMZNuFAAAAA=="

# We will use the full logo embedded in the original file
FULL_LOGO_B64 = """iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAIAAAB7GkOtAAAgAElEQVR4nO2dd3wURf/Hv7O7V3IhJCEJCSGU0DshgPQiICBIkSZFQFEURLEgKqIIKiAoHRAUpAkCUqT3JiWUACGkkEp6Lb8/Nr8LuWQvuSSXvd1kPq+X+5LbO2bnc7OT+ebMzBkQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQBEEQ"""

# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY        = "1E4080"   # medium navy blue - professional
NAVY_MID    = "2B5FAD"   # mid blue
NAVY_LIGHT  = "4A7FC1"   # lighter accent
NAVY_HEADER = "1A3A6C"   # dark header
ORANGE      = "E8722A"
ORANGE_DARK = "C45A1A"
CREAM       = "FFFBF5"
WHITE       = "FFFFFF"
LIGHT_BLUE  = "EBF3FC"
TABLE_ALT   = "F2F7FD"
SILVER      = "BDC8D8"
DARK_TEXT   = "1A2744"
MID_TEXT    = "4A5678"
GREEN       = "217346"
GREEN_LIGHT = "E8F5EE"

# ── BORDER HELPERS ────────────────────────────────────────────────────────────
def side(color, style="thin"):
    return Side(style=style, color=color)

def b_thin(c=SILVER):
    s = side(c)
    return Border(left=s, right=s, top=s, bottom=s)

def b_med(c=NAVY_MID):
    s = side(c, "medium")
    return Border(left=s, right=s, top=s, bottom=s)

def b_thick(c=NAVY):
    s = side(c, "thick")
    return Border(left=s, right=s, top=s, bottom=s)

def b_outer(color=NAVY_MID, style="medium"):
    """Only outer border (left,right,top,bottom) - good for merged ranges"""
    s = side(color, style)
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
    for col in range(c1, c2+1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=bg)
        if bdr:
            ws.cell(row=row, column=col).border = bdr

def mg(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def set_row_h(ws, r, h):
    ws.row_dimensions[r].height = h

# ── PHYSICS ───────────────────────────────────────────────────────────────────
def wt_kg(L, W, T):
    """Weight in kg given length, width, thickness all in mm, density 7.85 g/cc"""
    return (L * W * T * 7.85) / 1_000_000

def calc_components(rack):
    uw, ud, ul, ut = rack["uw"], rack["ud"], rack["ul"], rack["ut"]
    bh, bw, bl, bth = rack["bh"], rack["bw"], rack["bl"], rack["bth"]
    depth, dth = rack["depth"], rack["dth"]
    gap, cth = rack["gap"], rack["cth"]
    levels = rack["levels"]

    u_sheet = uw + ud + uw + 15
    u_wt = wt_kg(ul, u_sheet, ut)

    if rack["bt"] == "Pipe Beam":
        b_perim = 2 * (bh + bw)
    else:
        b_perim = (4*bh + 2*bw) + 20
    b_wt = wt_kg(bl, b_perim, bth)

    d_len = depth - 65
    d_wt = wt_kg(d_len, 92, dth)

    c_len = math.sqrt((d_len - 50)**2 + gap**2) + 50
    c_wt = wt_kg(c_len, 92, cth)

    main_wt  = round(u_wt*4 + b_wt*2*levels + d_wt*4 + c_wt*2, 3)
    addon_wt = round(u_wt*2 + b_wt*1*levels + d_wt*2 + c_wt*1, 3)

    # UDL = load per level per unit area (beam span × depth)
    # Assuming typical design load 500 kg/m²
    beam_span_m = bl / 1000
    depth_m = depth / 1000
    area_m2 = beam_span_m * depth_m
    udl_kg_per_m2 = 500  # standard design UDL

    return {
        "u_sheet": round(u_sheet, 1), "ul": ul, "ut": ut,
        "u_wt": round(u_wt, 4),
        "b_perim": round(b_perim, 1), "bl": bl, "bth": bth,
        "b_wt": round(b_wt, 4),
        "d_len": round(d_len, 1), "dth": dth,
        "d_wt": round(d_wt, 4),
        "c_len": round(c_len, 1), "cth": cth,
        "c_wt": round(c_wt, 4),
        "levels": levels,
        "main_wt": main_wt,
        "addon_wt": addon_wt,
        "area_m2": round(area_m2, 3),
        "udl_kg_m2": udl_kg_per_m2,
        "load_per_level_kg": round(udl_kg_per_m2 * area_m2, 1),
    }

# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — COMMERCIAL OFFER
# ═══════════════════════════════════════════════════════════════════════════════
def build_quotation_sheet(ws, client, product, offer_no, date_obj,
                          project_name, rack_data, rate_per_kg, logo_path=None):
    ws.sheet_view.showGridLines = False

    # Columns: A=1(margin), B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9(margin)
    col_w = {1:1.5, 2:5, 3:6, 4:28, 5:14, 6:9, 7:16, 8:17, 9:1.5}
    for col, w in col_w.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    R = 1

    # ── TOP ACCENT ───────────────────────────────────────────────────────────
    set_row_h(ws, R, 5); fill(ws, R, 1, 9, ORANGE); R += 1

    # ── LOGO + COMPANY NAME ──────────────────────────────────────────────────
    set_row_h(ws, R, 68)
    fill(ws, R, 1, 9, NAVY_HEADER)

    if logo_path and os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 190; img.height = 60
            img.anchor = f"B{R}"
            ws.add_image(img)
        except Exception:
            pass

    # Company name centered across B:H
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES"
    c.font = Font(name="Calibri", size=22, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    R += 1

    # Sub-tagline
    set_row_h(ws, R, 16); fill(ws, R, 1, 9, NAVY_MID)
    mg(ws, R, 2, R, 8)
    W(ws, R, 2, "DSS Dolphin Storage Solutions  ·  Modular Mezzanine & Racking Systems",
      sz=9, color="C8DEFF", bg=NAVY_MID, ha="center")
    R += 1

    # Address
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

    # Bottom accent
    set_row_h(ws, R, 5); fill(ws, R, 1, 9, ORANGE); R += 1
    set_row_h(ws, R, 9); R += 1

    # ── OFFER TITLE ──────────────────────────────────────────────────────────
    set_row_h(ws, R, 24)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "COMMERCIAL OFFER"
    c.font = Font(name="Calibri", size=13, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    c.border = b_med(NAVY_HEADER)
    R += 1

    # ── CUSTOMER DETAIL BLOCK ────────────────────────────────────────────────
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

    set_row_h(ws, R, 10); R += 1

    # ── RACK DETAILS SECTION ─────────────────────────────────────────────────
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "TECHNICAL DETAILS"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_MID)
    c.border = b_med(NAVY_MID)
    R += 1

    # Technical detail header row
    set_row_h(ws, R, 22)
    tech_hdrs = [(2,"MODULE","center"),(3,"UPRIGHT\nHEIGHT (mm)","center"),
                 (4,"BEAM\nLENGTH (mm)","center"),(5,"RACK\nDEPTH (mm)","center"),
                 (6,"LEVELS","center"),(7,"LOAD / LEVEL\n(kg)","center"),
                 (8,"PALLETS\n/ LEVEL","center")]
    for col, txt, al in tech_hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", sz=8, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        c.border = b_thin()
    R += 1

    for idx, rack in enumerate(rack_data):
        comp = calc_components(rack)
        bg = WHITE if idx % 2 == 0 else TABLE_ALT
        set_row_h(ws, R, 18)
        fill(ws, R, 2, 8, bg, b_thin())
        module_label = f"Module {rack['module']}"
        W(ws, R, 2, module_label, bold=True, sz=9, color=NAVY_MID, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 3, rack["ul"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 4, rack["bl"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 5, rack["depth"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 6, rack["levels"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 7, comp["load_per_level_kg"], sz=9, bg=bg, ha="center", bdr=b_thin(), fmt="#,##0")
        # Pallets: assume 1 pallet per 1.2m of beam length
        pallets = max(1, int(rack["bl"] / 1200))
        W(ws, R, 8, pallets, sz=9, bg=bg, ha="center", bdr=b_thin())
        R += 1

    set_row_h(ws, R, 10); R += 1

    # ── SCOPE OF SUPPLY ──────────────────────────────────────────────────────
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 8)
    c = ws.cell(row=R, column=2)
    c.value = "SCOPE OF SUPPLY"
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=ORANGE)
    c.border = b_med(ORANGE_DARK)
    R += 1

    # Table headers
    set_row_h(ws, R, 26)
    fill(ws, R, 2, 8, NAVY_HEADER)
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
        main_price  = comp["main_wt"] * rate_per_kg
        addon_price = comp["addon_wt"] * rate_per_kg
        main_total  = main_price * rack["main_qty"]
        total_basic += main_total

        # Description: "MODULE A" not "A MODULE A"
        desc = f"MODULE {rack['module']}"

        bg = WHITE if sr % 2 == 1 else TABLE_ALT
        set_row_h(ws, R, 20)
        fill(ws, R, 2, 8, bg, b_thin())
        W(ws, R, 2, sr, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
        mg(ws, R, 3, R, 4)
        W(ws, R, 3, f"  {desc}", bold=True, sz=9, color=DARK_TEXT, bg=bg, bdr=b_thin())
        W(ws, R, 5, "Main Rack", sz=9, color=NAVY_MID, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 6, rack["main_qty"], sz=9, bg=bg, ha="center", bdr=b_thin())
        W(ws, R, 7, round(main_price, 2), sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        W(ws, R, 8, round(main_total, 2), bold=True, sz=9, color=NAVY, bg=bg,
          ha="right", fmt="#,##0.00", bdr=b_thin())
        R += 1; sr += 1

        if rack["addon_qty"] > 0:
            addon_total  = addon_price * rack["addon_qty"]
            total_basic += addon_total
            bg = WHITE if sr % 2 == 1 else TABLE_ALT
            set_row_h(ws, R, 20)
            fill(ws, R, 2, 8, bg, b_thin())
            W(ws, R, 2, sr, sz=9, color=MID_TEXT, bg=bg, ha="center", bdr=b_thin())
            mg(ws, R, 3, R, 4)
            W(ws, R, 3, f"  {desc}", sz=9, color=MID_TEXT, bg=bg, bdr=b_thin())
            W(ws, R, 5, "Add-on Rack", sz=9, color=ORANGE, bg=bg, ha="center", bdr=b_thin())
            W(ws, R, 6, rack["addon_qty"], sz=9, bg=bg, ha="center", bdr=b_thin())
            W(ws, R, 7, round(addon_price, 2), sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
            W(ws, R, 8, round(addon_total, 2), bold=True, sz=9, color=NAVY, bg=bg,
              ha="right", fmt="#,##0.00", bdr=b_thin())
            R += 1; sr += 1

    # Subtotal bar
    set_row_h(ws, R, 22)
    fill(ws, R, 2, 8, LIGHT_BLUE)
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
    set_row_h(ws, R, 10); R += 1

    # ── PRICING SUMMARY ──────────────────────────────────────────────────────
    gst   = round(total_basic * 0.18, 2)
    grand = round(total_basic + gst, 2)

    def price_row(lbl, val, bg=WHITE, bold=False, vc=DARK_TEXT, is_text=False):
        nonlocal R
        set_row_h(ws, R, 20)
        fill(ws, R, 5, 8, bg)
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
        c_v.alignment = Alignment(horizontal="center" if is_text else "right",
                                   vertical="center", indent=1)
        c_v.fill = PatternFill("solid", fgColor=bg)
        c_v.border = b_thin()
        if not is_text:
            c_v.number_format = "#,##0.00"
        R += 1

    price_row("Basic Amount (₹)",    round(total_basic, 2), WHITE, True,  NAVY_HEADER)
    price_row("Freight Charges",     "Inclusive",           CREAM, False, ORANGE, True)
    price_row("Erection Charges",    "Inclusive",           CREAM, False, ORANGE, True)
    price_row("GST @ 18% (₹)",       gst,                   WHITE, False, MID_TEXT)

    # Grand Total row
    set_row_h(ws, R, 28)
    fill(ws, R, 2, 8, ORANGE)
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
    set_row_h(ws, R, 12); R += 1

    # ── TERMS & BANK ─────────────────────────────────────────────────────────
    set_row_h(ws, R, 20)
    fill(ws, R, 2, 8, NAVY_MID)
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
        c2.fill = PatternFill("solid", fgColor=CREAM)
        c2.border = b_thin()
        R += 1

    set_row_h(ws, R, 8); R += 1

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

    # Footer
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
def build_bom_sheet(ws, client, offer_no, date_obj, rack_data, rate_per_kg):
    ws.sheet_view.showGridLines = False

    # 14 columns: A=margin, B..M=data, N=margin
    # B=SR, C=Component, D=Section, E=Length, F=Thick, G=Wt/pcs,
    # H=Qty/Main, I=Qty/Addon, J=Main Total, K=Addon Total, L=UDL(kg/m²), M=Load/Level(kg)
    col_w = {1:1.5, 2:4, 3:22, 4:12, 5:10, 6:10, 7:11,
             8:10, 9:10, 10:13, 11:13, 12:12, 13:14, 14:1.5}
    for col, w in col_w.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    R = 1

    # Top accent
    set_row_h(ws, R, 5); fill(ws, R, 1, 14, ORANGE); R += 1

    # Header
    set_row_h(ws, R, 52); fill(ws, R, 1, 14, NAVY_HEADER)
    mg(ws, R, 2, R, 13)
    c = ws.cell(row=R, column=2)
    c.value = "BRIJ INDUSTRIES  —  BILL OF MATERIALS"
    c.font = Font(name="Calibri", size=17, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    R += 1

    set_row_h(ws, R, 14); fill(ws, R, 1, 14, NAVY_LIGHT)
    mg(ws, R, 2, R, 13)
    W(ws, R, 2, f"DSS Dolphin Storage Solutions  |  Offer No: {offer_no}  |  Customer: {client}  |  Date: {date_obj.strftime('%d %B %Y')}",
      sz=9, color=WHITE, bg=NAVY_LIGHT, ha="center")
    R += 1

    set_row_h(ws, R, 5); fill(ws, R, 1, 14, ORANGE); R += 1
    set_row_h(ws, R, 8); R += 1

    # Track totals for grand summary
    grand_main_wt  = 0.0
    grand_addon_wt = 0.0
    grand_total_val = 0.0

    for rack in rack_data:
        comp = calc_components(rack)

        # ── Module title bar ─────────────────────────────────────────────────
        set_row_h(ws, R, 22)
        mg(ws, R, 2, R, 13)
        c = ws.cell(row=R, column=2)
        # Fix: "MODULE A" not "A MODULE A"
        c.value = f"  MODULE {rack['module']}   |   Main Racks: {rack['main_qty']}   |   Add-on Racks: {rack['addon_qty']}   |   Levels: {rack['levels']}"
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.fill = PatternFill("solid", fgColor=NAVY_MID)
        c.border = b_med(NAVY_MID)
        R += 1

        # ── Column headers ───────────────────────────────────────────────────
        set_row_h(ws, R, 32)
        fill(ws, R, 2, 13, NAVY_HEADER)
        BOM_H = [
            (2,  "SR.",              "center"),
            (3,  "COMPONENT",        "left"),
            (4,  "SECTION\nPROFILE", "center"),
            (5,  "LENGTH\n(mm)",     "center"),
            (6,  "THICK\n(mm)",      "center"),
            (7,  "WT / PCS\n(kg)",   "center"),
            (8,  "QTY /\nMAIN",      "center"),
            (9,  "QTY /\nADD-ON",    "center"),
            (10, "MAIN\nTOTAL (kg)", "center"),
            (11, "ADD-ON\nTOTAL (kg)","center"),
            (12, "UDL\n(kg/m²)",     "center"),
            (13, "LOAD /\nLEVEL (kg)","center"),
        ]
        for col, txt, al in BOM_H:
            c = ws.cell(row=R, column=col)
            c.value = txt
            c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
            c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
            c.border = b_thin()
        R += 1

        # ── Component rows ───────────────────────────────────────────────────
        def comp_row(sr_n, comp_name, section, length, thick, wt_each,
                     qty_main, qty_addon, udl, load_level, bg):
            nonlocal R
            set_row_h(ws, R, 19)
            fill(ws, R, 2, 13, bg, b_thin())
            main_tot  = round(wt_each * qty_main,  3)
            addon_tot = round(wt_each * qty_addon, 3)
            row_vals = [
                (2,  sr_n,              "center", None),
                (3,  f"  {comp_name}",  "left",   None),
                (4,  section,           "center", None),
                (5,  length,            "center", None),
                (6,  thick,             "center", None),
                (7,  wt_each,           "center", "#,##0.000"),
                (8,  qty_main,          "center", None),
                (9,  qty_addon,         "center", None),
                (10, main_tot,          "right",  "#,##0.000"),
                (11, addon_tot,         "right",  "#,##0.000"),
                (12, udl,               "center", "#,##0"),
                (13, load_level,        "center", "#,##0"),
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

        levels = rack["levels"]
        beam_type_label = rack["bt"].replace(" Beam","")
        if rack["bt"] == "Pipe Beam":
            beam_sec = f"{rack['bh']}×{rack['bw']} Pipe"
        else:
            beam_sec = f"{rack['bh']}×{rack['bw']} Roll"

        # Upright (per unit)
        comp_row(1, "Upright / Column",
                 f"{rack['uw']}×{rack['ud']} Box",
                 rack["ul"], rack["ut"], comp["u_wt"],
                 4, 2, "—", "—", WHITE)

        # Beam (total for all levels, hence multiplied)
        beam_wt_all = round(comp["b_wt"] * levels, 3)
        comp_row(2, f"Beam ({beam_type_label})  ×{levels} levels",
                 beam_sec,
                 rack["bl"], rack["bth"], beam_wt_all,
                 2, 1, comp["udl_kg_m2"], int(comp["load_per_level_kg"]), TABLE_ALT)

        # Deep bar
        comp_row(3, "Deep Bar (Shelf Support)",
                 "92 mm flat",
                 int(comp["d_len"]), rack["dth"], comp["d_wt"],
                 4, 2, "—", "—", WHITE)

        # Cross brace
        comp_row(4, "Cross Brace (Diagonal)",
                 "92 mm flat",
                 int(comp["c_len"]), rack["cth"], comp["c_wt"],
                 2, 1, "—", "—", TABLE_ALT)

        # ── Module summary ───────────────────────────────────────────────────
        set_row_h(ws, R, 7); R += 1

        def mod_summary(lbl, v_main, v_addon, bg, bold=False, vc=DARK_TEXT, nf="#,##0.00"):
            nonlocal R
            set_row_h(ws, R, 20)
            fill(ws, R, 2, 13, bg)
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

        main_u_price  = round(comp["main_wt"]  * rate_per_kg, 2)
        addon_u_price = round(comp["addon_wt"] * rate_per_kg, 2)
        total_m_wt    = round(comp["main_wt"]  * rack["main_qty"],  2)
        total_a_wt    = round(comp["addon_wt"] * rack["addon_qty"], 2)
        total_m_val   = round(main_u_price  * rack["main_qty"],  2)
        total_a_val   = round(addon_u_price * rack["addon_qty"], 2)

        grand_main_wt  += total_m_wt
        grand_addon_wt += total_a_wt
        grand_total_val += total_m_val + total_a_val

        mod_summary("Wt — Single Rack (kg)",
                    comp["main_wt"], comp["addon_wt"], LIGHT_BLUE)
        mod_summary(f"Total Wt — All Racks  [Main ×{rack['main_qty']}  |  Add-on ×{rack['addon_qty']}]",
                    total_m_wt, total_a_wt,
                    LIGHT_BLUE, bold=True, vc=NAVY_HEADER)
        mod_summary(f"Total Value (₹)  [Main ×{rack['main_qty']}  |  Add-on ×{rack['addon_qty']}]",
                    total_m_val, total_a_val,
                    NAVY_HEADER, bold=True, vc=WHITE, nf="₹ #,##0.00")

        set_row_h(ws, R, 14); R += 1

    # ── TOTAL TONNAGE SUMMARY ─────────────────────────────────────────────────
    set_row_h(ws, R, 6); fill(ws, R, 1, 14, ORANGE); R += 1
    set_row_h(ws, R, 22)
    mg(ws, R, 2, R, 13)
    c = ws.cell(row=R, column=2)
    c.value = "TOTAL TONNAGE & PRICING SUMMARY"
    c.font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    c.border = b_med(NAVY_HEADER)
    R += 1

    # Sub-headers for tonnage table
    set_row_h(ws, R, 24)
    fill(ws, R, 2, 13, NAVY)
    ton_hdrs = [
        (2,  "DESCRIPTION",           "left"),
        (6,  "ALL MAIN RACKS\nWT (kg)", "center"),
        (8,  "ALL ADD-ON RACKS\nWT (kg)", "center"),
        (10, "COMBINED\nWT (kg)",     "center"),
        (11, "COMBINED\nWT (MT)",     "center"),
        (12, "BASIC\nAMOUNT (₹)",    "center"),
    ]
    for col, txt, al in ton_hdrs:
        c = ws.cell(row=R, column=col)
        c.value = txt
        c.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        c.border = b_thin()
    R += 1

    # Per-module tonnage rows
    for idx, rack in enumerate(rack_data):
        comp = calc_components(rack)
        m_wt = round(comp["main_wt"]  * rack["main_qty"],  2)
        a_wt = round(comp["addon_wt"] * rack["addon_qty"], 2)
        comb = round(m_wt + a_wt, 2)
        val  = round((comp["main_wt"] * rack["main_qty"] + comp["addon_wt"] * rack["addon_qty"]) * rate_per_kg, 2)
        bg   = WHITE if idx % 2 == 0 else TABLE_ALT

        set_row_h(ws, R, 19)
        fill(ws, R, 2, 13, bg, b_thin())
        mg(ws, R, 2, R, 5)
        W(ws, R, 2, f"  MODULE {rack['module']}  (Main ×{rack['main_qty']} | Add-on ×{rack['addon_qty']})",
          bold=True, sz=9, color=NAVY_MID, bg=bg, bdr=b_thin())
        W(ws, R, 6, m_wt,    sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        W(ws, R, 8, a_wt,    sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        W(ws, R, 10, comb,   sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        W(ws, R, 11, round(comb/1000, 3), sz=9, bg=bg, ha="right", fmt="#,##0.000", bdr=b_thin())
        W(ws, R, 12, val,    sz=9, bg=bg, ha="right", fmt="#,##0.00", bdr=b_thin())
        R += 1

    # Grand totals row
    grand_comb = round(grand_main_wt + grand_addon_wt, 2)
    gst_val    = round(grand_total_val * 0.18, 2)
    inc_gst    = round(grand_total_val + gst_val, 2)

    set_row_h(ws, R, 22)
    fill(ws, R, 2, 13, ORANGE)
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
        (12, round(grand_total_val,2), "#,##0.00"),
    ]:
        cv = ws.cell(row=R, column=col)
        cv.value = val
        cv.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        cv.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cv.border = b_med(NAVY_HEADER)
        cv.number_format = nf
    R += 1

    # GST + Grand incl GST
    for lbl, val, bg_l, bg_v in [
        ("GST @ 18% (₹)",                gst_val,  LIGHT_BLUE, LIGHT_BLUE),
        ("GRAND TOTAL incl. GST (₹)",     inc_gst,  NAVY_HEADER, ORANGE),
    ]:
        set_row_h(ws, R, 20)
        fill(ws, R, 2, 13, bg_l)
        mg(ws, R, 2, R, 11)
        c = ws.cell(row=R, column=2)
        c.value = lbl
        c.font = Font(name="Calibri", size=10, bold=(bg_l == NAVY_HEADER),
                      color=WHITE if bg_l == NAVY_HEADER else NAVY_HEADER)
        c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        c.fill = PatternFill("solid", fgColor=bg_l)
        c.border = b_med(NAVY_MID)
        mg(ws, R, 12, R, 13)
        cv = ws.cell(row=R, column=12)
        cv.value = val
        cv.font = Font(name="Calibri", size=11, bold=True,
                       color=WHITE)
        cv.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        cv.fill = PatternFill("solid", fgColor=bg_v)
        cv.border = b_med(NAVY_HEADER)
        cv.number_format = "₹ #,##0.00"
        R += 1

    # Footer
    set_row_h(ws, R, 5); fill(ws, R, 1, 14, ORANGE); R += 1
    set_row_h(ws, R, 14); fill(ws, R, 1, 14, NAVY_HEADER)
    mg(ws, R, 2, R, 13)
    W(ws, R, 2, "BRIJ INDUSTRIES  |  DSS Dolphin Storage Solutions  |  www.brijindustries.in",
      sz=8, color="8AADDD", bg=NAVY_HEADER, ha="center", italic=True)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.print_area = f"A1:{get_column_letter(14)}{R}"


# ═══════════════════════════════════════════════════════════════════════════════
#  MASTER BUILD
# ═══════════════════════════════════════════════════════════════════════════════
def build_excel(client, product, offer_no, date_obj, project_name,
                rack_data, rate_per_kg, out_path="quotation.xlsx", logo_path=None):
    wb = Workbook()
    ws_q = wb.active
    ws_q.title = "Commercial Offer"
    ws_b = wb.create_sheet("Bill of Materials")

    total_basic, gst, grand = build_quotation_sheet(
        ws_q, client, product, offer_no, date_obj,
        project_name, rack_data, rate_per_kg, logo_path
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
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>🐬  QUOTATION GENERATOR</h1>
    <p>BRIJ INDUSTRIES — DSS Dolphin Storage Solutions  |  Modular Mezzanine & Racking Systems</p>
</div>
""", unsafe_allow_html=True)

# ── Logo upload ────────────────────────────────────────────────────────────────
with st.expander("🖼️ Upload Company Logo (optional — for Excel header)", expanded=False):
    logo_file = st.file_uploader("Upload logo PNG/JPG", type=["png","jpg","jpeg"])

logo_path = None
if logo_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tf:
        tf.write(logo_file.read())
        logo_path = tf.name
    st.success(f"✅ Logo uploaded: {logo_file.name}")

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
    with st.expander(f"Module {chr(65+i)} — Configuration", expanded=(i == 0)):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown("**Quantities**")
            main_qty  = st.number_input("Main Rack Qty",    key=f"mq{i}", value=10, min_value=0)
            addon_qty = st.number_input("Add-on Rack Qty",  key=f"aq{i}", value=5,  min_value=0)
            levels    = st.number_input("No. of Levels",    key=f"lv{i}", value=3,  min_value=1)
        with c2:
            st.markdown("**Upright Details**")
            uw  = st.number_input("Width (mm)",     key=f"uw{i}", value=80)
            ud  = st.number_input("Depth (mm)",     key=f"ud{i}", value=60)
            ul  = st.number_input("Height (mm)",    key=f"ul{i}", value=3000)
            ut  = st.number_input("Thickness (mm)", key=f"ut{i}", value=1.6, format="%.1f")
        with c3:
            st.markdown("**Beam Details**")
            bt  = st.selectbox("Beam Type", ["Pipe Beam","Roll Formed Beam"], key=f"bt{i}")
            bh  = st.number_input("Height (mm)",    key=f"bh{i}", value=100)
            bw  = st.number_input("Width (mm)",     key=f"bw{i}", value=50)
            bl  = st.number_input("Length (mm)",    key=f"bl{i}", value=2000)
            bth = st.number_input("Thickness (mm)", key=f"bth{i}", value=1.6, format="%.1f")
        with c4:
            st.markdown("**Rack / Bracing**")
            depth = st.number_input("Rack Depth (mm)",         key=f"dp{i}", value=800)
            dth   = st.number_input("Deep Bar Thick (mm)",     key=f"dt{i}", value=1.6, format="%.1f")
            gap   = st.number_input("Cross Brace Gap (mm)",    key=f"gp{i}", value=500)
            cth   = st.number_input("Cross Brace Thick (mm)",  key=f"ct{i}", value=1.6, format="%.1f")

        rack_data.append({
            "module": chr(65+i),
            "name":   f"MODULE {chr(65+i)}",
            "main_qty": main_qty, "addon_qty": addon_qty, "levels": levels,
            "uw": uw, "ud": ud, "ul": ul, "ut": ut,
            "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
            "depth": depth, "dth": dth, "gap": gap, "cth": cth,
        })

st.divider()

# ── Live preview ──────────────────────────────────────────────────────────────
if rack_data:
    comp0    = calc_components(rack_data[0])
    prev_tot = sum(
        calc_components(r)["main_wt"] * rate_per_kg * r["main_qty"] +
        calc_components(r)["addon_wt"] * rate_per_kg * r["addon_qty"]
        for r in rack_data
    )
    prev_gst   = prev_tot * 0.18
    prev_grand = prev_tot + prev_gst
    all_wt     = sum(
        calc_components(r)["main_wt"] * r["main_qty"] +
        calc_components(r)["addon_wt"] * r["addon_qty"]
        for r in rack_data
    )

    p1,p2,p3,p4,p5,p6 = st.columns(6)
    p1.metric("Main Rack Wt",   f"{comp0['main_wt']:.2f} kg")
    p2.metric("Add-on Rack Wt", f"{comp0['addon_wt']:.2f} kg")
    p3.metric("Total Tonnage",  f"{all_wt/1000:.3f} MT")
    p4.metric("Basic Amount",   f"₹{prev_tot:,.0f}")
    p5.metric("GST (18%)",      f"₹{prev_gst:,.0f}")
    p6.metric("Grand Total",    f"₹{prev_grand:,.0f}")

st.divider()

# ── Generate ──────────────────────────────────────────────────────────────────
if st.button("🐬  GENERATE QUOTATION + BOM", type="primary", use_container_width=True):
    safe  = client.replace(" ","_").replace("/","-")
    fname = f"{safe}_Offer_{offer_no.replace('/','-')}.xlsx"
    out   = os.path.join(tempfile.gettempdir(), fname)

    try:
        basic, gst, grand = build_excel(
            client, product, offer_no, date, project_name,
            rack_data, rate_per_kg, out_path=out, logo_path=logo_path
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
