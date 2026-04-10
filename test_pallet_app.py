import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image
from io import BytesIO
import datetime
import base64

# ---------------- FUNCTIONS ---------------- #

def weight(length, width, thickness):
    return (length * width * thickness * 7.85) / 1_000_000

def upright_sheet_size(w, d):
    return w + d + w + 15

def pipe_beam_sheet_size(h, w):
    return 2 * (h + w)

def roll_beam_sheet_size(h, w):
    return (4*h + 2*w) + 20

def deep_bar_size(depth):
    return depth - 65

def cross_length(d, gap, ud):
    return ((d - 50)**2 + gap**2)**0.5 + 50

def apply_header_style(cell, bg_color='366092'):
    """Apply professional header styling with blue background"""
    cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border
    return cell

def apply_title_style(cell, size=14):
    """Apply title styling"""
    cell.font = Font(name='Arial', size=size, bold=True, color='366092')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    return cell

def apply_number_style(cell):
    """Apply number/currency styling"""
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.font = Font(name='Arial', size=10)
    cell.number_format = '#,##0.00'
    return cell

def apply_currency_style(cell):
    """Apply currency styling"""
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.font = Font(name='Arial', size=10, bold=True)
    cell.number_format = '#,##0.00'
    return cell

def create_logo():
    """Create a simple logo as bytes (Dolphin Storage Solutions)"""
    # You can replace this with actual logo file
    # For now, creating a text-based logo in Excel
    return None

# ---------------- UI ---------------- #

st.set_page_config(page_title="Professional Quotation Generator", layout="wide")

st.title("🏗 PROFESSIONAL QUOTATION + BOM GENERATOR (MEZZANINE FLOOR STYLE)")

col1, col2 = st.columns(2)

with col1:
    client = st.text_input("Customer Name (M/S)", value="STYLE BAZAAR")
    product = st.text_input("Product", value="MODULAR MEZZANINE FLOOR")
    offer_no = st.text_input("Offer No", value="DSS-IV/25-26/0712")
    
with col2:
    date = st.date_input("Date", datetime.date.today())
    project_name = st.text_input("Project Name", value="MODULE MEZZANINE FLOOR")
    
st.divider()

# Module/Rack Types
st.subheader("📦 MODULE / RACK CONFIGURATIONS")
rack_types = st.number_input("Number of Module/Rack Types", min_value=1, value=1)

rack_data = []

for i in range(int(rack_types)):
    st.markdown(f"### 🔹 MODULE {chr(65+i)} - Rack {i+1}")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        name = st.text_input(f"Module/Rack Name", key=f"name{i}", value=f"MODULE {chr(65+i)}")
        main_qty = st.number_input(f"Main Qty", key=f"m{i}", value=1)
        addon_qty = st.number_input(f"Addon Qty", key=f"a{i}", value=0)
        levels = st.number_input(f"Levels", key=f"l{i}", value=3)
        
    with col2:
        uw = st.number_input(f"Upright Width (mm)", key=f"uw{i}", value=80)
        ud = st.number_input(f"Upright Depth (mm)", key=f"ud{i}", value=60)
        ul = st.number_input(f"Upright Length (mm)", key=f"ul{i}", value=3000)
        ut = st.number_input(f"Upright Thickness (mm)", key=f"ut{i}", value=1.6)
        
    with col3:
        bt = st.selectbox(f"Beam Type", ["Pipe", "Roll"], key=f"bt{i}")
        bh = st.number_input(f"Beam Height (mm)", key=f"bh{i}", value=100)
        bw = st.number_input(f"Beam Width (mm)", key=f"bw{i}", value=50)
        bl = st.number_input(f"Beam Length (mm)", key=f"bl{i}", value=2000)
        bth = st.number_input(f"Beam Thickness (mm)", key=f"bth{i}", value=1.6)
    
    col4, col5 = st.columns(2)
    with col4:
        depth = st.number_input(f"Rack Depth (mm)", key=f"d{i}", value=800)
        dth = st.number_input(f"Deep Bar Thickness (mm)", key=f"dth{i}", value=1.6)
        
    with col5:
        gap = st.number_input(f"Gap (mm)", key=f"g{i}", value=500)
        method = st.number_input(f"Method (mm)", key=f"me{i}", value=50)
        cth = st.number_input(f"Cross Bar Thickness (mm)", key=f"cth{i}", value=1.6)
    
    rack_data.append({
        "module": chr(65+i),
        "name": name, "main_qty": main_qty, "addon_qty": addon_qty,
        "levels": levels, "uw": uw, "ud": ud, "ul": ul, "ut": ut,
        "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
        "depth": depth, "dth": dth, "gap": gap, "method": method, "cth": cth
    })
    st.divider()

rate = st.number_input("Rate per KG (₹)", value=85.00)

# ---------------- GENERATE ---------------- #

if st.button("🎯 GENERATE PROFESSIONAL QUOTATION", type="primary", use_container_width=True):
    
    # ================= CREATE QUOTATION WORKBOOK ================= #
    
    wb = Workbook()
    ws = wb.active
    ws.title = "COMMERCIAL OFFER"
    
    # Set default column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    
    # ================= HEADER SECTION ================= #
    
    # Row 1: Company Name with Logo
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = "DOLPHIN STORAGE SOLUTIONS"
    title_cell.font = Font(name='Arial', size=18, bold=True, color='366092')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells('A2:F2')
    sub_cell = ws['A2']
    sub_cell.value = "BRIJ INDUSTRIES"
    sub_cell.font = Font(name='Arial', size=12, bold=True)
    sub_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Row 3: Address
    ws.merge_cells('A3:F3')
    addr_cell = ws['A3']
    addr_cell.value = "ADDRESS:- 86/3/1 ROAD NO 7, MUNDKA INDUSTRIAL AREA SOUTH DELHI 110041"
    addr_cell.font = Font(name='Arial', size=9)
    addr_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Row 4: Contact Details
    ws.merge_cells('A4:F4')
    contact_cell = ws['A4']
    contact_cell.value = "E-MAIL:- Brijindustries09@rediffmail.com | WEBSITE:- WWW.BRIJINDUSTRIES.IN | MOBILE:- +919625589161, +919811096149"
    contact_cell.font = Font(name='Arial', size=9)
    contact_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Row 5: GST
    ws.merge_cells('A5:F5')
    gst_cell = ws['A5']
    gst_cell.value = "GST NO:- 07AAMFB6403G1ZM"
    gst_cell.font = Font(name='Arial', size=9, bold=True)
    gst_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Empty row
    ws.row_dimensions[6].height = 5
    
    # ================= CUSTOMER DETAILS ================= #
    
    details_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Row 7: M/S
    ws.merge_cells('A7:B7')
    ms_label = ws['A7']
    ms_label.value = "M/S:-"
    ms_label.font = Font(name='Arial', size=10, bold=True)
    ms_label.alignment = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells('C7:F7')
    ms_value = ws['C7']
    ms_value.value = client
    ms_value.font = Font(name='Arial', size=10)
    ms_value.alignment = Alignment(horizontal='left', vertical='center')
    
    # Row 8: Product
    ws.merge_cells('A8:B8')
    prod_label = ws['A8']
    prod_label.value = "PRODUCT:-"
    prod_label.font = Font(name='Arial', size=10, bold=True)
    prod_label.alignment = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells('C8:F8')
    prod_value = ws['C8']
    prod_value.value = product
    prod_value.font = Font(name='Arial', size=10)
    prod_value.alignment = Alignment(horizontal='left', vertical='center')
    
    # Row 9: Date
    ws.merge_cells('A9:B9')
    date_label = ws['A9']
    date_label.value = "DATE:-"
    date_label.font = Font(name='Arial', size=10, bold=True)
    date_label.alignment = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells('C9:D9')
    date_value = ws['C9']
    date_value.value = date.strftime("%d-%m-%Y")
    date_value.font = Font(name='Arial', size=10)
    date_value.alignment = Alignment(horizontal='left', vertical='center')
    
    # Row 10: Offer No
    ws.merge_cells('E10:F10')
    offer_label = ws['E10']
    offer_label.value = "OFFER NO:-"
    offer_label.font = Font(name='Arial', size=10, bold=True)
    offer_label.alignment = Alignment(horizontal='right', vertical='center')
    
    ws.merge_cells('A10:D10')
    offer_value = ws['A10']
    offer_value.value = offer_no
    offer_value.font = Font(name='Arial', size=10)
    offer_value.alignment = Alignment(horizontal='left', vertical='center')
    
    # Apply borders to customer details
    for row in range(7, 11):
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = details_border
    
    # Empty row
    ws.row_dimensions[11].height = 5
    
    # ================= PRODUCT TABLE HEADER ================= #
    
    # Row 12: Product Title
    ws.merge_cells('A12:F12')
    product_title = ws['A12']
    product_title.value = project_name
    product_title.font = Font(name='Arial', size=12, bold=True, color='366092')
    product_title.alignment = Alignment(horizontal='center', vertical='center')
    
    # Row 13: Table Headers
    headers = ["S.NO", "DESCRIPTION", "UNO", "QNTY", "AMOUNT (₹)", ""]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=13, column=col_idx, value=header)
        apply_header_style(cell)
    
    # ================= CALCULATIONS & DATA ================= #
    
    current_row = 14
    sr = 1
    total_basic_amount = 0
    
    # Store module details for summary
    module_summaries = []
    
    for rack in rack_data:
        # Calculate weights
        uwid = upright_sheet_size(rack["uw"], rack["ud"])
        u_wt = weight(rack["ul"], uwid, rack["ut"])
        
        beam_wid = pipe_beam_sheet_size(rack["bh"], rack["bw"]) if rack["bt"] == "Pipe" else roll_beam_sheet_size(rack["bh"], rack["bw"])
        b_wt = weight(rack["bl"], beam_wid, rack["bth"])
        
        d_len = deep_bar_size(rack["depth"])
        d_wt = weight(d_len, 92, rack["dth"])
        
        c_len = cross_length(d_len, rack["gap"], rack["ud"])
        c_wt = weight(c_len, 92, rack["cth"])
        
        components = [
            ("Upright", u_wt, 4, 2),
            ("Beam", b_wt, 2 * rack["levels"], 2 * rack["levels"]),
            ("Deep Bar", d_wt, 4, 2),
            ("Cross Bar", c_wt, 2, 1)
        ]
        
        rack_total = 0
        for comp, wt, m_qty, a_qty in components:
            total_qty = (m_qty * rack["main_qty"]) + (a_qty * rack["addon_qty"])
            total_wt = wt * total_qty
            rack_total += total_wt
        
        amount = rack_total * rate
        total_basic_amount += amount
        
        # Add module header row
        ws.merge_cells(f'A{current_row}:F{current_row}')
        module_cell = ws.cell(row=current_row, column=1, value=f"{rack['module']} - {rack['name']}")
        module_cell.font = Font(name='Arial', size=10, bold=True, color='366092')
        module_cell.fill = PatternFill(start_color='E6F0FA', end_color='E6F0FA', fill_type='solid')
        module_cell.alignment = Alignment(horizontal='center', vertical='center')
        current_row += 1
        
        # Add component rows
        for comp, wt, m_qty, a_qty in components:
            total_qty = (m_qty * rack["main_qty"]) + (a_qty * rack["addon_qty"])
            total_wt = wt * total_qty
            
            ws.cell(row=current_row, column=1, value=sr)
            ws.cell(row=current_row, column=2, value=comp)
            ws.cell(row=current_row, column=3, value="NOS")
            ws.cell(row=current_row, column=4, value=total_qty)
            ws.cell(row=current_row, column=5, value=round(amount, 2))
            
            # Apply number styling
            apply_number_style(ws.cell(row=current_row, column=4))
            apply_number_style(ws.cell(row=current_row, column=5))
            
            current_row += 1
            sr += 1
        
        # Add module total row
        ws.cell(row=current_row, column=4, value=f"Total {rack['module']}")
        ws.cell(row=current_row, column=5, value=round(amount, 2))
        ws.cell(row=current_row, column=4).font = Font(bold=True)
        ws.cell(row=current_row, column=5).font = Font(bold=True)
        apply_number_style(ws.cell(row=current_row, column=5))
        current_row += 1
        
        module_summaries.append({
            "module": rack['module'],
            "name": rack['name'],
            "amount": amount
        })
        
        current_row += 1  # Empty row between modules
    
    # ================= SUMMARY SECTION ================= #
    
    current_row += 1
    
    # Basic Amount
    ws.cell(row=current_row, column=4, value="BASIC AMOUNT")
    ws.cell(row=current_row, column=5, value=round(total_basic_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True)
    apply_currency_style(ws.cell(row=current_row, column=5))
    
    current_row += 1
    
    # Freight Charge
    ws.cell(row=current_row, column=4, value="FREIGHT CHARGE")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(bold=True)
    
    current_row += 1
    
    # Erection Charge
    ws.cell(row=current_row, column=4, value="ERECTION CHARGE")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(bold=True)
    
    current_row += 1
    
    # TOTAL AMOUNT
    ws.cell(row=current_row, column=4, value="TOTAL AMOUNT")
    ws.cell(row=current_row, column=5, value=round(total_basic_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True, size=11)
    apply_currency_style(ws.cell(row=current_row, column=5))
    
    current_row += 1
    
    # GST 18%
    gst_amount = total_basic_amount * 0.18
    ws.cell(row=current_row, column=4, value="GST 18%")
    ws.cell(row=current_row, column=5, value=round(gst_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True)
    apply_currency_style(ws.cell(row=current_row, column=5))
    
    current_row += 1
    
    # GRAND TOTAL
    grand_total = total_basic_amount + gst_amount
    ws.cell(row=current_row, column=4, value="GRAND TOTAL")
    ws.cell(row=current_row, column=5, value=round(grand_total, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True, size=12, color='FF0000')
    ws.cell(row=current_row, column=5).font = Font(bold=True, size=12, color='FF0000')
    apply_currency_style(ws.cell(row=current_row, column=5))
    
    # Apply borders to summary section
    for row in range(current_row-6, current_row+1):
        for col in [4, 5]:
            ws.cell(row=row, column=col).border = details_border
    
    # ================= FOOTER NOTES ================= #
    
    current_row += 2
    
    notes = [
        "NOTES:",
        "1. The above prices are exclusive of GST @18%",
        "2. Freight and installation charges are inclusive",
        "3. Payment Terms: 50% advance, 50% against delivery",
        "4. Delivery Period: 10-12 weeks from advance payment",
        "5. Warranty: 12 months from commissioning"
    ]
    
    for note in notes:
        ws.cell(row=current_row, column=1, value=note)
        if note == "NOTES:":
            ws.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1
    
    # Save file
    filename = f"{client.replace(' ', '_')}_COMMERCIAL_OFFER.xlsx"
    wb.save(filename)
    
    st.success("✅ PROFESSIONAL QUOTATION GENERATED SUCCESSFULLY!")
    st.balloons()
    
    # Download button
    with open(filename, "rb") as f:
        st.download_button(
            label="📊 DOWNLOAD COMMERCIAL OFFER (EXCEL)",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Display preview of modules
    st.subheader("📋 QUOTATION SUMMARY")
    st.write(f"**Customer:** {client}")
    st.write(f"**Offer No:** {offer_no}")
    st.write(f"**Date:** {date.strftime('%d-%m-%Y')}")
    
    for mod in module_summaries:
        st.write(f"**{mod['module']} - {mod['name']}:** ₹{mod['amount']:,.2f}")
    
    st.write(f"---")
    st.write(f"**Grand Total:** ₹{grand_total:,.2f} (Including GST 18%)")
