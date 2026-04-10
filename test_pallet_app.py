import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, GradientFill
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import datetime
import io
import base64
from PIL import Image as PILImage
import requests

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

def calculate_rack_weight(rack):
    """Calculate total weight for a rack"""
    uwid = upright_sheet_size(rack["uw"], rack["ud"])
    u_wt = weight(rack["ul"], uwid, rack["ut"])
    
    beam_wid = pipe_beam_sheet_size(rack["bh"], rack["bw"]) if rack["bt"] == "Pipe" else roll_beam_sheet_size(rack["bh"], rack["bw"])
    b_wt = weight(rack["bl"], beam_wid, rack["bth"])
    
    d_len = deep_bar_size(rack["depth"])
    d_wt = weight(d_len, 92, rack["dth"])
    
    c_len = cross_length(d_len, rack["gap"], rack["ud"])
    c_wt = weight(c_len, 92, rack["cth"])
    
    # Total weight per rack set
    total_weight = (u_wt * 4) + (b_wt * 2 * rack["levels"]) + (d_wt * 4) + (c_wt * 2)
    return total_weight

def create_logo_image():
    """Create a professional logo image with DSS and Dolphin Storage Solutions"""
    # Create a PIL Image for the logo
    img = PILImage.new('RGB', (300, 100), color='#1B3A5C')
    from PIL import ImageDraw, ImageFont
    draw = ImageDraw.Draw(img)
    
    # Draw dolphin-like shape (simplified)
    draw.ellipse([20, 20, 80, 80], fill='#FF6B35', outline='#FFFFFF', width=3)
    draw.arc([30, 35, 70, 75], start=0, end=180, fill='#FFFFFF', width=4)
    
    # Draw text
    draw.text((100, 25), "DSS", fill='#FF6B35', font=None)
    draw.text((100, 50), "DOLPHIN", fill='#FFFFFF', font=None)
    draw.text((100, 75), "STORAGE SOLUTIONS", fill='#FF6B35', font=None)
    
    # Convert to bytes
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    return img_bytes

def apply_header_style(cell, bg_color='1B3A5C'):
    """Apply professional header styling with dark blue background"""
    cell.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return cell

def apply_orange_accent(cell):
    """Apply orange accent styling"""
    cell.font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color='FF6B35', end_color='FF6B35', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    return cell

def apply_light_blue_bg(cell):
    """Apply light blue background for alternating rows"""
    cell.fill = PatternFill(start_color='E8F0FE', end_color='E8F0FE', fill_type='solid')
    return cell

def apply_border(cell, border_style='thin'):
    """Apply border to cell"""
    border = Border(
        left=Side(style=border_style, color='CCCCCC'),
        right=Side(style=border_style, color='CCCCCC'),
        top=Side(style=border_style, color='CCCCCC'),
        bottom=Side(style=border_style, color='CCCCCC')
    )
    cell.border = border
    return cell

# ---------------- UI ---------------- #

st.set_page_config(page_title="Professional Quotation Generator", layout="wide")

# Custom CSS for better UI
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #1B3A5C 0%, #FF6B35 100%);
    }
    .main-header {
        background: linear-gradient(135deg, #1B3A5C, #FF6B35);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🏗 PROFESSIONAL QUOTATION GENERATOR</h1><p>Modular Mezzanine Floor & Racking Systems</p></div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    client = st.text_input("🏢 Customer Name (M/S)", value="STYLE BAZAAR")
    product = st.text_input("📦 Product", value="MODULAR MEZZANINE FLOOR")
    offer_no = st.text_input("📄 Offer No", value="DSS-IV/25-26/0712")
    
with col2:
    date = st.date_input("📅 Date", datetime.date.today())
    project_name = st.text_input("🏗️ Project Name", value="MODULE MEZZANINE FLOOR")
    color_scheme = st.selectbox("🎨 Color Scheme", ["Blue & Orange (Default)", "Blue & Gold", "Navy & Coral"])

st.divider()

# Module/Rack Types
st.subheader("📦 RACK CONFIGURATIONS")
st.info("💡 Each rack type will appear as a separate module in the quotation")

rack_types = st.number_input("Number of Rack Types", min_value=1, value=1)

rack_data = []

for i in range(int(rack_types)):
    with st.expander(f"🔹 RACK TYPE {chr(65+i)} - Configuration", expanded=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            name = st.text_input(f"Rack Name", key=f"name{i}", value=f"MODULE {chr(65+i)}")
            main_qty = st.number_input(f"Main Rack Qty", key=f"m{i}", value=10, min_value=0)
            addon_qty = st.number_input(f"Add-on Rack Qty", key=f"a{i}", value=5, min_value=0)
            levels = st.number_input(f"Number of Levels", key=f"l{i}", value=3, min_value=1)
            
        with col2:
            uw = st.number_input(f"Upright Width (mm)", key=f"uw{i}", value=80)
            ud = st.number_input(f"Upright Depth (mm)", key=f"ud{i}", value=60)
            ul = st.number_input(f"Upright Length (mm)", key=f"ul{i}", value=3000)
            ut = st.number_input(f"Upright Thickness (mm)", key=f"ut{i}", value=1.6, format="%.1f")
            
        with col3:
            bt = st.selectbox(f"Beam Type", ["Pipe Beam", "Roll Formed Beam"], key=f"bt{i}")
            bh = st.number_input(f"Beam Height (mm)", key=f"bh{i}", value=100)
            bw = st.number_input(f"Beam Width (mm)", key=f"bw{i}", value=50)
            bl = st.number_input(f"Beam Length (mm)", key=f"bl{i}", value=2000)
            bth = st.number_input(f"Beam Thickness (mm)", key=f"bth{i}", value=1.6, format="%.1f")
    
    rack_data.append({
        "module": chr(65+i),
        "name": name, 
        "main_qty": main_qty, 
        "addon_qty": addon_qty,
        "levels": levels, 
        "uw": uw, "ud": ud, "ul": ul, "ut": ut,
        "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
        "depth": 800, "dth": 1.6, "gap": 500, "method": 50, "cth": 1.6
    })

rate_per_kg = st.number_input("💰 Rate per KG (₹)", value=85.00, min_value=0.0, format="%.2f")

# Calculate rates for main and add-on racks
st.divider()
st.subheader("💰 PRICE CALCULATION")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 🏗️ Main Rack Price")
    st.caption("Complete rack system including all components")
    
with col2:
    st.markdown("### 📦 Add-on Rack Price")
    st.caption("Extension rack (without end frames)")

# Calculate sample prices for preview
if rack_data:
    sample_rack = rack_data[0]
    sample_weight = calculate_rack_weight(sample_rack)
    main_rack_price = sample_weight * rate_per_kg
    addon_rack_price = (sample_weight * 0.7) * rate_per_kg  # Add-on is ~70% of main
    
    col1.metric("Estimated Main Rack Price", f"₹{main_rack_price:,.2f}")
    col2.metric("Estimated Add-on Rack Price", f"₹{addon_rack_price:,.2f}")

# ---------------- GENERATE ---------------- #

if st.button("🎯 GENERATE PROFESSIONAL QUOTATION", type="primary", use_container_width=True):
    
    # ================= CREATE QUOTATION WORKBOOK ================= #
    
    wb = Workbook()
    ws = wb.active
    ws.title = "COMMERCIAL OFFER"
    
    # Set column widths
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    
    # ================= HEADER SECTION WITH LOGO ================= #
    
    # Create logo area (merge cells for logo and company name)
    ws.merge_cells('A1:C1')
    ws.merge_cells('D1:F1')
    
    # Logo cell
    logo_cell = ws['A1']
    logo_cell.value = "🐬 DSS"
    logo_cell.font = Font(name='Segoe UI', size=20, bold=True, color='FF6B35')
    logo_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    # Company name cell
    company_cell = ws['D1']
    company_cell.value = "DOLPHIN STORAGE SOLUTIONS"
    company_cell.font = Font(name='Segoe UI', size=16, bold=True, color='1B3A5C')
    company_cell.alignment = Alignment(horizontal='right', vertical='center')
    
    # Subtitle row
    ws.merge_cells('A2:F2')
    subtitle_cell = ws['A2']
    subtitle_cell.value = "BRIJ INDUSTRIES"
    subtitle_cell.font = Font(name='Segoe UI', size=12, bold=True, color='666666')
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Contact info with orange accent
    ws.merge_cells('A3:F3')
    contact_cell = ws['A3']
    contact_cell.value = "📍 86/3/1 ROAD NO 7, MUNDKA INDUSTRIAL AREA, SOUTH DELHI - 110041"
    contact_cell.font = Font(name='Segoe UI', size=9)
    contact_cell.alignment = Alignment(horizontal='center', vertical='center')
    contact_cell.fill = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
    
    ws.merge_cells('A4:F4')
    email_cell = ws['A4']
    email_cell.value = "✉️ brijindustries09@rediffmail.com  |  🌐 WWW.BRIJINDUSTRIES.IN  |  📞 +91 9625589161, +91 9811096149"
    email_cell.font = Font(name='Segoe UI', size=9)
    email_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells('A5:F5')
    gst_cell = ws['A5']
    gst_cell.value = "✅ GST NO: 07AAMFB6403G1ZM"
    gst_cell.font = Font(name='Segoe UI', size=9, bold=True)
    gst_cell.alignment = Alignment(horizontal='center', vertical='center')
    gst_cell.fill = PatternFill(start_color='FF6B35', end_color='FF6B35', fill_type='solid')
    gst_cell.font = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
    
    # Decorative line
    ws.row_dimensions[6].height = 5
    
    # ================= CUSTOMER DETAILS CARD ================= #
    
    # Create a bordered box for customer details
    for row in range(7, 12):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(
                left=Side(style='thin', color='1B3A5C'),
                right=Side(style='thin', color='1B3A5C'),
                top=Side(style='thin', color='1B3A5C') if row == 7 else None,
                bottom=Side(style='thin', color='1B3A5C') if row == 11 else None
            )
    
    # Fill background
    for row in range(7, 12):
        for col in range(1, 7):
            ws.cell(row=row, column=col).fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
    
    # Customer details
    details = [
        ("M/S:", client, "DATE:", date.strftime("%d-%m-%Y")),
        ("PRODUCT:", product, "OFFER NO:", offer_no),
        ("PROJECT:", project_name, "", "")
    ]
    
    row = 7
    for detail in details:
        # Left label
        cell1 = ws.cell(row=row, column=1, value=detail[0])
        cell1.font = Font(bold=True, color='1B3A5C')
        cell1.alignment = Alignment(horizontal='left', vertical='center')
        
        # Left value
        cell2 = ws.cell(row=row, column=2, value=detail[1])
        cell2.font = Font(color='333333')
        cell2.alignment = Alignment(horizontal='left', vertical='center')
        
        # Right label
        if detail[2]:
            cell3 = ws.cell(row=row, column=4, value=detail[2])
            cell3.font = Font(bold=True, color='1B3A5C')
            cell3.alignment = Alignment(horizontal='left', vertical='center')
        
        # Right value
        if detail[3]:
            cell4 = ws.cell(row=row, column=5, value=detail[3])
            cell4.font = Font(color='333333')
            cell4.alignment = Alignment(horizontal='left', vertical='center')
        
        row += 1
    
    # Empty row
    ws.row_dimensions[12].height = 10
    
    # ================= PRODUCT TABLE ================= #
    
    # Section header with orange accent
    ws.merge_cells('A13:F13')
    section_header = ws['A13']
    section_header.value = "📊 COMMERCIAL OFFER DETAILS"
    section_header.font = Font(name='Segoe UI', size=14, bold=True, color='FFFFFF')
    section_header.fill = PatternFill(start_color='FF6B35', end_color='FF6B35', fill_type='solid')
    section_header.alignment = Alignment(horizontal='center', vertical='center')
    
    # Table headers
    headers = ["#", "DESCRIPTION", "TYPE", "QTY", "UNIT PRICE (₹)", "TOTAL (₹)"]
    header_row = 14
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1B3A5C', end_color='1B3A5C', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        apply_border(cell)
    
    # ================= CALCULATIONS ================= #
    
    current_row = 15
    sr = 1
    total_basic_amount = 0
    
    for rack in rack_data:
        # Calculate rack weight and price
        rack_weight = calculate_rack_weight(rack)
        main_rack_price = rack_weight * rate_per_kg
        addon_rack_price = (rack_weight * 0.7) * rate_per_kg  # Add-on is 70% of main
        
        # Main Rack row
        main_total = main_rack_price * rack["main_qty"]
        total_basic_amount += main_total
        
        # Alternating row colors
        bg_color = 'F0F7FF' if sr % 2 == 1 else 'FFFFFF'
        
        ws.cell(row=current_row, column=1, value=sr)
        ws.cell(row=current_row, column=2, value=f"{rack['module']} - {rack['name']}")
        ws.cell(row=current_row, column=3, value="MAIN RACK")
        ws.cell(row=current_row, column=4, value=rack["main_qty"])
        ws.cell(row=current_row, column=5, value=round(main_rack_price, 2))
        ws.cell(row=current_row, column=6, value=round(main_total, 2))
        
        # Apply styling
        for col in range(1, 7):
            cell = ws.cell(row=current_row, column=col)
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
            apply_border(cell)
            if col in [4, 5, 6]:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0.00'
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        current_row += 1
        sr += 1
        
        # Add-on Rack row (if quantity > 0)
        if rack["addon_qty"] > 0:
            addon_total = addon_rack_price * rack["addon_qty"]
            total_basic_amount += addon_total
            
            bg_color = 'F0F7FF' if sr % 2 == 1 else 'FFFFFF'
            
            ws.cell(row=current_row, column=1, value=sr)
            ws.cell(row=current_row, column=2, value=f"{rack['module']} - {rack['name']} (Extension)")
            ws.cell(row=current_row, column=3, value="ADD-ON RACK")
            ws.cell(row=current_row, column=4, value=rack["addon_qty"])
            ws.cell(row=current_row, column=5, value=round(addon_rack_price, 2))
            ws.cell(row=current_row, column=6, value=round(addon_total, 2))
            
            for col in range(1, 7):
                cell = ws.cell(row=current_row, column=col)
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
                apply_border(cell)
                if col in [4, 5, 6]:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.number_format = '#,##0.00'
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            current_row += 1
            sr += 1
    
    # Subtotal row
    current_row += 1
    subtotal_row = current_row
    
    ws.merge_cells(f'A{subtotal_row}:D{subtotal_row}')
    subtotal_label = ws.cell(row=subtotal_row, column=1, value="SUB TOTAL")
    subtotal_label.font = Font(bold=True, size=11)
    subtotal_label.alignment = Alignment(horizontal='right', vertical='center')
    
    subtotal_value = ws.cell(row=subtotal_row, column=6, value=round(total_basic_amount, 2))
    subtotal_value.font = Font(bold=True, size=11)
    subtotal_value.alignment = Alignment(horizontal='right', vertical='center')
    subtotal_value.number_format = '#,##0.00'
    
    for col in [1, 2, 3, 4, 5, 6]:
        cell = ws.cell(row=subtotal_row, column=col)
        cell.fill = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')
        apply_border(cell)
    
    # ================= SUMMARY SECTION WITH BOXES ================= #
    
    current_row += 2
    
    # Create summary box
    summary_start = current_row
    
    # Basic Amount
    ws.cell(row=current_row, column=4, value="BASIC AMOUNT")
    ws.cell(row=current_row, column=5, value=round(total_basic_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True, size=10)
    ws.cell(row=current_row, column=5).font = Font(bold=True, size=10)
    ws.cell(row=current_row, column=5).number_format = '#,##0.00'
    
    current_row += 1
    
    # Freight
    ws.cell(row=current_row, column=4, value="FREIGHT CHARGES")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(size=10)
    ws.cell(row=current_row, column=5).font = Font(size=10, italic=True)
    
    current_row += 1
    
    # Erection
    ws.cell(row=current_row, column=4, value="ERECTION CHARGES")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(size=10)
    ws.cell(row=current_row, column=5).font = Font(size=10, italic=True)
    
    current_row += 1
    
    # Total Amount
    ws.cell(row=current_row, column=4, value="TOTAL AMOUNT")
    ws.cell(row=current_row, column=5, value=round(total_basic_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True, size=11, color='1B3A5C')
    ws.cell(row=current_row, column=5).font = Font(bold=True, size=11, color='1B3A5C')
    ws.cell(row=current_row, column=5).number_format = '#,##0.00'
    
    current_row += 1
    
    # GST
    gst_amount = total_basic_amount * 0.18
    ws.cell(row=current_row, column=4, value="GST (18%)")
    ws.cell(row=current_row, column=5, value=round(gst_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(size=10)
    ws.cell(row=current_row, column=5).font = Font(size=10)
    ws.cell(row=current_row, column=5).number_format = '#,##0.00'
    
    current_row += 1
    
    # Grand Total with orange highlight
    grand_total = total_basic_amount + gst_amount
    ws.merge_cells(f'D{current_row}:E{current_row}')
    grand_cell = ws.cell(row=current_row, column=4, value=f"GRAND TOTAL: ₹{grand_total:,.2f}")
    grand_cell.font = Font(name='Segoe UI', size=14, bold=True, color='FFFFFF')
    grand_cell.fill = PatternFill(start_color='FF6B35', end_color='FF6B35', fill_type='solid')
    grand_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Apply borders to summary section
    for row in range(summary_start, current_row):
        for col in [4, 5]:
            apply_border(ws.cell(row=row, column=col))
    
    # ================= FOOTER NOTES ================= #
    
    current_row += 2
    
    footer_bg = PatternFill(start_color='F5F5F5', end_color='F5F5F5', fill_type='solid')
    
    notes = [
        "📋 TERMS & CONDITIONS",
        "✓ Payment Terms: 50% advance, 50% against delivery",
        "✓ Delivery Period: 10-12 weeks from advance payment",
        "✓ Warranty: 12 months from date of commissioning",
        "✓ GST @18% extra as applicable",
        "✓ Freight & Installation charges are inclusive",
        "",
        "🏦 BANK DETAILS",
        "Account Name: BRIJ INDUSTRIES",
        "Bank: HDFC Bank Ltd.",
        "Account No: 50100512345678",
        "IFSC: HDFC0001234",
        "",
        "For any queries, please contact our sales team.",
        "📞 +91 9625589161  |  ✉️ sales@brijindustries.in"
    ]
    
    for note in notes:
        if note.startswith("📋") or note.startswith("🏦"):
            ws.merge_cells(f'A{current_row}:F{current_row}')
            cell = ws.cell(row=current_row, column=1, value=note)
            cell.font = Font(bold=True, size=11, color='1B3A5C')
            cell.fill = footer_bg
            cell.alignment = Alignment(horizontal='left', vertical='center')
        elif note == "":
            pass
        else:
            ws.merge_cells(f'A{current_row}:F{current_row}')
            cell = ws.cell(row=current_row, column=1, value=f"  {note}")
            cell.font = Font(size=9)
            cell.fill = footer_bg
            cell.alignment = Alignment(horizontal='left', vertical='center')
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
    
    # Display preview
    st.subheader("📋 QUOTATION SUMMARY")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Customer", client)
        st.metric("Total Basic Amount", f"₹{total_basic_amount:,.2f}")
    with col2:
        st.metric("Offer No", offer_no)
        st.metric("GST (18%)", f"₹{gst_amount:,.2f}")
    with col3:
        st.metric("Date", date.strftime("%d-%m-%Y"))
        st.metric("GRAND TOTAL", f"₹{grand_total:,.2f}", delta="Including GST")
    
    st.divider()
    st.info("💡 The Excel file has been formatted with professional blue & orange color scheme, including borders, headers, and proper number formatting.")
