import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import datetime

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

def apply_header_style(cell):
    """Apply professional header styling"""
    cell.font = Font(name='Calibri', size=11, bold=True, color='000000')
    cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border
    return cell

def apply_title_style(cell):
    """Apply title styling"""
    cell.font = Font(name='Calibri', size=14, bold=True)
    cell.alignment = Alignment(horizontal='left', vertical='center')
    return cell

def apply_number_style(cell):
    """Apply number/currency styling"""
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.font = Font(name='Calibri', size=11)
    return cell

def apply_bold_style(cell):
    """Apply bold text style"""
    cell.font = Font(name='Calibri', size=11, bold=True)
    cell.alignment = Alignment(horizontal='right', vertical='center')
    return cell

def format_worksheet(ws, column_widths):
    """Auto-adjust column widths"""
    for col_idx, width in enumerate(column_widths, 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

# ---------------- UI ---------------- #

st.title("🏗 PROFESSIONAL QUOTATION + BOM GENERATOR")

client = st.text_input("Client Name")
offer_no = st.text_input("Offer No")

rack_types = st.number_input("No. of Rack Types", min_value=1)

rack_data = []

for i in range(int(rack_types)):
    st.header(f"Rack {i+1}")

    name = st.text_input("Rack Name", key=i)
    main_qty = st.number_input("Main Qty", key=f"m{i}")
    addon_qty = st.number_input("Addon Qty", key=f"a{i}")
    levels = st.number_input("Levels", key=f"l{i}")

    uw = st.number_input("Upright Width (mm)", key=f"uw{i}")
    ud = st.number_input("Upright Depth (mm)", key=f"ud{i}")
    ul = st.number_input("Upright Length (mm)", key=f"ul{i}")
    ut = st.number_input("Upright Thickness (mm)", key=f"ut{i}")

    bt = st.selectbox("Beam Type", ["Pipe", "Roll"], key=f"bt{i}")
    bh = st.number_input("Beam Height (mm)", key=f"bh{i}")
    bw = st.number_input("Beam Width (mm)", key=f"bw{i}")
    bl = st.number_input("Beam Length (mm)", key=f"bl{i}")
    bth = st.number_input("Beam Thickness (mm)", key=f"bth{i}")

    depth = st.number_input("Rack Depth (mm)", key=f"d{i}")
    dth = st.number_input("Deep Bar Thickness (mm)", key=f"dth{i}")

    gap = st.number_input("Gap (mm)", key=f"g{i}")
    method = st.number_input("Method (mm)", key=f"me{i}")
    cth = st.number_input("Cross Bar Thickness (mm)", key=f"cth{i}")

    rack_data.append({
        "name": name, "main_qty": main_qty, "addon_qty": addon_qty,
        "levels": levels, "uw": uw, "ud": ud, "ul": ul, "ut": ut,
        "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
        "depth": depth, "dth": dth, "gap": gap, "method": method, "cth": cth
    })

rate = st.number_input("Rate per KG (₹)", value=100)

# ---------------- GENERATE ---------------- #

if st.button("GENERATE FILES"):

    # ================= QUOTATION SHEET ================= #
    
    wb1 = Workbook()
    ws1 = wb1.active
    ws1.title = "Quotation"
    
    # Add company info with professional styling
    ws1.merge_cells('A1:F1')
    title_cell = ws1['A1']
    title_cell.value = "BRIJ INDUSTRIES"
    apply_title_style(title_cell)
    
    ws1.merge_cells('A2:F2')
    client_cell = ws1['A2']
    client_cell.value = f"M/S: {client}"
    client_cell.font = Font(name='Calibri', size=11)
    client_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    ws1.merge_cells('A3:F3')
    offer_cell = ws1['A3']
    offer_cell.value = f"OFFER NO: {offer_no}"
    offer_cell.font = Font(name='Calibri', size=11)
    offer_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    ws1.merge_cells('A4:F4')
    date_cell = ws1['A4']
    date_cell.value = f"DATE: {datetime.date.today()}"
    date_cell.font = Font(name='Calibri', size=11)
    date_cell.alignment = Alignment(horizontal='left', vertical='center')
    
    # Empty row
    ws1.append([])
    
    # Headers
    headers = ["S.No", "Description", "Unit", "Qty", "Price (₹)", "Amount (₹)"]
    ws1.append(headers)
    
    # Apply header styling
    for col_idx, header in enumerate(headers, 1):
        apply_header_style(ws1.cell(row=6, column=col_idx))
    
    # ================= CALCULATIONS ================= #
    
    sr = 1
    total = 0
    current_row = 7
    
    for r in rack_data:
        uwid = upright_sheet_size(r["uw"], r["ud"])
        u_wt = weight(r["ul"], uwid, r["ut"])
        
        beam_wid = pipe_beam_sheet_size(r["bh"], r["bw"]) if r["bt"] == "Pipe" else roll_beam_sheet_size(r["bh"], r["bw"])
        b_wt = weight(r["bl"], beam_wid, r["bth"])
        
        d_len = deep_bar_size(r["depth"])
        d_wt = weight(d_len, 92, r["dth"])
        
        c_len = cross_length(d_len, r["gap"], r["ud"])
        c_wt = weight(c_len, 92, r["cth"])
        
        components = [
            ("Upright", u_wt, 4, 2),
            ("Beam", b_wt, 2 * r["levels"], 2 * r["levels"]),
            ("Deep Bar", d_wt, 4, 2),
            ("Cross Bar", c_wt, 2, 1)
        ]
        
        rack_total = 0
        
        for comp, wt, m_qty, a_qty in components:
            main_wt = wt * m_qty
            addon_wt = wt * a_qty
            
            total_qty = (m_qty * r["main_qty"]) + (a_qty * r["addon_qty"])
            total_wt = wt * total_qty
            rack_total += total_wt
        
        amount = rack_total * rate
        total += amount
        
        # Add row to quotation
        row_data = [sr, f"{r['name']} Rack", "NOS", r["main_qty"] + r["addon_qty"], rate, round(amount, 2)]
        ws1.append(row_data)
        
        # Apply number styling
        for col_idx in [4, 5, 6]:
            apply_number_style(ws1.cell(row=current_row, column=col_idx))
        
        current_row += 1
        sr += 1
    
    # Add summary section
    current_row += 1
    
    # Basic Total
    ws1.cell(row=current_row, column=5, value="Basic")
    apply_bold_style(ws1.cell(row=current_row, column=5))
    ws1.cell(row=current_row, column=6, value=round(total, 2))
    apply_number_style(ws1.cell(row=current_row, column=6))
    
    current_row += 1
    
    # GST 18%
    gst = total * 0.18
    ws1.cell(row=current_row, column=5, value="GST 18%")
    apply_bold_style(ws1.cell(row=current_row, column=5))
    ws1.cell(row=current_row, column=6, value=round(gst, 2))
    apply_number_style(ws1.cell(row=current_row, column=6))
    
    current_row += 1
    
    # Grand Total
    ws1.cell(row=current_row, column=5, value="Grand Total")
    apply_bold_style(ws1.cell(row=current_row, column=5))
    ws1.cell(row=current_row, column=6, value=round(total + gst, 2))
    apply_number_style(ws1.cell(row=current_row, column=6))
    
    # Add border to summary section
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in range(current_row-2, current_row+1):
        for col in [5, 6]:
            ws1.cell(row=row, column=col).border = thin_border
    
    # Format columns
    format_worksheet(ws1, [8, 30, 10, 10, 15, 18])
    
    # ================= BOM SHEET ================= #
    
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "BOM"
    
    # BOM Headers
    bom_headers = ["Rack", "Component", "Single Main Qty", "Single Addon Qty", 
                   "Total Qty", "Unit Weight (kg)", "Main Weight (kg)", 
                   "Addon Weight (kg)", "Total Weight (kg)"]
    ws2.append(bom_headers)
    
    # Apply header styling
    for col_idx, header in enumerate(bom_headers, 1):
        apply_header_style(ws2.cell(row=1, column=col_idx))
    
    current_row = 2
    
    for r in rack_data:
        uwid = upright_sheet_size(r["uw"], r["ud"])
        u_wt = weight(r["ul"], uwid, r["ut"])
        
        beam_wid = pipe_beam_sheet_size(r["bh"], r["bw"]) if r["bt"] == "Pipe" else roll_beam_sheet_size(r["bh"], r["bw"])
        b_wt = weight(r["bl"], beam_wid, r["bth"])
        
        d_len = deep_bar_size(r["depth"])
        d_wt = weight(d_len, 92, r["dth"])
        
        c_len = cross_length(d_len, r["gap"], r["ud"])
        c_wt = weight(c_len, 92, r["cth"])
        
        components = [
            ("Upright", u_wt, 4, 2),
            ("Beam", b_wt, 2 * r["levels"], 2 * r["levels"]),
            ("Deep Bar", d_wt, 4, 2),
            ("Cross Bar", c_wt, 2, 1)
        ]
        
        for comp, wt, m_qty, a_qty in components:
            main_wt = wt * m_qty
            addon_wt = wt * a_qty
            
            total_qty = (m_qty * r["main_qty"]) + (a_qty * r["addon_qty"])
            total_wt = wt * total_qty
            
            row_data = [
                r["name"], comp,
                m_qty, a_qty, total_qty,
                round(wt, 2),
                round(main_wt, 2),
                round(addon_wt, 2),
                round(total_wt, 2)
            ]
            ws2.append(row_data)
            
            # Apply number styling to numeric columns
            for col_idx in [3, 4, 5, 6, 7, 8, 9]:
                apply_number_style(ws2.cell(row=current_row, column=col_idx))
            
            current_row += 1
    
    # Format BOM columns
    format_worksheet(ws2, [15, 15, 15, 15, 12, 15, 15, 15, 15])
    
    # SAVE FILES
    q_file = f"{client}_Quotation.xlsx"
    b_file = f"{client}_BOM.xlsx"
    
    wb1.save(q_file)
    wb2.save(b_file)
    
    st.success("✅ Excel Files Generated Successfully!")
    
    col1, col2 = st.columns(2)
    
    with col1:
        with open(q_file, "rb") as f:
            st.download_button(
                "📊 Download Quotation (Excel)",
                f,
                file_name=q_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with col2:
        with open(b_file, "rb") as f:
            st.download_button(
                "📋 Download BOM (Excel)",
                f,
                file_name=b_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
