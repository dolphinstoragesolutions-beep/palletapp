import streamlit as st
import datetime
from xml.dom import minidom
from xml.etree import ElementTree as ET

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

def create_excel_xml(sheet_name, headers, data_rows, title_info=None):
    """Create Excel XML format matching reference file"""
    
    # Create root Workbook element
    workbook = ET.Element('Workbook', {
        'xmlns': 'urn:schemas-microsoft-com:office:spreadsheet',
        'xmlns:o': 'urn:schemas-microsoft-com:office:office',
        'xmlns:x': 'urn:schemas-microsoft-com:office:excel',
        'xmlns:ss': 'urn:schemas-microsoft-com:office:spreadsheet',
        'xmlns:html': 'http://www.w3.org/TR/REC-html40'
    })
    
    # Document properties
    doc_props = ET.SubElement(workbook, 'DocumentProperties', {
        'xmlns': 'urn:schemas-microsoft-com:office:office'
    })
    ET.SubElement(doc_props, 'Author').text = 'Brij Industries'
    ET.SubElement(doc_props, 'LastAuthor').text = 'Brij Industries'
    ET.SubElement(doc_props, 'Created').text = datetime.datetime.now().isoformat()
    ET.SubElement(doc_props, 'Company').text = 'Brij Industries'
    ET.SubElement(doc_props, 'Version').text = '16.00'
    
    # Excel workbook settings
    excel_workbook = ET.SubElement(workbook, 'ExcelWorkbook', {
        'xmlns': 'urn:schemas-microsoft-com:office:excel'
    })
    ET.SubElement(excel_workbook, 'WindowHeight').text = '15915'
    ET.SubElement(excel_workbook, 'WindowWidth').text = '28800'
    ET.SubElement(excel_workbook, 'WindowTopX').text = '120'
    ET.SubElement(excel_workbook, 'WindowTopY').text = '75'
    ET.SubElement(excel_workbook, 'ProtectStructure').text = 'False'
    ET.SubElement(excel_workbook, 'ProtectWindows').text = 'False'
    
    # Styles definition (matching reference file)
    styles = ET.SubElement(workbook, 'Styles')
    
    # Default style
    default_style = ET.SubElement(styles, 'Style', {'ss:ID': 'Default', 'ss:Name': 'Normal'})
    ET.SubElement(default_style, 'Alignment', {'ss:Vertical': 'Bottom'})
    ET.SubElement(default_style, 'Borders')
    ET.SubElement(default_style, 'Font', {'ss:FontName': 'Calibri', 'ss:Size': '11'})
    ET.SubElement(default_style, 'Interior')
    ET.SubElement(default_style, 'NumberFormat')
    ET.SubElement(default_style, 'Protection')
    
    # Header style (bold, centered)
    header_style = ET.SubElement(styles, 'Style', {'ss:ID': 's22'})
    ET.SubElement(header_style, 'Font', {
        'ss:FontName': 'Calibri',
        'ss:Size': '11',
        'ss:Bold': '1'
    })
    ET.SubElement(header_style, 'Alignment', {
        'ss:Horizontal': 'Center',
        'ss:Vertical': 'Bottom'
    })
    ET.SubElement(header_style, 'Interior', {
        'ss:Color': '#D3D3D3',
        'ss:Pattern': 'Solid'
    })
    
    # Title style
    title_style = ET.SubElement(styles, 'Style', {'ss:ID': 's26'})
    ET.SubElement(title_style, 'Font', {
        'ss:FontName': 'Calibri',
        'ss:Size': '14',
        'ss:Bold': '1'
    })
    ET.SubElement(title_style, 'Alignment', {'ss:Horizontal': 'Left'})
    
    # Right-align number style
    right_style = ET.SubElement(styles, 'Style', {'ss:ID': 's29'})
    ET.SubElement(right_style, 'NumberFormat', {'ss:Format': 'Standard'})
    ET.SubElement(right_style, 'Alignment', {'ss:Horizontal': 'Right'})
    
    # Currency style
    currency_style = ET.SubElement(styles, 'Style', {'ss:ID': 's30'})
    ET.SubElement(currency_style, 'NumberFormat', {'ss:Format': '₹#,##0.00'})
    ET.SubElement(currency_style, 'Alignment', {'ss:Horizontal': 'Right'})
    
    # Worksheet
    worksheet = ET.SubElement(workbook, 'Worksheet', {'ss:Name': sheet_name})
    
    # Table
    table = ET.SubElement(worksheet, 'Table', {
        'ss:ExpandedColumnCount': str(len(headers) if headers else '6'),
        'x:FullColumns': '1',
        'x:FullRows': '1',
        'ss:DefaultColumnWidth': '65',
        'ss:DefaultRowHeight': '15'
    })
    
    # Add title rows if provided
    row_idx = 1
    if title_info:
        # Title
        title_row = ET.SubElement(table, 'Row', {'ss:Height': '20'})
        title_cell = ET.SubElement(title_row, 'Cell', {'ss:StyleID': 's26'})
        title_data = ET.SubElement(title_cell, 'Data', {'ss:Type': 'String'})
        title_data.text = title_info.get('title', 'BRIJ INDUSTRIES')
        ET.SubElement(title_row, 'Cell')  # Empty cells for remaining columns
        ET.SubElement(title_row, 'Cell')
        ET.SubElement(title_row, 'Cell')
        ET.SubElement(title_row, 'Cell')
        ET.SubElement(title_row, 'Cell')
        
        # Client
        client_row = ET.SubElement(table, 'Row')
        client_cell = ET.SubElement(client_row, 'Cell')
        client_data = ET.SubElement(client_cell, 'Data', {'ss:Type': 'String'})
        client_data.text = f"M/S: {title_info.get('client', '')}"
        for _ in range(5):
            ET.SubElement(client_row, 'Cell')
        
        # Offer No
        offer_row = ET.SubElement(table, 'Row')
        offer_cell = ET.SubElement(offer_row, 'Cell')
        offer_data = ET.SubElement(offer_cell, 'Data', {'ss:Type': 'String'})
        offer_data.text = f"OFFER NO: {title_info.get('offer_no', '')}"
        for _ in range(5):
            ET.SubElement(offer_row, 'Cell')
        
        # Date
        date_row = ET.SubElement(table, 'Row')
        date_cell = ET.SubElement(date_row, 'Cell')
        date_data = ET.SubElement(date_cell, 'Data', {'ss:Type': 'String'})
        date_data.text = f"DATE: {title_info.get('date', '')}"
        for _ in range(5):
            ET.SubElement(date_row, 'Cell')
        
        # Empty row
        ET.SubElement(table, 'Row')
    
    # Headers row
    if headers:
        header_row = ET.SubElement(table, 'Row')
        for header in headers:
            header_cell = ET.SubElement(header_row, 'Cell', {'ss:StyleID': 's22'})
            header_data = ET.SubElement(header_cell, 'Data', {'ss:Type': 'String'})
            header_data.text = str(header)
    
    # Data rows
    for row_data in data_rows:
        data_row = ET.SubElement(table, 'Row')
        for idx, cell_value in enumerate(row_data):
            # Determine cell style
            if isinstance(cell_value, (int, float)) and idx > 0:
                style = 's29'
                data_type = 'Number'
            elif isinstance(cell_value, (int, float)):
                style = 's29'
                data_type = 'Number'
            else:
                style = ''
                data_type = 'String'
            
            cell = ET.SubElement(data_row, 'Cell', {'ss:StyleID': style} if style else {})
            cell_data = ET.SubElement(cell, 'Data', {'ss:Type': data_type})
            
            if data_type == 'Number':
                cell_data.text = str(float(cell_value) if isinstance(cell_value, int) else cell_value)
            else:
                cell_data.text = str(cell_value)
    
    # Worksheet options
    ws_options = ET.SubElement(worksheet, 'WorksheetOptions', {'xmlns': 'urn:schemas-microsoft-com:office:excel'})
    ET.SubElement(ws_options, 'PageSetup')
    ET.SubElement(ws_options, 'FitToPage')
    ET.SubElement(ws_options, 'Print')
    ET.SubElement(ws_options, 'Zoom').text = '100'
    ET.SubElement(ws_options, 'Selected')
    ET.SubElement(ws_options, 'Panes')
    ET.SubElement(ws_options, 'ProtectObjects').text = 'False'
    ET.SubElement(ws_options, 'ProtectScenarios').text = 'False'
    
    # Convert to string with proper formatting
    xml_str = ET.tostring(workbook, encoding='utf-8')
    dom = minidom.parseString(xml_str)
    return dom.toprettyxml(indent='  ', encoding='utf-8')

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

    uw = st.number_input("Upright Width", key=f"uw{i}")
    ud = st.number_input("Upright Depth", key=f"ud{i}")
    ul = st.number_input("Upright Length", key=f"ul{i}")
    ut = st.number_input("Thickness", key=f"ut{i}")

    bt = st.selectbox("Beam Type", ["Pipe", "Roll"], key=f"bt{i}")
    bh = st.number_input("Beam Height", key=f"bh{i}")
    bw = st.number_input("Beam Width", key=f"bw{i}")
    bl = st.number_input("Beam Length", key=f"bl{i}")
    bth = st.number_input("Beam Thickness", key=f"bth{i}")

    depth = st.number_input("Rack Depth", key=f"d{i}")
    dth = st.number_input("Deep Thickness", key=f"dth{i}")

    gap = st.number_input("Gap", key=f"g{i}")
    method = st.number_input("Method", key=f"me{i}")
    cth = st.number_input("Cross Thickness", key=f"cth{i}")

    rack_data.append({
        "name": name, "main_qty": main_qty, "addon_qty": addon_qty,
        "levels": levels, "uw": uw, "ud": ud, "ul": ul, "ut": ut,
        "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
        "depth": depth, "dth": dth, "gap": gap, "method": method, "cth": cth
    })

rate = st.number_input("Rate per KG", value=100)

# ---------------- GENERATE ---------------- #

if st.button("GENERATE FILES"):
    
    # ================= CALCULATIONS ================= #
    
    quotation_rows = []
    bom_rows = []
    
    sr = 1
    total = 0
    
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
            
            bom_rows.append([
                r["name"], comp,
                m_qty, a_qty, total_qty,
                round(wt, 2),
                round(main_wt, 2),
                round(addon_wt, 2),
                round(total_wt, 2)
            ])
        
        amount = rack_total * rate
        total += amount
        
        quotation_rows.append([
            sr,
            f"{r['name']} Rack",
            "NOS",
            r["main_qty"] + r["addon_qty"],
            rate,
            round(amount, 2)
        ])
        
        sr += 1
    
    gst = total * 0.18
    
    # Add summary rows
    quotation_rows.append(["", "", "", "", "Basic", round(total, 2)])
    quotation_rows.append(["", "", "", "", "GST 18%", round(gst, 2)])
    quotation_rows.append(["", "", "", "", "Grand Total", round(total + gst, 2)])
    
    # ================= CREATE XML FILES ================= #
    
    # Quotation XML
    quotation_headers = ["S.No", "Description", "Unit", "Qty", "Price", "Amount"]
    title_info = {
        'title': 'BRIJ INDUSTRIES',
        'client': client,
        'offer_no': offer_no,
        'date': str(datetime.date.today())
    }
    
    quotation_xml = create_excel_xml("Quotation", quotation_headers, quotation_rows, title_info)
    
    # BOM XML
    bom_headers = ["Rack", "Component", "Single Main Qty", "Single Addon Qty", 
                   "Total Qty", "Unit Weight", "Main Weight", "Addon Weight", "Total Weight"]
    
    bom_xml = create_excel_xml("BOM", bom_headers, bom_rows, None)
    
    # SAVE FILES
    q_file = f"{client}_Quotation.xml"
    b_file = f"{client}_BOM.xml"
    
    with open(q_file, "wb") as f:
        f.write(quotation_xml)
    
    with open(b_file, "wb") as f:
        f.write(bom_xml)
    
    st.success("✅ Files Generated in Excel XML Format")
    
    with open(q_file, "rb") as f:
        st.download_button("📊 Download Quotation (Excel XML)", f, file_name=q_file)
    
    with open(b_file, "rb") as f:
        st.download_button("📋 Download BOM (Excel XML)", f, file_name=b_file)
