import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import datetime

# ---------------- FUNCTIONS ---------------- #

def upright_sheet_size(width, depth):
    return width + depth + width + 15

def pipe_beam_sheet_size(height, width):
    return 2 * (height + width)

def roll_form_beam_sheet_size(height, width):
    return ((4*height) + (2*width)) + 20

def deep_bar_size(rack_depth):
    return rack_depth - 65

def cross_bar_length(deep_bar, gap, upright_depth):
    return ((deep_bar - 50)**2 + gap**2) ** 0.5 + 50

def weight(length, width, thickness):
    return (length * width * thickness * 7.85) / 1_000_000


# ---------------- UI ---------------- #

st.title("🏗 Pallet Rack Quotation Generator")

client = st.text_input("Client Name")
rack_types = st.number_input("Number of Rack Types", min_value=1, step=1)

rack_data = []

# ---------------- INPUT ---------------- #

for i in range(rack_types):

    st.header(f"Rack Type {i+1}")

    name = st.text_input("Rack Name", key=f"name{i}")
    main_qty = st.number_input("Main Qty", key=f"main{i}")
    addon_qty = st.number_input("Add-On Qty", key=f"addon{i}")
    levels = st.number_input("Levels", min_value=1, key=f"levels{i}")

    st.subheader("Upright")
    uw = st.number_input("Width", key=f"uw{i}")
    ud = st.number_input("Depth", key=f"ud{i}")
    ul = st.number_input("Length", key=f"ul{i}")
    ut = st.number_input("Thickness", key=f"ut{i}")

    st.subheader("Beam")
    bt = st.selectbox("Type", ["Pipe", "Roll"], key=f"bt{i}")
    bh = st.number_input("Height", key=f"bh{i}")
    bw = st.number_input("Width", key=f"bw{i}")
    bl = st.number_input("Length", key=f"bl{i}")
    bth = st.number_input("Thickness", key=f"bth{i}")

    st.subheader("Deep Bar")
    depth = st.number_input("Rack Depth", key=f"depth{i}")
    dth = st.number_input("Thickness", key=f"dth{i}")

    st.subheader("Cross Bar")
    gap = st.number_input("Gap (600–900)", min_value=600, max_value=900, step=10, key=f"gap{i}")
    method = st.number_input("Method", key=f"method{i}")
    cth = st.number_input("Thickness", key=f"cth{i}")

    rack_data.append({
        "name": name,
        "main_qty": main_qty,
        "addon_qty": addon_qty,
        "levels": levels,
        "uw": uw, "ud": ud, "ul": ul, "ut": ut,
        "bt": bt, "bh": bh, "bw": bw, "bl": bl, "bth": bth,
        "depth": depth, "dth": dth,
        "gap": gap, "method": method, "cth": cth
    })

# Accessories
st.header("Accessories")
cg_qty = st.number_input("Column Guard Qty")
rc_qty = st.number_input("Row Connector Qty")

# Rate Input
st.header("Pricing")
default_rate = st.number_input("Default Rate (₹ per item)", value=100)

# ---------------- GENERATE ---------------- #

if st.button("Generate Professional Quotation"):

    wb = Workbook()
    ws = wb.active
    ws.title = "Quotation"

    # Styles
    bold = Font(bold=True)
    big_bold = Font(size=14, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left")

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Header
    ws.merge_cells("A1:F1")
    ws["A1"] = "YOUR COMPANY NAME PVT. LTD."
    ws["A1"].font = big_bold
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = "DETAILED QUOTATION FOR HDR PALLET RACKING SYSTEM"
    ws["A2"].font = bold
    ws["A2"].alignment = center

    ws["A4"] = "Client Name:"
    ws["B4"] = client

    ws["D4"] = "Date:"
    ws["E4"] = datetime.date.today().strftime("%d-%m-%Y")

    ws["A5"] = "Subject:"
    ws["B5"] = "Supply of Pallet Racking System"

    # Table Header
    headers = ["Sr No", "Description", "Qty", "Unit Wt (Kg)", "Rate (₹)", "Amount (₹)"]
    row = 7

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = bold
        cell.alignment = center
        cell.border = border

    row += 1
    sr = 1
    subtotal = 0

    # Data
    for data in rack_data:

        beam_per_rack = 2 * data["levels"]

        # Upright
        uwid = upright_sheet_size(data["uw"], data["ud"])
        u_wt = weight(data["ul"], uwid, data["ut"])
        u_qty = (4 * data["main_qty"]) + (2 * data["addon_qty"])

        # Beam
        if data["bt"] == "Pipe":
            bwid = pipe_beam_sheet_size(data["bh"], data["bw"])
        else:
            bwid = roll_form_beam_sheet_size(data["bh"], data["bw"])

        b_wt = weight(data["bl"], bwid, data["bth"]) + 1.15
        b_qty = beam_per_rack * (data["main_qty"] + data["addon_qty"])

        # Deep Bar
        dlen = deep_bar_size(data["depth"])
        d_wt = weight(dlen, 92, data["dth"])
        d_qty = (4 * data["main_qty"]) + (2 * data["addon_qty"])

        # Cross Bar
        eff_h = data["ul"] - data["method"]
        num_cross = int(eff_h // data["gap"])

        clen = cross_bar_length(dlen, data["gap"], data["ud"])
        c_wt = weight(clen, 92, data["cth"])
        c_qty = num_cross * (2 * data["main_qty"] + data["addon_qty"])

        items = [
            ("Upright", u_qty, u_wt),
            ("Beam", b_qty, b_wt),
            ("Deep Bar", d_qty, d_wt),
            ("Cross Bar", c_qty, c_wt)
        ]

        for comp, qty, unit_wt in items:

            rate = default_rate
            amount = qty * rate

            ws.append([
                sr,
                f"{data['name']} - {comp}",
                qty,
                round(unit_wt, 2),
                rate,
                round(amount, 2)
            ])

            for col in range(1, 7):
                ws.cell(row=row, column=col).border = border

            subtotal += amount
            row += 1
            sr += 1

    # Accessories
    ws.append([])
    row += 1

    ws.append([sr, "Column Guard", cg_qty, "", 500, cg_qty * 500])
    subtotal += cg_qty * 500
    sr += 1

    ws.append([sr, "Row Connector", rc_qty, "", 200, rc_qty * 200])
    subtotal += rc_qty * 200

    # Totals
    row = ws.max_row + 2

    ws[f"E{row}"] = "Subtotal"
    ws[f"F{row}"] = round(subtotal, 2)

    gst = subtotal * 0.18
    ws[f"E{row+1}"] = "GST (18%)"
    ws[f"F{row+1}"] = round(gst, 2)

    grand_total = subtotal + gst
    ws[f"E{row+2}"] = "Grand Total"
    ws[f"F{row+2}"] = round(grand_total, 2)

    ws[f"E{row+2}"].font = bold
    ws[f"F{row+2}"].font = bold

    # Terms
    row += 4
    ws.merge_cells(f"A{row}:F{row}")
    ws[f"A{row}"] = "Terms & Conditions"
    ws[f"A{row}"].font = bold

    terms = [
        "1. GST extra as applicable.",
        "2. Delivery within 3-4 weeks.",
        "3. Payment: 50% advance, balance before dispatch.",
        "4. Transportation extra.",
        "5. Installation extra."
    ]

    for t in terms:
        row += 1
        ws.merge_cells(f"A{row}:F{row}")
        ws[f"A{row}"] = t

    # Save
    filename = f"{client}_Quotation.xlsx"
    wb.save(filename)

    st.success(f"✅ Quotation Generated: {filename}")

    with open(filename, "rb") as f:
        st.download_button("Download Quotation", f, file_name=filename)
