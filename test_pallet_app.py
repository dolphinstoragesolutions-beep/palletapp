import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font
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
    a = deep_bar - 50
    b = gap
    return ((a**2 + b**2) ** 0.5) + 50

def weight(length, width, thickness):
    return (length * width * thickness * 7.85) / 1_000_000


# ---------------- UI ---------------- #

st.title("🏗 Pallet Rack BOQ & Tonnage Calculator")

client = st.text_input("Client Name")
rack_types = st.number_input("Number of Rack Types", min_value=1, step=1)

rack_data = []

# ---------------- MULTIPLE RACK TYPES ---------------- #

for i in range(rack_types):

    st.header(f"Rack Type {i+1}")

    name = st.text_input(f"Rack Name", key=f"name{i}")
    main_qty = st.number_input(f"Main Qty", key=f"main{i}")
    addon_qty = st.number_input(f"Add-On Qty", key=f"addon{i}")
    levels = st.number_input(f"Levels", min_value=1, key=f"levels{i}")

    # Upright
    st.subheader("Upright")
    uw = st.number_input("Width", key=f"uw{i}")
    ud = st.number_input("Depth", key=f"ud{i}")
    ul = st.number_input("Length", key=f"ul{i}")
    ut = st.number_input("Thickness", key=f"ut{i}")

    # Beam
    st.subheader("Beam")
    bt = st.selectbox("Type", ["Pipe", "Roll"], key=f"bt{i}")
    bh = st.number_input("Height", key=f"bh{i}")
    bw = st.number_input("Width", key=f"bw{i}")
    bl = st.number_input("Length", key=f"bl{i}")
    bth = st.number_input("Thickness", key=f"bth{i}")

    # Deep Bar
    st.subheader("Deep Bar")
    depth = st.number_input("Rack Depth", key=f"depth{i}")
    dth = st.number_input("Thickness", key=f"dth{i}")

    # Cross Bar
    st.subheader("Cross Bar")
    gap = st.number_input(
    f"Enter Gap (600–900) - Rack {i+1}",
    min_value=600,
    max_value=900,
    step=10,
    key=f"gap_{i}"
    )

    method = st.number_input(
    f"Enter Method Value - Rack {i+1}",
    min_value=0,
    step=50,
    key=f"method_{i}"
    )
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

# ---------------- ACCESSORIES ---------------- #

st.header("Accessories")

cg_qty = st.number_input("Column Guard Qty")
rc_qty = st.number_input("Row Connector Qty")

# ---------------- CALCULATE ---------------- #

if st.button("Generate BOQ & Excel"):

    wb = Workbook()
    ws = wb.active
    ws.title = "Rack BOQ"

    headers = ["Rack Type","Component","Size","Length","Thickness","Main Qty","Add-On Qty","Total Qty","Unit Wt","Main Wt","Add-On Wt"]
    ws.append(headers)

    total_uprights = 0
    grand_main = 0
    grand_addon = 0

    for data in rack_data:

        beam_per_rack = 2 * data["levels"]

        # Upright
        uwid = upright_sheet_size(data["uw"], data["ud"])
        u_wt = weight(data["ul"], uwid, data["ut"])

        u_main_qty = 4 * data["main_qty"]
        u_addon_qty = 2 * data["addon_qty"]
        u_total = u_main_qty + u_addon_qty

        total_uprights += u_total

        u_main = u_wt * u_main_qty
        u_addon = u_wt * u_addon_qty

        # Beam
        if data["bt"] == "Pipe":
            bwid = pipe_beam_sheet_size(data["bh"], data["bw"])
        else:
            bwid = roll_form_beam_sheet_size(data["bh"], data["bw"])

        b_wt = weight(data["bl"], bwid, data["bth"]) + 1.15
        b_main_qty = beam_per_rack * data["main_qty"]
        b_addon_qty = beam_per_rack * data["addon_qty"]

        b_main = b_wt * b_main_qty
        b_addon = b_wt * b_addon_qty

        # Deep Bar
        dlen = deep_bar_size(data["depth"])
        d_wt = weight(dlen, 92, data["dth"])

        d_main_qty = 4 * data["main_qty"]
        d_addon_qty = 2 * data["addon_qty"]

        d_main = d_wt * d_main_qty
        d_addon = d_wt * d_addon_qty

        # Cross Bar
        eff_h = data["ul"] - data["method"]
        num_cross = int(eff_h // data["gap"])

        clen = cross_bar_length(dlen, data["gap"], data["ud"])
        c_wt = weight(clen, 92, data["cth"])

        c_main_qty = num_cross * 2 * data["main_qty"]
        c_addon_qty = num_cross * data["addon_qty"]

        c_main = c_wt * c_main_qty
        c_addon = c_wt * c_addon_qty

        total_main = u_main + b_main + d_main + c_main
        total_addon = u_addon + b_addon + d_addon + c_addon

        grand_main += total_main
        grand_addon += total_addon

        ws.append([data["name"],"TOTAL","","","","","", "", "", round(total_main,2), round(total_addon,2)])

    # ---------------- ACCESSORIES SHEET ---------------- #

    ws2 = wb.create_sheet("Accessories")
    ws2.append(["Item","Qty","Unit Wt","Total"])

    ws2.append(["Column Guard", cg_qty, 3.75, cg_qty*3.75])
    ws2.append(["Row Connector", rc_qty, 1, rc_qty*1])

    ws2.append(["Base Plate", total_uprights, "Included", "-"])

    # ---------------- SAVE ---------------- #

    filename = f"{client}_BOQ.xlsx"
    wb.save(filename)

    st.success(f"✅ File Generated: {filename}")

    with open(filename, "rb") as f:
        st.download_button("Download Excel", f, file_name=filename)
