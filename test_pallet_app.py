import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
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


# ---------------- UI ---------------- #

st.title("🏗 PROFESSIONAL QUOTATION + BOM GENERATOR")

client = st.text_input("Client Name")
offer_no = st.text_input("Offer No")

rack_types = st.number_input("No. of Rack Types", min_value=1)

rack_data = []

for i in range(rack_types):
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

    rack_data.append(locals())

rate = st.number_input("Rate per KG", value=100)

# ---------------- GENERATE ---------------- #

if st.button("GENERATE FILES"):

    # ================= QUOTATION ================= #

    wb1 = Workbook()
    ws = wb1.active
    ws.title = "Quotation"

    bold = Font(bold=True)
    center = Alignment(horizontal="center")

    ws["A1"] = "BRIJ INDUSTRIES"
    ws["A2"] = f"M/S: {client}"
    ws["A3"] = f"OFFER NO: {offer_no}"
    ws["A4"] = f"DATE: {datetime.date.today()}"

    ws.append([])
    ws.append(["S.No","Description","Unit","Qty","Price","Amount"])

    sr = 1
    total = 0

    # ================= BOM ================= #

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "BOM"

    ws2.append([
        "Rack","Component",
        "Single Main Qty","Single Addon Qty","Total Qty",
        "Unit Weight","Main Weight","Addon Weight","Total Weight"
    ])

    # ================= CALCULATION ================= #

    for r in rack_data:

        uwid = upright_sheet_size(r["uw"], r["ud"])
        u_wt = weight(r["ul"], uwid, r["ut"])

        beam_wid = pipe_beam_sheet_size(r["bh"], r["bw"]) if r["bt"]=="Pipe" else roll_beam_sheet_size(r["bh"], r["bw"])
        b_wt = weight(r["bl"], beam_wid, r["bth"])

        d_len = deep_bar_size(r["depth"])
        d_wt = weight(d_len, 92, r["dth"])

        c_len = cross_length(d_len, r["gap"], r["ud"])
        c_wt = weight(c_len, 92, r["cth"])

        components = [
            ("Upright", u_wt, 4, 2),
            ("Beam", b_wt, 2*r["levels"], 2*r["levels"]),
            ("Deep Bar", d_wt, 4, 2),
            ("Cross Bar", c_wt, 2, 1)
        ]

        rack_total = 0

        for comp, wt, m_qty, a_qty in components:

            main_wt = wt * m_qty
            addon_wt = wt * a_qty

            total_qty = (m_qty*r["main_qty"]) + (a_qty*r["addon_qty"])
            total_wt = wt * total_qty

            rack_total += total_wt

            # BOM
            ws2.append([
                r["name"], comp,
                m_qty, a_qty, total_qty,
                round(wt,2),
                round(main_wt,2),
                round(addon_wt,2),
                round(total_wt,2)
            ])

        # Quotation line
        amount = rack_total * rate
        total += amount

        ws.append([
            sr,
            f"{r['name']} Rack",
            "NOS",
            r["main_qty"] + r["addon_qty"],
            rate,
            round(amount,2)
        ])

        sr += 1

    gst = total * 0.18

    ws.append([])
    ws.append(["","","","","Basic",total])
    ws.append(["","","","","GST 18%",gst])
    ws.append(["","","","","Grand Total",total+gst])

    # SAVE FILES
    q_file = f"{client}_Quotation.xlsx"
    b_file = f"{client}_BOM.xlsx"

    wb1.save(q_file)
    wb2.save(b_file)

    st.success("✅ Files Generated")

    with open(q_file,"rb") as f:
        st.download_button("Download Quotation",f,file_name=q_file)

    with open(b_file,"rb") as f:
        st.download_button("Download BOM",f,file_name=b_file)
