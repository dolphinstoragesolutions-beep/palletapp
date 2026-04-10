import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import datetime
import io
import base64
from PIL import Image as PILImage
import tempfile
import os

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
    
    beam_wid = pipe_beam_sheet_size(rack["bh"], rack["bw"]) if rack["bt"] == "Pipe Beam" else roll_beam_sheet_size(rack["bh"], rack["bw"])
    b_wt = weight(rack["bl"], beam_wid, rack["bth"])
    
    d_len = deep_bar_size(rack["depth"])
    d_wt = weight(d_len, 92, rack["dth"])
    
    c_len = cross_length(d_len, rack["gap"], rack["ud"])
    c_wt = weight(c_len, 92, rack["cth"])
    
    # Total weight per rack set
    total_weight = (u_wt * 4) + (b_wt * 2 * rack["levels"]) + (d_wt * 4) + (c_wt * 2)
    return total_weight

def fireworks_animation():
    """Create fireworks animation using HTML/CSS/JS"""
    fireworks_html = """
    <div id="fireworks-container" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none; z-index: 9999;">
        <canvas id="fireworks-canvas" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%;"></canvas>
    </div>
    <script>
    (function() {
        const canvas = document.getElementById('fireworks-canvas');
        if (!canvas) return;
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
        const ctx = canvas.getContext('2d');
        
        let particles = [];
        let rockets = [];
        
        function random(min, max) {
            return Math.random() * (max - min) + min;
        }
        
        class Rocket {
            constructor(x, y, targetX, targetY) {
                this.x = x;
                this.y = y;
                this.targetX = targetX;
                this.targetY = targetY;
                this.velocity = 5;
                this.exploded = false;
            }
            
            update() {
                const dx = this.targetX - this.x;
                const dy = this.targetY - this.y;
                const distance = Math.sqrt(dx * dx + dy * dy);
                
                if (distance < 5) {
                    this.exploded = true;
                    this.explode();
                    return false;
                }
                
                const angle = Math.atan2(dy, dx);
                this.x += Math.cos(angle) * this.velocity;
                this.y += Math.sin(angle) * this.velocity;
                return true;
            }
            
            explode() {
                const particleCount = 80;
                for (let i = 0; i < particleCount; i++) {
                    const angle = random(0, Math.PI * 2);
                    const speed = random(2, 8);
                    const vx = Math.cos(angle) * speed;
                    const vy = Math.sin(angle) * speed;
                    const color = `hsl(${random(0, 360)}, 100%, 60%)`;
                    particles.push(new Particle(this.x, this.y, vx, vy, color));
                }
            }
            
            draw() {
                ctx.beginPath();
                ctx.arc(this.x, this.y, 3, 0, Math.PI * 2);
                ctx.fillStyle = '#FFA500';
                ctx.fill();
            }
        }
        
        class Particle {
            constructor(x, y, vx, vy, color) {
                this.x = x;
                this.y = y;
                this.vx = vx;
                this.vy = vy;
                this.color = color;
                this.life = 1;
                this.decay = 0.02;
            }
            
            update() {
                this.x += this.vx;
                this.y += this.vy;
                this.vy += 0.2;
                this.life -= this.decay;
                return this.life > 0;
            }
            
            draw() {
                ctx.beginPath();
                ctx.arc(this.x, this.y, 2, 0, Math.PI * 2);
                ctx.fillStyle = this.color;
                ctx.globalAlpha = this.life;
                ctx.fill();
                ctx.globalAlpha = 1;
            }
        }
        
        function createRocket() {
            const x = random(100, canvas.width - 100);
            const y = canvas.height;
            const targetX = random(100, canvas.width - 100);
            const targetY = random(100, canvas.height * 0.6);
            rockets.push(new Rocket(x, y, targetX, targetY));
        }
        
        function animate() {
            ctx.fillStyle = 'rgba(0, 0, 0, 0.1)';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            
            for (let i = rockets.length - 1; i >= 0; i--) {
                const alive = rockets[i].update();
                rockets[i].draw();
                if (!alive || rockets[i].exploded) {
                    rockets.splice(i, 1);
                }
            }
            
            for (let i = particles.length - 1; i >= 0; i--) {
                const alive = particles[i].update();
                particles[i].draw();
                if (!alive) {
                    particles.splice(i, 1);
                }
            }
            
            if (rockets.length < 5 && particles.length < 50) {
                createRocket();
            }
            
            requestAnimationFrame(animate);
        }
        
        // Start animation
        for (let i = 0; i < 3; i++) {
            setTimeout(() => createRocket(), i * 500);
        }
        animate();
        
        // Remove after 8 seconds
        setTimeout(() => {
            const container = document.getElementById('fireworks-container');
            if (container) container.remove();
        }, 8000);
    })();
    </script>
    """
    return st.components.v1.html(fireworks_html, height=0)

def apply_header_style(cell, bg_color='1B3A5C'):
    """Apply professional header styling with dark blue background"""
    cell.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return cell

def apply_orange_accent(cell):
    """Apply light orange accent styling"""
    cell.font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color='FF8C42', end_color='FF8C42', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    return cell

def apply_light_blue_bg(cell):
    """Apply light blue background for alternating rows"""
    cell.fill = PatternFill(start_color='E8F0FE', end_color='E8F0FE', fill_type='solid')
    return cell

def apply_border(cell, border_style='thin'):
    """Apply border to cell"""
    border = Border(
        left=Side(style=border_style, color='DDDDDD'),
        right=Side(style=border_style, color='DDDDDD'),
        top=Side(style=border_style, color='DDDDDD'),
        bottom=Side(style=border_style, color='DDDDDD')
    )
    cell.border = border
    return cell

def add_logo_to_excel(ws, logo_bytes):
    """Add logo to Excel worksheet"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
            tmp_file.write(logo_bytes)
            tmp_file_path = tmp_file.name
        
        img = XLImage(tmp_file_path)
        img.width = 150
        img.height = 80
        img.anchor = 'A1'
        ws.add_image(img)
        
        os.unlink(tmp_file_path)
        return True
    except Exception as e:
        return False

# ---------------- UI ---------------- #

st.set_page_config(page_title="Professional Quotation Generator", layout="wide", page_icon="🎆")

# Custom CSS for better UI
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #1B3A5C 0%, #FF8C42 100%);
    }
    .main-header {
        background: linear-gradient(135deg, #1B3A5C, #FF8C42);
        padding: 20px;
        border-radius: 15px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        margin-bottom: 20px;
    }
    .stButton > button {
        background: linear-gradient(135deg, #FF8C42, #FF6B00);
        color: white;
        font-weight: bold;
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 20px rgba(255,140,66,0.4);
    }
    .css-1d391kg {
        background-color: rgba(255,255,255,0.1);
        border-radius: 10px;
        padding: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🎆 PROFESSIONAL QUOTATION GENERATOR</h1><p>Modular Mezzanine Floor & Racking Systems</p></div>', unsafe_allow_html=True)

# Logo upload section
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("### 🖼️ Upload Your Company Logo")
    uploaded_logo = st.file_uploader("", type=['png', 'jpg', 'jpeg', 'gif'], label_visibility="collapsed")
    
    if uploaded_logo is not None:
        logo_bytes = uploaded_logo.getvalue()
        st.image(logo_bytes, width=150)
        st.success("✅ Logo uploaded successfully!")
    else:
        st.info("📁 Please upload your company logo (DSS Dolphin Storage Solutions)")
        # Default placeholder logo text
        logo_bytes = None

st.divider()

col1, col2 = st.columns(2)

with col1:
    client = st.text_input("🏢 Customer Name (M/S)", value="STYLE BAZAAR")
    product = st.text_input("📦 Product", value="MODULAR MEZZANINE FLOOR")
    offer_no = st.text_input("📄 Offer No", value="DSS-IV/25-26/0712")
    
with col2:
    date = st.date_input("📅 Date", datetime.date.today())
    project_name = st.text_input("🏗️ Project Name", value="MODULE MEZZANINE FLOOR")

st.divider()

# Module/Rack Types
st.subheader("📦 RACK CONFIGURATIONS")
st.info("💡 Configure your rack types - each will appear as a separate module")

rack_types = st.number_input("Number of Rack Types", min_value=1, value=1)

rack_data = []

for i in range(int(rack_types)):
    with st.expander(f"🔹 RACK TYPE {chr(65+i)} - Complete Configuration", expanded=i==0):
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
        
        # Hidden parameters for calculation
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
st.subheader("💰 PRICE CALCULATION PREVIEW")

if rack_data:
    sample_rack = rack_data[0]
    sample_weight = calculate_rack_weight(sample_rack)
    main_rack_price = sample_weight * rate_per_kg
    addon_rack_price = (sample_weight * 0.7) * rate_per_kg
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📊 Sample Rack Weight", f"{sample_weight:.2f} kg")
    with col2:
        st.metric("🏗️ Main Rack Price", f"₹{main_rack_price:,.2f}")
    with col3:
        st.metric("📦 Add-on Rack Price", f"₹{addon_rack_price:,.2f}", delta="~70% of main")

# ---------------- GENERATE ---------------- #

if st.button("🎆 GENERATE PROFESSIONAL QUOTATION", type="primary", use_container_width=True):
    
    # Trigger fireworks animation
    fireworks_animation()
    
    # ================= CREATE QUOTATION WORKBOOK ================= #
    
    wb = Workbook()
    ws = wb.active
    ws.title = "COMMERCIAL OFFER"
    
    # Set column widths
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    
    # ================= ADD LOGO TO EXCEL ================= #
    
    if uploaded_logo is not None:
        add_logo_to_excel(ws, uploaded_logo.getvalue())
        logo_row_offset = 3
    else:
        logo_row_offset = 0
    
    # ================= HEADER SECTION ================= #
    
    # Company header
    if not uploaded_logo:
        ws.merge_cells('A1:C1')
        ws.merge_cells('D1:F1')
        
        logo_cell = ws['A1']
        logo_cell.value = "🐬 DSS"
        logo_cell.font = Font(name='Segoe UI', size=22, bold=True, color='FF8C42')
        logo_cell.alignment = Alignment(horizontal='left', vertical='center')
        
        company_cell = ws['D1']
        company_cell.value = "DOLPHIN STORAGE SOLUTIONS"
        company_cell.font = Font(name='Segoe UI', size=16, bold=True, color='1B3A5C')
        company_cell.alignment = Alignment(horizontal='right', vertical='center')
    else:
        ws.row_dimensions[1].height = 80
        ws.merge_cells('D1:F1')
        company_cell = ws['D1']
        company_cell.value = "DOLPHIN STORAGE SOLUTIONS"
        company_cell.font = Font(name='Segoe UI', size=16, bold=True, color='1B3A5C')
        company_cell.alignment = Alignment(horizontal='right', vertical='center')
    
    # Subtitle
    start_row = 2 if not uploaded_logo else 3
    ws.merge_cells(f'A{start_row}:F{start_row}')
    subtitle_cell = ws[f'A{start_row}']
    subtitle_cell.value = "BRIJ INDUSTRIES"
    subtitle_cell.font = Font(name='Segoe UI', size=12, bold=True, color='FF8C42')
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Contact info
    contact_row = start_row + 1
    ws.merge_cells(f'A{contact_row}:F{contact_row}')
    contact_cell = ws[f'A{contact_row}']
    contact_cell.value = "📍 86/3/1 ROAD NO 7, MUNDKA INDUSTRIAL AREA, SOUTH DELHI - 110041"
    contact_cell.font = Font(name='Segoe UI', size=9)
    contact_cell.alignment = Alignment(horizontal='center', vertical='center')
    contact_cell.fill = PatternFill(start_color='FFF5EE', end_color='FFF5EE', fill_type='solid')
    
    email_row = contact_row + 1
    ws.merge_cells(f'A{email_row}:F{email_row}')
    email_cell = ws[f'A{email_row}']
    email_cell.value = "✉️ brijindustries09@rediffmail.com  |  🌐 WWW.BRIJINDUSTRIES.IN  |  📞 +91 9625589161, +91 9811096149"
    email_cell.font = Font(name='Segoe UI', size=9)
    email_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    gst_row = email_row + 1
    ws.merge_cells(f'A{gst_row}:F{gst_row}')
    gst_cell = ws[f'A{gst_row}']
    gst_cell.value = "✅ GST NO: 07AAMFB6403G1ZM"
    gst_cell.font = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
    gst_cell.alignment = Alignment(horizontal='center', vertical='center')
    gst_cell.fill = PatternFill(start_color='FF8C42', end_color='FF8C42', fill_type='solid')
    
    # Decorative line
    ws.row_dimensions[gst_row + 1].height = 8
    
    # ================= CUSTOMER DETAILS CARD ================= #
    
    details_start = gst_row + 2
    
    # Create bordered box
    for row in range(details_start, details_start + 5):
        for col in range(1, 7):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(
                left=Side(style='thin', color='1B3A5C'),
                right=Side(style='thin', color='1B3A5C'),
                top=Side(style='thin', color='1B3A5C') if row == details_start else None,
                bottom=Side(style='thin', color='1B3A5C') if row == details_start + 4 else None
            )
            cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    # Customer details
    details = [
        ("🏢 M/S:", client, "📅 DATE:", date.strftime("%d-%m-%Y")),
        ("📦 PRODUCT:", product, "📄 OFFER NO:", offer_no),
        ("🏗️ PROJECT:", project_name, "", "")
    ]
    
    row = details_start
    for detail in details:
        # Left label
        cell1 = ws.cell(row=row, column=1, value=detail[0])
        cell1.font = Font(bold=True, size=10, color='1B3A5C')
        cell1.alignment = Alignment(horizontal='left', vertical='center')
        
        # Left value
        cell2 = ws.cell(row=row, column=2, value=detail[1])
        cell2.font = Font(size=10, color='333333')
        cell2.alignment = Alignment(horizontal='left', vertical='center')
        
        # Right label
        if detail[2]:
            cell3 = ws.cell(row=row, column=4, value=detail[2])
            cell3.font = Font(bold=True, size=10, color='1B3A5C')
            cell3.alignment = Alignment(horizontal='left', vertical='center')
        
        # Right value
        if detail[3]:
            cell4 = ws.cell(row=row, column=5, value=detail[3])
            cell4.font = Font(size=10, color='333333')
            cell4.alignment = Alignment(horizontal='left', vertical='center')
        
        row += 1
    
    # Empty row
    ws.row_dimensions[details_start + 5].height = 10
    
    # ================= PRODUCT TABLE ================= #
    
    table_start = details_start + 6
    
    # Section header with orange accent
    ws.merge_cells(f'A{table_start}:F{table_start}')
    section_header = ws[f'A{table_start}']
    section_header.value = "📊 COMMERCIAL OFFER DETAILS"
    section_header.font = Font(name='Segoe UI', size=14, bold=True, color='FFFFFF')
    section_header.fill = PatternFill(start_color='FF8C42', end_color='FF8C42', fill_type='solid')
    section_header.alignment = Alignment(horizontal='center', vertical='center')
    
    # Table headers
    headers = ["#", "DESCRIPTION", "TYPE", "QTY", "UNIT PRICE (₹)", "TOTAL (₹)"]
    header_row = table_start + 1
    
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1B3A5C', end_color='1B3A5C', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        apply_border(cell)
    
    # ================= CALCULATIONS ================= #
    
    current_row = header_row + 1
    sr = 1
    total_basic_amount = 0
    
    for rack in rack_data:
        # Calculate rack weight and price
        rack_weight = calculate_rack_weight(rack)
        main_rack_price = rack_weight * rate_per_kg
        addon_rack_price = (rack_weight * 0.7) * rate_per_kg
        
        # Main Rack row
        main_total = main_rack_price * rack["main_qty"]
        total_basic_amount += main_total
        
        # Alternating row colors
        bg_color = 'FFF9F0' if sr % 2 == 1 else 'FFFFFF'
        
        ws.cell(row=current_row, column=1, value=sr)
        ws.cell(row=current_row, column=2, value=f"{rack['module']} - {rack['name']}")
        ws.cell(row=current_row, column=3, value="🏗️ MAIN RACK")
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
            
            bg_color = 'FFF9F0' if sr % 2 == 1 else 'FFFFFF'
            
            ws.cell(row=current_row, column=1, value=sr)
            ws.cell(row=current_row, column=2, value=f"{rack['module']} - {rack['name']}")
            ws.cell(row=current_row, column=3, value="📦 ADD-ON RACK")
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
    subtotal_label = ws.cell(row=subtotal_row, column=1, value="💰 SUB TOTAL")
    subtotal_label.font = Font(bold=True, size=11, color='1B3A5C')
    subtotal_label.alignment = Alignment(horizontal='right', vertical='center')
    
    subtotal_value = ws.cell(row=subtotal_row, column=6, value=round(total_basic_amount, 2))
    subtotal_value.font = Font(bold=True, size=11, color='1B3A5C')
    subtotal_value.alignment = Alignment(horizontal='right', vertical='center')
    subtotal_value.number_format = '#,##0.00'
    
    for col in range(1, 7):
        cell = ws.cell(row=subtotal_row, column=col)
        cell.fill = PatternFill(start_color='FFF0E6', end_color='FFF0E6', fill_type='solid')
        apply_border(cell)
    
    # ================= SUMMARY SECTION ================= #
    
    current_row += 2
    summary_start = current_row
    
    # Basic Amount
    ws.cell(row=current_row, column=4, value="BASIC AMOUNT")
    ws.cell(row=current_row, column=5, value=round(total_basic_amount, 2))
    ws.cell(row=current_row, column=4).font = Font(bold=True, size=10)
    ws.cell(row=current_row, column=5).font = Font(bold=True, size=10)
    ws.cell(row=current_row, column=5).number_format = '#,##0.00'
    
    current_row += 1
    
    # Freight
    ws.cell(row=current_row, column=4, value="🚚 FREIGHT CHARGES")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(size=10)
    ws.cell(row=current_row, column=5).font = Font(size=10, italic=True, color='FF8C42')
    
    current_row += 1
    
    # Erection
    ws.cell(row=current_row, column=4, value="🔧 ERECTION CHARGES")
    ws.cell(row=current_row, column=5, value="INCLUSIVE")
    ws.cell(row=current_row, column=4).font = Font(size=10)
    ws.cell(row=current_row, column=5).font = Font(size=10, italic=True, color='FF8C42')
    
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
    grand_cell = ws.cell(row=current_row, column=4, value=f"🎯 GRAND TOTAL: ₹{grand_total:,.2f}")
    grand_cell.font = Font(name='Segoe UI', size=14, bold=True, color='FFFFFF')
    grand_cell.fill = PatternFill(start_color='FF8C42', end_color='FF8C42', fill_type='solid')
    grand_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Apply borders to summary section
    for row in range(summary_start, current_row + 1):
        for col in [4, 5]:
            apply_border(ws.cell(row=row, column=col))
    
    # ================= FOOTER NOTES ================= #
    
    current_row += 2
    
    footer_bg = PatternFill(start_color='FFF5EE', end_color='FFF5EE', fill_type='solid')
    
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
        "✨ For queries, contact our sales team:",
        "📞 +91 9625589161  |  ✉️ sales@brijindustries.in"
    ]
    
    for note in notes:
        if note.startswith("📋") or note.startswith("🏦") or note.startswith("✨"):
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
            cell.font = Font(size=9, color='555555')
            cell.fill = footer_bg
            cell.alignment = Alignment(horizontal='left', vertical='center')
        current_row += 1
    
    # Save file
    filename = f"{client.replace(' ', '_')}_COMMERCIAL_OFFER.xlsx"
    wb.save(filename)
    
    st.success("✅ PROFESSIONAL QUOTATION GENERATED SUCCESSFULLY!")
    
    # Download button
    with open(filename, "rb") as f:
        st.download_button(
            label="🎆 DOWNLOAD COMMERCIAL OFFER (EXCEL)",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Display preview
    st.subheader("📋 QUOTATION SUMMARY")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🏢 Customer", client)
    with col2:
        st.metric("📄 Offer No", offer_no)
    with col3:
        st.metric("💰 Basic Amount", f"₹{total_basic_amount:,.2f}")
    with col4:
        st.metric("🎯 GRAND TOTAL", f"₹{grand_total:,.2f}", delta="Including GST")
    
    st.divider()
    st.info("💡 The Excel file features: Light Orange (#FF8C42) & Dark Blue (#1B3A5C) color scheme | Professional borders | Logo integration | Fireworks animation on generation")
