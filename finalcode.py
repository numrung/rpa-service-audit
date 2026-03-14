import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import plotly.express as px
import urllib.parse
import platform

# --- เช็คระบบปฏิบัติการเพื่อเลือกใช้ Library ---
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    try:
        import win32com.client as win32
    except ImportError:
        IS_WINDOWS = False

# --- ตั้งค่า UI ---
st.set_page_config(page_title="CMS Service Audit & Auto Mail", page_icon="🚗", layout="wide")

st.title("🚗 ระบบบริหารจัดการกำหนดซ่อมบำรุงรถยนต์ (Smart Alert)")
st.write(f"📅 วันที่ปัจจุบัน (ค.ศ.): {datetime.now().strftime('%d/%m/%Y')}")

# --- Helper Functions ---
def clean_plate(text):
    if pd.isna(text): return ""
    text = str(text).replace(" ", "").replace("-", "").strip()
    match = re.search(r'(\d{6,7}|\d?[ก-ฮ]{2,3}\d{1,4})', text)
    return match.group(1) if match else text[:7]

def parse_thai_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        if isinstance(date_val, datetime):
            if date_val.year > 2500: return date_val.replace(year=date_val.year - 543)
            return date_val
        date_str = str(date_val).strip()
        for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%d-%m-%y', '%d/%m/%y'):
            try:
                dt = datetime.strptime(date_str, fmt)
                if dt.year > 2500: dt = dt.replace(year=dt.year - 543)
                return dt
            except: continue
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except: return pd.NaT

def preview_outlook_windows(to, cc, subject, body):
    """ฟังก์ชันเปิด Outlook สำหรับ Windows Local"""
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = str(cc) if pd.notna(cc) else ""
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Display()
        return True
    except Exception as e:
        st.error(f"Error opening Outlook: {e}")
        return False

# --- ส่วนของ Sidebar ---
with st.sidebar:
    st.header("📂 อัปโหลดข้อมูล")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (Sorted)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขอะไหล่ (Logic)", type=['xlsx'])
    file_email = st.file_uploader("4. ข้อมูล Email (Email.xlsx)", type=['xlsx'])
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลและเตรียมเมล")

# --- ระบบหลัก ---
if process_btn:
    if not (file_input and file_mileage and file_logic and file_email):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 4 ไฟล์!")
    else:
        try:
            # 1. อ่านข้อมูล
            df_logic = pd.read_excel(file_logic)
            df_m = pd.read_excel(file_mileage, header=2)
            df_new = pd.read_excel(file_input, skiprows=2)
            df_email = pd.read_excel(file_email) # ไฟล์ที่พี่ส่งมา (Name, to, CC)
            
            df_new.columns = df_new.columns.str.strip()
            
            # 2. คลีนข้อมูลและคำนวณ (Logic เดิมของพี่)
            df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
            
            df_new = df_new.dropna(subset=['ป้ายทะเบียนรถ'])
            df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
            df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict).fillna(0)

            # คอลัมน์วันที่
            t_col = 'วันที่เข้าศูนย์บริการ'
            for c in df_new.columns:
                if 'วันที่' in str(c): t_col = c; break

            def calc_logic(row):
                detail = str(row.get('รายละเอียดการเข้าศูนย์', '')).lower()
                p_km, p_mo = 10000.0, 6.0
                found = []
                for _, lg in df_logic.iterrows():
                    kw = re.sub(r'\(.*\)', '', str(lg['รายการอะไหล่/การบริการ'])).lower().strip()
                    if kw in detail:
                        found.append(str(lg['รายการอะไหล่/การบริการ']))
                        if pd.notna(lg['ระยะเปลี่ยนถ่าย (กม.)']): p_km = min(p_km, float(lg['ระยะเปลี่ยนถ่าย (กม.)']))
                        if pd.notna(lg['ระยะเวลา (เดือน)']): p_mo = min(p_mo, float(lg['ระยะเวลา (เดือน)']))
                if not found: found = ["ตรวจเช็คทั่วไป"]
                km_val = str(row.get('เลขไมล์ที่เข้าศูนย์บริการ', '0')).replace(',', '')
                curr_km = float(km_val) if km_val != 'nan' and km_val.strip() != '' else 0
                date_in = parse_thai_date(row.get(t_col))
                next_d = date_in + timedelta(days=int(p_mo * 30.44)) if pd.notna(date_in) else None
                return pd.Series([curr_km + p_km, next_d, ", ".join(found), curr_km, date_in])

            df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ', 'ไมล์ที่เข้าล่าสุด', 'วันที่เข้าล่าสุด']] = df_new.apply(calc_logic, axis=1)

            # 3. ตรวจสอบสถานะ
            today = datetime.now()
            def get_status(row):
                if row['ไมล์ปัจจุบัน'] == 0: return "🔍 ไม่พบข้อมูลไมล์"
                diff_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
                diff_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
                if diff_km <= 0 or diff_days <= 0: return f"🔴 ถึงกำหนด ({row['รายการ']})"
                elif diff_km <= 1000 or diff_days <= 15: return f"🟡 ใกล้ถึง (เหลือ {int(diff_km):,} กม.)"
                return "🟢 ปกติ"

            df_new['สถานะการแจ้งเตือน'] = df_new.apply(get_status, axis=1)

            # 4. Merge กับ Email Config
            # สมมติคอลัมน์ในไฟล์หลักที่เก็บชื่อคนขับชื่อ 'ชื่อคนขับ' (พี่เปลี่ยนให้ตรงกับไฟล์พี่นะครับ)
            name_col_in_main = 'ชื่อคนขับ' if 'ชื่อคนขับ' in df_new.columns else df_new.columns[0]
            df_final = pd.merge(df_new, df_email, left_on=name_col_in_main, right_on='Name', how='left')

            # --- ส่วนการแสดงผล Metrics ---
            c1, c2, c3 = st.columns(3)
            red_zone = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴")]
            yellow_zone = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🟡")]
            c1.metric("รถทั้งหมด", len(df_final))
            c2.metric("ต้องเข้าศูนย์ด่วน", len(red_zone), delta_color="inverse")
            c3.metric("ใกล้ถึงกำหนด", len(yellow_zone))

            # --- ส่วนระบบส่ง Email ---
            st.subheader("📧 ระบบแจ้งเตือนผ่าน Email (Outlook)")
            df_alert = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴|🟡")].copy()

            if df_alert.empty:
                st.success("✅ เยี่ยมมาก! ไม่มีรถที่ถึงกำหนดส่งเมลในขณะนี้")
            else:
                for idx, row in df_alert.iterrows():
                    with st.container():
                        col_text, col_btn = st.columns([4, 1])
                        t_to = row['to'] if pd.notna(row['to']) else ""
                        t_cc = row['CC'] if pd.notna(row['CC']) else ""
                        
                        col_text.write(f"🚗 **{row['ป้ายทะเบียนรถ']}** | 👤 {row['Name']} | {row['สถานะการแจ้งเตือน']}")
                        
                        subject = f"แจ้งเตือนซ่อมบำรุงรถยนต์: {row['ป้ายทะเบียนรถ']}"
                        
                        if IS_WINDOWS:
                            # แบบ Windows: ใช้ win32com (สวยงาม)
                            body_html = f"<h3>เรียน คุณ {row['Name']}</h3><p>รถทะเบียน <b>{row['ป้ายทะเบียนรถ']}</b> {row['สถานะการแจ้งเตือน']}</p><p>ไมล์ปัจจุบัน: {row['ไมล์ปัจจุบัน']:,} กม.</p>"
                            if col_btn.button(f"Preview (Outlook)", key=f"win_{idx}"):
                                preview_outlook_windows(t_to, t_cc, subject, body_html)
                        else:
                            # แบบ Cloud: ใช้ mailto (ปลอดภัย)
                            body_plain = f"เรียน คุณ {row['Name']}\n\nรถทะเบียน {row['ป้ายทะเบียนรถ']} {row['สถานะการแจ้งเตือน']}\nไมล์ปัจจุบัน: {row['ไมล์ปัจจุบัน']:,} กม.\n\nกรุณานำรถเข้าศูนย์บริการตามกำหนด"
                            mailto_url = f"mailto:{t_to}?cc={t_cc}&subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body_plain)}"
                            col_btn.markdown(f'<a href="{mailto_url}"><button style="background-color:#0078d4;color:white;border:none;border-radius:5px;padding:5px 10px;cursor:pointer;">Open Outlook</button></a>', unsafe_allow_html=True)

            # --- ปุ่มดาวน์โหลดรายงาน ---
            st.divider()
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลดรายงานสรุป (.xlsx)", buffer.getvalue(), "Service_Audit_Final.xlsx")

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# --- ส่วนท้าย (Footer) ---
if not IS_WINDOWS:
    st.info("💡 หมายเหตุ: ปัจจุบันรันบนระบบ Cloud/Non-Windows ระบบจะเปิด Outlook ผ่าน Browser แทนการใช้ Automation")
