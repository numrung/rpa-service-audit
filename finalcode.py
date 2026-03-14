import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import plotly.express as px
import urllib.parse
import platform

# --- เช็คระบบปฏิบัติการ (สำหรับ Outlook Preview) ---
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

# --- ส่วนที่ 0: เครื่องมือเตรียมไฟล์ (Sort ข้อมูล) ---
st.subheader("🛠️ ส่วนที่ 0: เครื่องมือเตรียมไฟล์ (จัดระเบียบข้อมูลก่อน)")
with st.expander("🪄 คลิกเพื่อใช้งานเครื่องมือจัดกลุ่มทะเบียนรถ (Input_Service_Sorted)"):
    st.info("ใช้สำหรับจัดกลุ่มรถทะเบียนเดียวกันให้อยู่ติดกัน และเรียงวันที่จากเก่าไปใหม่")
    prep_file = st.file_uploader("อัปโหลดไฟล์ Input_Service_Data.xlsx", type=['xlsx'], key="prep_tool")
    
    if prep_file:
        if st.button("🚀 กดจัดกลุ่มข้อมูล"):
            try:
                df_prep = pd.read_excel(prep_file, skiprows=2)
                df_prep.columns = df_prep.columns.str.strip()
                
                p_date_col = 'วันที่เข้าศูนย์บริการ'
                if p_date_col not in df_prep.columns:
                    for c in df_prep.columns:
                        if 'วันที่' in str(c): p_date_col = c; break
                
                df_prep['tmp_date'] = df_prep[p_date_col].apply(parse_thai_date)
                df_prep = df_prep.sort_values(by=['ป้ายทะเบียนรถ', 'tmp_date'], ascending=[True, True])
                df_prep = df_prep.drop(columns=['tmp_date'])
                
                output_prep = io.BytesIO()
                with pd.ExcelWriter(output_prep, engine='xlsxwriter') as writer:
                    df_prep.to_excel(writer, index=False, startrow=2)
                
                st.success("✅ จัดระเบียบเสร็จแล้ว!")
                st.download_button("📥 ดาวน์โหลดไฟล์ Sorted", output_prep.getvalue(), "Input_Service_Sorted.xlsx")
            except Exception as e:
                st.error(f"❌ Error: {e}")

st.divider()

# --- ส่วนของระบบหลัก (Sidebar) ---
with st.sidebar:
    st.header("📂 อัปโหลดข้อมูลระบบหลัก")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (ใช้ไฟล์ Sorted)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขอะไหล่ (Logic)", type=['xlsx'])
    file_email = st.file_uploader("4. ข้อมูล Email (Email.xlsx)", type=['xlsx'])
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลระบบ")

if process_btn:
    if not (file_input and file_mileage and file_logic and file_email):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบ!")
    else:
        try:
            with st.spinner('กำลังคำนวณ...'):
                df_logic = pd.read_excel(file_logic)
                df_m = pd.read_excel(file_mileage, header=2)
                df_new = pd.read_excel(file_input, skiprows=2)
                df_email = pd.read_excel(file_email)
                df_new.columns = df_new.columns.str.strip()

                # --- Logic ประมวลผลหลัก ---
                df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
                mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
                
                df_new = df_new.dropna(subset=['ป้ายทะเบียนรถ'])
                df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict).fillna(0)

                target_date_col = 'วันที่เข้าศูนย์บริการ'
                for c in df_new.columns:
                    if 'วันที่' in str(c): target_date_col = c; break

                def calculate_service(row):
                    detail = str(row.get('รายละเอียดการเข้าศูนย์', '')).lower()
                    plus_km, plus_mo = 10000.0, 6.0
                    found_items = []
                    for _, lg in df_logic.iterrows():
                        kw = re.sub(r'\(.*\)', '', str(lg['รายการอะไหล่/การบริการ'])).lower().strip()
                        if kw in detail:
                            found_items.append(str(lg['รายการอะไหล่/การบริการ']))
                            if pd.notna(lg['ระยะเปลี่ยนถ่าย (กม.)']): plus_km = min(plus_km, float(lg['ระยะเปลี่ยนถ่าย (กม.)']))
                            if pd.notna(lg['ระยะเวลา (เดือน)']): plus_mo = min(plus_mo, float(lg['ระยะเวลา (เดือน)']))
                    if not found_items: found_items = ["ตรวจเช็คทั่วไป"]
                    km_val = str(row.get('เลขไมล์ที่เข้าศูนย์บริการ', '0')).replace(',', '')
                    curr_km = float(km_val) if km_val != 'nan' and km_val.strip() != '' else 0
                    date_in = parse_thai_date(row.get(target_date_col))
                    next_date = date_in + timedelta(days=int(plus_mo * 30.44)) if pd.notna(date_in) else None
                    return pd.Series([curr_km + plus_km, next_date, ", ".join(found_items), curr_km, date_in])

                df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ', 'ไมล์ที่เข้าล่าสุด', 'วันที่เข้าล่าสุด']] = df_new.apply(calculate_service, axis=1)

                today = datetime.now()
                def get_status(row):
                    if row['ไมล์ปัจจุบัน'] == 0: return "🔍 ไม่พบข้อมูลไมล์"
                    diff_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
                    diff_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
                    if diff_km <= 0 or diff_days <= 0: return f"🔴 ถึงกำหนด ({row['รายการ']})"
                    elif diff_km <= 1000 or diff_days <= 15: return f"🟡 ใกล้ถึง (เหลือ {int(diff_km):,} กม.)"
                    return "🟢 ปกติ"

                df_new['สถานะการแจ้งเตือน'] = df_new.apply(get_status, axis=1)

                # --- Merge ข้อมูล Email ---
                # พี่เช็คชื่อคอลัมน์ในไฟล์หลักที่ใช้ Match กับ Name ใน Email.xlsx นะครับ (ตัวอย่างใช้ 'ป้ายทะเบียนรถ' หรือ 'ชื่อผู้รับผิดชอบ')
                match_col = 'ป้ายทะเบียนรถ' # หรือเปลี่ยนเป็น 'ชื่อพนักงาน' ตามจริง
                if 'Name' in df_email.columns:
                    df_final = pd.merge(df_new, df_email, left_on=match_col, right_on='Name', how='left')
                else:
                    df_final = df_new.copy()
                    st.warning("⚠️ คอลัมน์ 'Name' ไม่พบในไฟล์ Email.xlsx")

            # --- ส่วนแสดงผล Dashboard & Table ---
            c1, c2, c3 = st.columns(3)
            red_n = len(df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴")])
            yellow_n = len(df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🟡")])
            c1.metric("รถทั้งหมด", len(df_final))
            c2.metric("ถึงกำหนด", red_n)
            c3.metric("ใกล้ถึง", yellow_n)

            # --- ระบบส่ง Email ---
            st.subheader("📧 เตรียมแจ้งเตือน Email")
            df_alert = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴|🟡")].copy()
            
            for idx, row in df_alert.iterrows():
                col_t, col_b = st.columns([4, 1])
                t_to = row['to'] if 'to' in row and pd.notna(row['to']) else ""
                t_cc = row['CC'] if 'CC' in row and pd.notna(row['CC']) else ""
                
                col_t.write(f"🚗 {row['ป้ายทะเบียนรถ']} | 👤 {row.get('Name', 'ไม่ระบุ')} | {row['สถานะการแจ้งเตือน']}")
                
                mail_sub = f"แจ้งเตือนซ่อมบำรุงรถยนต์: {row['ป้ายทะเบียนรถ']}"
                if IS_WINDOWS:
                    mail_body = f"เรียน คุณ {row.get('Name','')}<br>รถทะเบียน <b>{row['ป้ายทะเบียนรถ']}</b> ถึงกำหนดเช็คระยะ<br>ไมล์ปัจจุบัน: {row['ไมล์ปัจจุบัน']:,} กม."
                    if col_b.button(f"Preview", key=f"btn_{idx}"):
                        preview_outlook_windows(t_to, t_cc, mail_sub, mail_body)
                else:
                    mail_plain = f"เรียน คุณ {row.get('Name','')}\nรถทะเบียน {row['ป้ายทะเบียนรถ']} ถึงกำหนดเช็คระยะ\nไมล์ปัจจุบัน: {row['ไมล์ปัจจุบัน']:,} กม."
                    mailto_url = f"mailto:{t_to}?cc={t_cc}&subject={urllib.parse.quote(mail_sub)}&body={urllib.parse.quote(mail_plain)}"
                    col_b.markdown(f'<a href="{mailto_url}"><button style="background-color:#0078d4;color:white;border:none;border-radius:5px;padding:5px 10px;cursor:pointer;">Open</button></a>', unsafe_allow_html=True)

            # --- ปุ่มดาวน์โหลดรายงาน ---
            st.divider()
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลดรายงาน (.xlsx)", buffer.getvalue(), "Service_Audit_Final.xlsx")

        except Exception as e:
            st.error(f"❌ Error: {e}")
