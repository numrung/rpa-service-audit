import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
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

# --- ส่วนที่ 0: เครื่องมือเตรียมไฟล์ ---
st.subheader("🛠️ ส่วนที่ 0: เครื่องมือเตรียมไฟล์")
with st.expander("🪄 คลิกเพื่อใช้งานเครื่องมือจัดกลุ่มทะเบียนรถ"):
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
            except Exception as e: st.error(f"❌ Error: {e}")

st.divider()

# --- ส่วนของระบบหลัก (Sidebar) ---
with st.sidebar:
    st.header("📂 อัปโหลดข้อมูลระบบหลัก")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (Sorted)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขอะไหล่", type=['xlsx'])
    file_email = st.file_uploader("4. ข้อมูล Email.xlsx", type=['xlsx'])
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลระบบ")

if process_btn:
    if not (file_input and file_mileage and file_logic and file_email):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบ!")
    else:
        try:
            with st.spinner('กำลังประมวลผล...'):
                df_logic = pd.read_excel(file_logic)
                df_m = pd.read_excel(file_mileage, header=2)
                df_new = pd.read_excel(file_input, skiprows=2)
                df_email = pd.read_excel(file_email, sheet_name='เงื่อนไข')
                
                df_new.columns = df_new.columns.str.strip()
                df_email.columns = df_email.columns.str.strip()

                name_col_main = ""
                for c in df_new.columns:
                    if any(x in str(c) for x in ['ชื่อ', 'พนักงาน', 'ผู้รับผิดชอบ']):
                        name_col_main = c; break
                
                df_email['Name_Clean'] = df_email['Name'].astype(str).str.replace(' ', '').str.strip()
                df_new['Name_Match'] = df_new[name_col_main].astype(str).str.replace(' ', '').str.strip() if name_col_main else ""

                df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
                mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
                
                df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict).fillna(0)

                t_date_col = 'วันที่เข้าศูนย์บริการ'
                for c in df_new.columns:
                    if 'วันที่' in str(c): t_date_col = c; break

                def calc_serv(row):
                    det = str(row.get('รายละเอียดการเข้าศูนย์', '')).lower()
                    p_km, p_mo = 10000.0, 6.0
                    items = []
                    for _, lg in df_logic.iterrows():
                        kw = re.sub(r'\(.*\)', '', str(lg['รายการอะไหล่/การบริการ'])).lower().strip()
                        if kw in det:
                            items.append(str(lg['รายการอะไหล่/การบริการ']))
                            if pd.notna(lg['ระยะเปลี่ยนถ่าย (กม.)']): p_km = min(p_km, float(lg['ระยะเปลี่ยนถ่าย (กม.)']))
                            if pd.notna(lg['ระยะเวลา (เดือน)']): p_mo = min(p_mo, float(lg['ระยะเวลา (เดือน)']))
                    if not items: items = ["ตรวจเช็คทั่วไป"]
                    km_in = float(str(row.get('เลขไมล์ที่เข้าศูนย์บริการ', '0')).replace(',', '')) if str(row.get('เลขไมล์ที่เข้าศูนย์บริการ', '0')) != 'nan' else 0
                    dt_in = parse_thai_date(row.get(t_date_col))
                    nxt_dt = dt_in + timedelta(days=int(p_mo * 30.44)) if pd.notna(dt_in) else None
                    return pd.Series([km_in + p_km, nxt_dt, ", ".join(items)])

                df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ']] = df_new.apply(calc_serv, axis=1)

                today = datetime.now()
                def check_status(row):
                    if row['ไมล์ปัจจุบัน'] == 0: return "🔍 ไม่พบข้อมูลไมล์"
                    d_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
                    d_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
                    if d_km <= 0 or d_days <= 0: return f"🔴 ถึงกำหนด ({row['รายการ']})"
                    elif d_km <= 1000 or d_days <= 15: return f"🟡 ใกล้ถึง (เหลือ {int(d_km):,} กม.) ({row['รายการ']})"
                    return "🟢 ปกติ"

                df_new['สถานะการแจ้งเตือน'] = df_new.apply(check_status, axis=1)
                df_final = pd.merge(df_new, df_email, left_on='Name_Match', right_on='Name_Clean', how='left')

            st.subheader("📧 รายการแจ้งเตือน Email")
            df_alert = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴|🟡")].copy()
            
            if not df_alert.empty:
                # --- [ใหม่] ปุ่มสำหรับส่งทั้งหมด ---
                if IS_WINDOWS:
                    if st.button("📧 ส่ง Email ทั้งหมด (Send All Alerts)", type="primary"):
                        count_sent = 0
                        for _, row in df_alert.iterrows():
                            d_name = row['Name'] if pd.notna(row['Name']) else row.get(name_col_main, "ไม่ระบุชื่อ")
                            m_sub = f"แจ้งเตือนซ่อมบำรุง: {row['ป้ายทะเบียนรถ']}"
                            m_body = (
                                f"<div style='font-family: Tahoma; font-size: 14px;'>"
                                f"เรียน คุณ {d_name}<br><br>"
                                f"รถทะเบียน: <b>{row['ป้ายทะเบียนรถ']}</b><br>"
                                f"สถานะ: {row['สถานะการแจ้งเตือน']}<br>"
                                f"--------------------------------------------------<br>"
                                f"<b>ไมล์ปัจจุบัน:</b> {int(row['ไมล์ปัจจุบัน']):,} กม.<br>"
                                f"<b>กำหนดนัดหมายที่:</b> {int(row['ไมล์นัดหมาย']):,} กม.<br>"
                                f"<b>รายการบริการ:</b> {row['รายการ']}<br>"
                                f"--------------------------------------------------<br></div>"
                            )
                            if preview_outlook_windows(row.get('to', ''), row.get('CC', ''), m_sub, m_body):
                                count_sent += 1
                        st.success(f"🚀 เปิดหน้าต่าง Outlook เรียบร้อยแล้ว {count_sent} รายการ")
                
                st.divider()
                for idx, row in df_alert.iterrows():
                    with st.container():
                        col_t, col_b = st.columns([4, 1])
                        display_name = row['Name'] if pd.notna(row['Name']) else row.get(name_col_main, "ไม่ระบุชื่อ")
                        m_sub = f"แจ้งเตือนซ่อมบำรุง: {row['ป้ายทะเบียนรถ']}"
                        
                        col_t.write(f"🚗 **{row['ป้ายทะเบียนรถ']}** | 👤 {display_name} | {row['สถานะการแจ้งเตือน']}")
                        
                        m_html = (f"<div style='font-family: Tahoma; font-size: 14px;'>เรียน คุณ {display_name}<br><br>"
                                  f"รถทะเบียน: <b>{row['ป้ายทะเบียนรถ']}</b><br>สถานะ: {row['สถานะการแจ้งเตือน']}<br>"
                                  f"--------------------------------------------------<br>"
                                  f"<b>ไมล์ปัจจุบัน:</b> {int(row['ไมล์ปัจจุบัน']):,} กม.<br>"
                                  f"<b>กำหนดนัดหมายที่:</b> {int(row['ไมล์นัดหมาย']):,} กม.<br>"
                                  f"<b>รายการบริการ:</b> {row['รายการ']}<br>"
                                  f"--------------------------------------------------<br></div>")
                        
                        m_plain = (f"เรียน คุณ {display_name}\n\nรถทะเบียน: {row['ป้ายทะเบียนรถ']}\nสถานะ: {row['สถานะการแจ้งเตือน']}\n"
                                   f"--------------------------------------------------\n"
                                   f"ไมล์ปัจจุบัน: {int(row['ไมล์ปัจจุบัน']):,} กม.\nกำหนดนัดหมายที่: {int(row['ไมล์นัดหมาย']):,} กม.\n"
                                   f"รายการบริการ: {row['รายการ']}\n--------------------------------------------------")

                        if IS_WINDOWS:
                            if col_b.button(f"Preview", key=f"btn_{idx}"):
                                preview_outlook_windows(row.get('to', ''), row.get('CC', ''), m_sub, m_html)
                        else:
                            m_url = f"mailto:{row.get('to','') or ''}?cc={row.get('CC','') or ''}&subject={urllib.parse.quote(m_sub)}&body={urllib.parse.quote(m_plain)}"
                            col_b.markdown(f'<a href="{m_url}"><button style="background-color:#0078d4;color:white;border:none;border-radius:5px;padding:5px 10px;cursor:pointer;width:100%;">Open</button></a>', unsafe_allow_html=True)
            else:
                st.success("✅ ไม่มีรายการที่ต้องแจ้งเตือน")

            st.divider()
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลด Service_Audit_Final.xlsx", output.getvalue(), "Service_Audit_Final.xlsx")

        except Exception as e:
            st.error(f"❌ Error: {e}")
