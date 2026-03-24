import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import urllib.parse

st.set_page_config(page_title="CMS Service Audit", page_icon="🚗", layout="wide")

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

# --- Sidebar ---
with st.sidebar:
    st.header("📂 อัปโหลดไฟล์")
    f_in = st.file_uploader("1. ไฟล์ Sorted", type=['xlsx'])
    f_mi = st.file_uploader("2. ไฟล์เลขไมล์ปัจจุบัน", type=['xlsx'])
    f_lg = st.file_uploader("3. ไฟล์เงื่อนไข (Logic)", type=['xlsx'])
    f_em = st.file_uploader("4. ไฟล์ Email.xlsx", type=['xlsx'])
    process = st.button("🚀 ประมวลผล")

if process and f_in and f_mi and f_lg and f_em:
    try:
        df_logic = pd.read_excel(f_lg)
        df_m = pd.read_excel(f_mi, header=2)
        df_new = pd.read_excel(f_in, skiprows=2)
        df_email = pd.read_excel(f_em, sheet_name='เงื่อนไข')
        
        df_new.columns = df_new.columns.str.strip()
        df_email.columns = df_email.columns.str.strip()

        # 1. จัดการเลขไมล์ปัจจุบัน
        df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
        mile_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
        df_new['ไมล์ปัจจุบัน'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate).map(mile_dict).fillna(0)

        # 2. ฟังก์ชันคำนวณ (ปรับให้ดึงรายการอะไหล่แม่นยำขึ้น)
        def calc_serv(row):
            # ดึงรายละเอียดมาลบช่องว่างและทำเป็นตัวพิมพ์เล็ก
            detail = str(row.get('รายละเอียดการเข้าศูนย์', '')).replace(' ', '').lower()
            p_km, p_mo = 10000.0, 6.0
            found_items = []
            
            for _, lg in df_logic.iterrows():
                # ดึง Keyword จาก Logic มาลบช่องว่างและวงเล็บออก
                raw_kw = str(lg['รายการอะไหล่/การบริการ'])
                clean_kw = re.sub(r'\(.*\)', '', raw_kw).replace(' ', '').lower().strip()
                
                # เช็คว่า Keyword "แฝงอยู่ใน" รายละเอียดหรือไม่
                if clean_kw != "" and clean_kw in detail:
                    found_items.append(raw_kw)
                    if pd.notna(lg['ระยะเปลี่ยนถ่าย (กม.)']): p_km = min(p_km, float(lg['ระยะเปลี่ยนถ่าย (กม.)']))
                    if pd.notna(lg['ระยะเวลา (เดือน)']): p_mo = min(p_mo, float(lg['ระยะเวลา (เดือน)']))
            
            res_items = ", ".join(found_items) if found_items else "ตรวจเช็คทั่วไป"
            km_in = float(str(row.get('เลขไมล์ที่เข้าศูนย์บริการ', '0')).replace(',', ''))
            dt_in = parse_thai_date(row.get('วันที่เข้าศูนย์บริการ'))
            return pd.Series([km_in + p_km, dt_in + timedelta(days=int(p_mo * 30.44)) if pd.notna(dt_in) else None, res_items])

        df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ']] = df_new.apply(calc_serv, axis=1)

        # 3. เช็คสถานะ
        today = datetime.now()
        def check_status(row):
            d_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
            d_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
            if d_km <= 0 or d_days <= 0: return f"🔴 ถึงกำหนด ({row['รายการ']})"
            elif d_km <= 1000 or d_days <= 15: return f"🟡 ใกล้ถึง ({row['รายการ']})" # เพิ่มรายการตรงนี้
            return "🟢 ปกติ"

        df_new['สถานะการแจ้งเตือน'] = df_new.apply(check_status, axis=1)

        # 4. Merge ข้อมูล Email
        df_email['MatchName'] = df_email['Name'].astype(str).str.replace(' ', '').str.strip()
        name_col = next((c for c in df_new.columns if any(x in str(c) for x in ['ชื่อ', 'พนักงาน'])), df_new.columns[0])
        df_new['MatchName'] = df_new[name_col].astype(str).str.replace(' ', '').str.strip()
        df_final = pd.merge(df_new, df_email, on='MatchName', how='left')

        # 5. แสดงรายการแจ้งเตือน
        st.subheader("📧 รายการแจ้งเตือน (ส่งผ่าน Mailto)")
        df_alert = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴|🟡")].copy()
        
        for idx, row in df_alert.iterrows():
            c1, c2 = st.columns([4, 1])
            d_name = row['Name'] if pd.notna(row['Name']) else row[name_col]
            c1.write(f"🚗 **{row['ป้ายทะเบียนรถ']}** | 👤 {d_name} | {row['สถานะการแจ้งเตือน']}")
            
            # สร้าง URL Mailto (ไม่ทำสีตามข้อจำกัด Browser)
            m_sub = f"แจ้งเตือนซ่อมบำรุง: {row['ป้ายทะเบียนรถ']}"
            m_body = f"เรียน คุณ {d_name}\n\nรถทะเบียน {row['ป้ายทะเบียนรถ']} มีสถานะ {row['สถานะการแจ้งเตือน']}\nไมล์ปัจจุบัน: {int(row['ไมล์ปัจจุบัน']):,} กม.\nกำหนดนัดหมาย: {int(row['ไมล์นัดหมาย']):,} กม.\nรายการ: {row['รายการ']}"
            m_url = f"mailto:{row['to']}?cc={row.get('CC', '')}&subject={urllib.parse.quote(m_sub)}&body={urllib.parse.quote(m_body)}"
            
            c2.markdown(f'<a href="{m_url}"><button style="background-color:#0078d4;color:white;border:none;border-radius:5px;padding:5px 10px;cursor:pointer;">Open</button></a>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
