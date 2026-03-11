import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import time
import plotly.express as px

# --- ตั้งค่า UI หน้าเว็บ ---
st.set_page_config(page_title="CMS Service Audit Dashboard", page_icon="🚗", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

st.title("🚗 ระบบบริหารจัดการกำหนดซ่อมบำรุงรถยนต์")
st.write(f"📅 วันที่ปัจจุบัน: {datetime.now().strftime('%d/%m/%Y')}")

def clean_plate(text):
    if pd.isna(text): return ""
    text = str(text).replace(" ", "").replace("-", "").strip()
    match = re.search(r'(\d{6,7}|\d?[ก-ฮ]{2,3}\d{1,4})', text)
    return match.group(1) if match else text[:7]

# --- Sidebar ---
with st.sidebar:
    st.header("📂 เมนูจัดการข้อมูล")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (Input)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน (Mileage)", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขการเปลี่ยนอะไหล่ (Logic)", type=['xlsx'])
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลระบบ")

if process_btn:
    if not (file_input and file_mileage and file_logic):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ไฟล์!")
    else:
        try:
            status_placeholder = st.empty()
            progress_bar = st.progress(0)
            
            with st.spinner('กำลังประมวลผล...'):
                # Step 1
                status_placeholder.write("⏳ 1/4: กำลังอ่านไฟล์...")
                df_logic = pd.read_excel(file_logic)
                df_m = pd.read_excel(file_mileage, header=2)
                df_new = pd.read_excel(file_input, skiprows=2)
                progress_bar.progress(25)

                # Step 2
                status_placeholder.write("⏳ 2/4: กำลังตรวจสอบทะเบียน...")
                df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
                mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
                
                df_new = df_new.dropna(subset=['ป้ายทะเบียนรถ'])
                df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict)
                progress_bar.progress(50)

                # Step 3
                status_placeholder.write("⏳ 3/4: กำลังคำนวณวันนัดหมาย...")
                def calculate_service(row):
                    detail = str(row['รายละเอียดการเข้าศูนย์']).lower()
                    plus_km, plus_mo = 10000.0, 6.0
                    found_items = []
                    for _, lg in df_logic.iterrows():
                        item_name = str(lg['รายการอะไหล่/การบริการ'])
                        kw = re.sub(r'\(.*\)', '', item_name).lower().strip()
                        if kw in detail:
                            found_items.append(item_name)
                            if pd.notna(lg['ระยะเปลี่ยนถ่าย (กม.)']):
                                plus_km = min(plus_km, float(lg['ระยะเปลี่ยนถ่าย (กม.)']))
                            if pd.notna(lg['ระยะเวลา (เดือน)']):
                                plus_mo = min(plus_mo, float(lg['ระยะเวลา (เดือน)']))
                    
                    if not found_items: found_items = ["ตรวจเช็คทั่วไป"]
                    curr_km = float(str(row['เลขไมล์ที่เข้าศูนย์บริการ']).replace(',', ''))
                    date_in = pd.to_datetime(row['วันที่เข้าศูนย์บริการ'], dayfirst=True, errors='coerce')
                    next_date = date_in + timedelta(days=int(plus_mo * 30.44)) if pd.notna(date_in) else None
                    return pd.Series([curr_km + plus_km, next_date, ", ".join(found_items)])

                df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ']] = df_new.apply(calculate_service, axis=1)
                progress_bar.progress(75)

                # Step 4
                status_placeholder.write("⏳ 4/4: สรุปสถานะการแจ้งเตือน...")
                today = datetime.now()
                def get_status(row):
                    if pd.isna(row['ไมล์ปัจจุบัน']): return "🔍 ไม่พบข้อมูลไมล์"
                    is_overdue = False
                    if row['ไมล์นัดหมาย'] < 900000 and row['ไมล์ปัจจุบัน'] >= row['ไมล์นัดหมาย']: is_overdue = True
                    if pd.notna(row['วันที่นัดหมาย']) and today >= row['วันที่นัดหมาย']: is_overdue = True
                    return f"🔴 ถึงกำหนด ({row['รายการ']})" if is_overdue else "🟢 ปกติ"

                df_new['สถานะการแจ้งเตือน'] = df_new.apply(get_status, axis=1)
                progress_bar.progress(100)
                status_placeholder.success("✅ ประมวลผลเสร็จสมบูรณ์!")

            # --- Dashboard Visualization ---
            st.divider()
            c1, c2, c3 = st.columns(3)
            num_all = len(df_new)
            num_overdue = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🔴")])
            num_ok = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🟢")])

            c1.metric("รถทั้งหมด", f"{num_all} คัน")
            c2.metric("ต้องซ่อมด่วน", f"{num_overdue} คัน", delta=num_overdue, delta_color="inverse")
            c3.metric("สถานะปกติ", f"{num_ok} คัน")

            col_chart, col_table = st.columns([1, 2])
            with col_chart:
                st.write("📊 สัดส่วนสถานะ")
                # แก้ไขจุดที่เคย Error: มั่นใจว่าวงเล็บปิดครบถ้วน
                chart_data = pd.DataFrame({
                    'สถานะ': ['ถึงกำหนด', 'ปกติ'], 
                    'จำนวน': [num_overdue, num_ok]
                })
                fig = px.pie(chart_data, values='จำนวน', names='สถานะ', 
                             color='สถานะ', color_discrete_map={'ถึงกำหนด':'#EF5350', 'ปกติ':'#66BB6A'})
                st.plotly_chart(fig, use_container_width=True)

            with col_table:
                st.write("📋 ตารางตรวจสอบสถานะ")
                def color_row(val):
                    color = '#ffebee' if '🔴' in val else '#e8f5e9' if '🟢' in val else 'white'
                    return f'background-color: {color}'
                st.dataframe(df_new[['ป้ายทะเบียนรถ', 'ไมล์ปัจจุบัน', 'สถานะการแจ้งเตือน']].style.applymap(color_row, subset=['สถานะการแจ้งเตือน']), use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_new.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลดรายงาน Excel", data=buffer.getvalue(), file_name="Service_Audit_Report.xlsx", mime="application/vnd.ms-excel")

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")