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

st.title("🚗 ระบบบริหารจัดการกำหนดซ่อมบำรุงรถยนต์ (Smart Alert)")
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
                status_placeholder.write("⏳ 2/4: กำลังตรวจสอบทะเบียนและประวัติล่าสุด...")
                df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
                mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
                
                df_new = df_new.dropna(subset=['ป้ายทะเบียนรถ'])
                df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict)
                progress_bar.progress(50)

                # Step 3
                status_placeholder.write("⏳ 3/4: กำลังคำนวณวันนัดหมายถัดไป...")
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
                    
                    raw_km = str(row['เลขไมล์ที่เข้าศูนย์บริการ']).replace(',', '')
                    curr_km = float(raw_km) if raw_km != 'nan' else 0
                    
                    date_in = pd.to_datetime(row['วันที่เข้าศูนย์บริการ'], dayfirst=True, errors='coerce')
                    next_date = date_in + timedelta(days=int(plus_mo * 30.44)) if pd.notna(date_in) else None
                    
                    # คืนค่าเพิ่ม: ไมล์นัดหมาย, วันนัดหมาย, รายการ, ไมล์ที่เข้าล่าสุด, วันที่เข้าล่าสุด
                    return pd.Series([curr_km + plus_km, next_date, ", ".join(found_items), curr_km, date_in])

                df_new[['ไมล์นัดหมาย', 'วันที่นัดหมาย', 'รายการ', 'ไมล์ที่เข้าล่าสุด', 'วันที่เข้าล่าสุด']] = df_new.apply(calculate_service, axis=1)
                progress_bar.progress(75)

                # Step 4
                status_placeholder.write("⏳ 4/4: สรุปสถานะ 3 ระดับ (แดง/เหลือง/เขียว)...")
                today = datetime.now()

                def get_status(row):
                    if pd.isna(row['ไมล์ปัจจุบัน']) or row['ไมล์ปัจจุบัน'] == 0: 
                        return "🔍 ไม่พบข้อมูลไมล์"
                    
                    warning_km = 1000  # ระยะเตือนล่วงหน้า
                    warning_days = 15   # วันเตือนล่วงหน้า
                    
                    diff_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
                    diff_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
                    
                    # 1. สีแดง: ถึงกำหนดแล้ว
                    if diff_km <= 0 or diff_days <= 0:
                        return f"🔴 ถึงกำหนด ({row['รายการ']})"
                    
                    # 2. สีเหลือง: ใกล้ถึงกำหนด
                    elif diff_km <= warning_km or diff_days <= warning_days:
                        return f"🟡 ใกล้ถึง (เหลือ {int(diff_km):,} กม.)"
                    
                    # 3. สีเขียว: ปกติ
                    else:
                        return "🟢 ปกติ"

                df_new['สถานะการแจ้งเตือน'] = df_new.apply(get_status, axis=1)
                
                # แปลงตัวเลขให้สวยงาม
                for col in ['ไมล์ปัจจุบัน', 'ไมล์นัดหมาย', 'ไมล์ที่เข้าล่าสุด']:
                    df_new[col] = pd.to_numeric(df_new[col], errors='coerce').fillna(0).astype(int)

                progress_bar.progress(100)
                status_placeholder.success("✅ ประมวลผลเสร็จสมบูรณ์!")

            # --- Dashboard Visualization ---
            st.divider()
            c1, c2, c3, c4 = st.columns(4)
            num_all = len(df_new)
            num_red = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🔴")])
            num_yellow = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🟡")])
            num_green = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🟢")])

            c1.metric("รถทั้งหมด", f"{num_all} คัน")
            c2.metric("ต้องซ่อมด่วน", f"{num_red} คัน", delta=num_red, delta_color="inverse")
            c3.metric("ใกล้ถึงกำหนด", f"{num_yellow} คัน")
            c4.metric("สถานะปกติ", f"{num_green} คัน")

            col_chart, col_table = st.columns([1, 2])
            with col_chart:
                st.write("📊 สัดส่วนสถานะสุขภาพ Fleet")
                chart_data = pd.DataFrame({
                    'สถานะ': ['ถึงกำหนด', 'ใกล้ถึง', 'ปกติ'], 
                    'จำนวน': [num_red, num_yellow, num_green]
                })
                fig = px.pie(chart_data, values='จำนวน', names='สถานะ', 
                             color='สถานะ', color_discrete_map={'ถึงกำหนด':'#EF5350', 'ใกล้ถึง':'#FFB74D', 'ปกติ':'#66BB6A'})
                st.plotly_chart(fig, use_container_width=True)

            with col_table:
                st.write("📋 ตารางตรวจสอบประวัติและสถานะล่าสุด")
                
                # จัดลำดับคอลัมน์ให้อ่านง่าย
                output_columns = [
                    'ป้ายทะเบียนรถ', 
                    'วันที่เข้าล่าสุด', 
                    'ไมล์ที่เข้าล่าสุด', 
                    'ไมล์ปัจจุบัน', 
                    'สถานะการแจ้งเตือน'
                ]
                
                def color_row(val):
                    if '🔴' in val: return 'background-color: #ffebee'
                    if '🟡' in val: return 'background-color: #fff3e0'
                    if '🟢' in val: return 'background-color: #e8f5e9'
                    return ''
                
                # แสดงผลตารางพร้อม Format
                df_show = df_new[output_columns].copy()
                df_show['วันที่เข้าล่าสุด'] = df_show['วันที่เข้าล่าสุด'].dt.strftime('%d/%m/%Y')
                
                st.dataframe(
                    df_show.style.applymap(color_row, subset=['สถานะการแจ้งเตือน'])
                    .format({"ไมล์ที่เข้าล่าสุด": "{:,d}", "ไมล์ปัจจุบัน": "{:,d}"}),
                    use_container_width=True
                )

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_new.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลดรายงานฉบับเต็ม (.xlsx)", data=buffer.getvalue(), file_name="Service_Audit_Final.xlsx", mime="application/vnd.ms-excel")

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            st.exception(e)
