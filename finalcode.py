import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import io
import plotly.express as px

# --- ตั้งค่า UI ---
st.set_page_config(page_title="CMS Service Audit Dashboard", page_icon="🚗", layout="wide")

st.title("🚗 ระบบบริหารจัดการกำหนดซ่อมบำรุงรถยนต์ (Smart Alert)")
st.write(f"📅 วันที่ปัจจุบัน (ค.ศ.): {datetime.now().strftime('%d/%m/%Y')}")

# --- ส่วนของฟังก์ชันตัวช่วย (Helper Functions) ---

def clean_plate(text):
    if pd.isna(text): return ""
    text = str(text).replace(" ", "").replace("-", "").strip()
    match = re.search(r'(\d{6,7}|\d?[ก-ฮ]{2,3}\d{1,4})', text)
    return match.group(1) if match else text[:7]

def parse_thai_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        if isinstance(date_val, datetime):
            if date_val.year > 2500:
                return date_val.replace(year=date_val.year - 543)
            return date_val
        
        date_str = str(date_val).strip()
        for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%d-%m-%y', '%d/%m/%y'):
            try:
                dt = datetime.strptime(date_str, fmt)
                if dt.year > 2500:
                    dt = dt.replace(year=dt.year - 543)
                return dt
            except:
                continue
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except:
        return pd.NaT

# --- ส่วนที่ 0: เครื่องมือเตรียมไฟล์ (ปุ่มอัปเดต/จัดระเบียบข้อมูล) ---
st.subheader("🛠️ ส่วนที่ 0: เครื่องมือเตรียมไฟล์ (จัดระเบียบข้อมูลก่อน)")
with st.expander("🪄 คลิกเพื่อใช้งานเครื่องมือจัดกลุ่มทะเบียนรถ (Input_Service_Sorted)"):
    st.info("ใช้สำหรับจัดกลุ่มรถทะเบียนเดียวกันให้อยู่ติดกัน และเรียงวันที่จากเก่าไปใหม่ เพื่อให้ดูง่ายและระบบแม่นยำขึ้น")
    prep_file = st.file_uploader("อัปโหลดไฟล์ Input_Service_Data.xlsx เพื่อจัดระเบียบ", type=['xlsx'], key="prep_tool")
    
    if prep_file:
        if st.button("🚀 กดจัดกลุ่มข้อมูล"):
            try:
                # อ่านไฟล์เดิม (ข้าม 2 แถวแรกตามโครงสร้าง CMS)
                df_prep = pd.read_excel(prep_file, skiprows=2)
                df_prep.columns = df_prep.columns.str.strip()
                
                # หาคอลัมน์วันที่
                p_date_col = 'วันที่เข้าศูนย์บริการ'
                if p_date_col not in df_prep.columns:
                    for c in df_prep.columns:
                        if 'วันที่' in str(c): p_date_col = c; break
                
                # เรียงข้อมูล: ทะเบียนรถ (ก-ฮ) และ วันที่ (เก่า -> ใหม่)
                df_prep['tmp_date'] = df_prep[p_date_col].apply(parse_thai_date)
                df_prep = df_prep.sort_values(by=['ป้ายทะเบียนรถ', 'tmp_date'], ascending=[True, True])
                df_prep = df_prep.drop(columns=['tmp_date'])
                
                # สร้างไฟล์ส่งกลับ
                output_prep = io.BytesIO()
                with pd.ExcelWriter(output_prep, engine='xlsxwriter') as writer:
                    df_prep.to_excel(writer, index=False, startrow=2) # เขียนเริ่มแถวที่ 3
                
                st.success("✅ จัดระเบียบเสร็จแล้ว! กรุณาดาวน์โหลดไฟล์นี้ไปใช้ในขั้นตอนที่ 1 ด้านล่าง")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ที่จัดระเบียบแล้ว (Sorted)",
                    data=output_prep.getvalue(),
                    file_name="Input_Service_Sorted.xlsx",
                    mime="application/vnd.ms-excel"
                )
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")

st.divider()

# --- ส่วนของระบบหลัก (Main System) ---
with st.sidebar:
    st.header("📂 เมนูจัดการข้อมูลระบบหลัก")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (ใช้ไฟล์ที่ Sorted แล้ว)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน (Mileage)", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขการเปลี่ยนอะไหล่ (Logic)", type=['xlsx'])
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลระบบ")

if process_btn:
    if not (file_input and file_mileage and file_logic):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบ!")
    else:
        try:
            with st.spinner('กำลังคำนวณ...'):
                df_logic = pd.read_excel(file_logic)
                df_m = pd.read_excel(file_mileage, header=2)
                df_new = pd.read_excel(file_input, skiprows=2)
                df_new.columns = df_new.columns.str.strip()

                # หาคอลัมน์วันที่
                target_date_col = 'วันที่เข้าศูนย์บริการ'
                if target_date_col not in df_new.columns:
                    for c in df_new.columns:
                        if 'วันที่' in str(c): target_date_col = c; break

                # ดึงไมล์ล่าสุด
                df_m['Key'] = df_m['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
                mileage_dict = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('Key')['เลขไมล์สิ้นสุด'].last().to_dict()
                
                df_new = df_new.dropna(subset=['ป้ายทะเบียนรถ'])
                df_new['ทะเบียน_Clean'] = df_new['ป้ายทะเบียนรถ'].apply(clean_plate)
                df_new['ไมล์ปัจจุบัน'] = df_new['ทะเบียน_Clean'].map(mileage_dict)

                # คำนวณ Logic
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
                
                # แจ้งเตือน 3 สี
                today = datetime.now()
                def get_status(row):
                    if pd.isna(row['ไมล์ปัจจุบัน']) or row['ไมล์ปัจจุบัน'] == 0: return "🔍 ไม่พบข้อมูลไมล์"
                    diff_km = row['ไมล์นัดหมาย'] - row['ไมล์ปัจจุบัน']
                    diff_days = (row['วันที่นัดหมาย'] - today).days if pd.notna(row['วันที่นัดหมาย']) else 999
                    if diff_km <= 0 or diff_days <= 0: return f"🔴 ถึงกำหนด ({row['รายการ']})"
                    elif diff_km <= 1000 or diff_days <= 15: return f"🟡 ใกล้ถึง (เหลือ {int(diff_km):,} กม.)"
                    else: return "🟢 ปกติ"

                df_new['สถานะการแจ้งเตือน'] = df_new.apply(get_status, axis=1)
                
                # แปลงตัวเลขเป็น Int
                for col in ['ไมล์ปัจจุบัน', 'ไมล์นัดหมาย', 'ไมล์ที่เข้าล่าสุด']:
                    df_new[col] = pd.to_numeric(df_new[col], errors='coerce').fillna(0).astype(int)

            # --- ส่วนแสดงผล ---
            c1, c2, c3, c4 = st.columns(4)
            num_red = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🔴")])
            num_yellow = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🟡")])
            num_green = len(df_new[df_new['สถานะการแจ้งเตือน'].str.contains("🟢")])
            c1.metric("รถทั้งหมด", f"{len(df_new)} คัน")
            c2.metric("ต้องซ่อมด่วน", f"{num_red} คัน", delta=num_red, delta_color="inverse")
            c3.metric("ใกล้ถึงกำหนด", f"{num_yellow} คัน")
            c4.metric("สถานะปกติ", f"{num_green} คัน")

            col_chart, col_table = st.columns([1, 2])
            with col_chart:
                fig = px.pie(names=['ถึงกำหนด', 'ใกล้ถึง', 'ปกติ'], values=[num_red, num_yellow, num_green],
                             color=['ถึงกำหนด', 'ใกล้ถึง', 'ปกติ'],
                             color_discrete_map={'ถึงกำหนด':'#EF5350', 'ใกล้ถึง':'#FFB74D', 'ปกติ':'#66BB6A'})
                st.plotly_chart(fig, use_container_width=True)

            with col_table:
                st.write("📋 ตารางสถานะ")
                output_cols = ['ป้ายทะเบียนรถ', 'วันที่เข้าล่าสุด', 'ไมล์ที่เข้าล่าสุด', 'ไมล์ปัจจุบัน', 'สถานะการแจ้งเตือน']
                df_show = df_new[output_cols].copy()
                df_show['วันที่เข้าล่าสุด'] = df_show['วันที่เข้าล่าสุด'].apply(lambda x: f"{x.day:02d}/{x.month:02d}/{x.year + 543}" if pd.notnull(x) else "ไม่ระบุ")
                
                def color_row(val):
                    if '🔴' in val: return 'background-color: #ffebee'
                    if '🟡' in val: return 'background-color: #fff3e0'
                    if '🟢' in val: return 'background-color: #e8f5e9'
                    return ''
                
                st.dataframe(df_show.style.applymap(color_row, subset=['สถานะการแจ้งเตือน'])
                             .format({"ไมล์ที่เข้าล่าสุด": "{:,d}", "ไมล์ปัจจุบัน": "{:,d}"}), use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_new.to_excel(writer, index=False)
            st.download_button("📥 ดาวน์โหลดรายงาน (.xlsx)", data=buffer.getvalue(), file_name="Service_Audit_Final.xlsx")

        except Exception as e:
            st.error(f"❌ Error: {e}")
