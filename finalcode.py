import streamlit as st
import pandas as pd
import win32com.client as win32 # สำหรับคุม Outlook
import io
from datetime import datetime

# --- ฟังก์ชันสำหรับเปิด Outlook Preview ---
def preview_outlook_email(to_email, cc_email, subject, body):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.CC = cc_email
        mail.Subject = subject
        mail.HTMLBody = body
        # ใช้ .Display() เพื่อให้หน้าต่างเด้งขึ้นมาพรีวิว (ไม่ส่งทันที)
        mail.Display()
        return True
    except Exception as e:
        st.error(f"ไม่สามารถเปิด Outlook ได้: {e}")
        return False

# --- ปรับปรุงส่วน UI การอัปโหลด (Sidebar) ---
with st.sidebar:
    st.header("📂 เมนูจัดการข้อมูลระบบหลัก")
    file_input = st.file_uploader("1. ข้อมูลการเข้าศูนย์ (ใช้ไฟล์ที่ Sorted แล้ว)", type=['xlsx'])
    file_mileage = st.file_uploader("2. ข้อมูลเลขไมล์ปัจจุบัน (Mileage)", type=['xlsx'])
    file_logic = st.file_uploader("3. เงื่อนไขการเปลี่ยนอะไหล่ (Logic)", type=['xlsx'])
    file_email_config = st.file_uploader("4. ไฟล์เงื่อนไข Email (Email.xlsx)", type=['xlsx']) # <--- เพิ่มช่องนี้
    st.divider()
    process_btn = st.button("🚀 เริ่มประมวลผลระบบ")

# --- ภายใน Logic หลังกดปุ่ม process_btn ---
if process_btn:
    if not (file_input and file_mileage and file_logic and file_email_config):
        st.error("⚠️ กรุณาอัปโหลดไฟล์ให้ครบรวมถึง Email.xlsx ด้วยครับ!")
    else:
        # ... (โค้ดส่วนคำนวณเดิมของพี่) ...
        # สมมติว่าได้ df_new ออกมาแล้ว
        
        # --- [ส่วนที่เพิ่มใหม่] Logic การเชื่อมข้อมูล Email ---
        df_email = pd.read_excel(file_email_config, sheet_name='เงื่อนไข') 
        # ปรับชื่อคอลัมน์ให้ตรงกับไฟล์พี่ (Name, to, CC)
        df_email.columns = ['Name', 'Email_To', 'Email_CC'] 
        
        # Map ข้อมูล Email เข้ากับ df_new (โดยใช้ชื่อคนขับ/ผู้รับผิดชอบเป็น Key)
        # หมายเหตุ: ใน df_new พี่ต้องมีคอลัมน์ที่เก็บชื่อที่ตรงกับในไฟล์ Email.xlsx นะครับ
        df_final = pd.merge(df_new, df_email, left_on='ชื่อผู้รับผิดชอบ', right_on='Name', how='left')

        # --- ส่วนแสดงผลปุ่มส่งเมล ---
        st.subheader("📧 ระบบส่งการแจ้งเตือนผ่าน Email")
        
        # กรองเฉพาะรถที่ "ถึงกำหนด" (สีแดง) หรือ "ใกล้ถึง" (สีเหลือง)
        df_to_alert = df_final[df_final['สถานะการแจ้งเตือน'].str.contains("🔴|🟡")]

        if not df_to_alert.empty:
            for index, row in df_to_alert.iterrows():
                col_info, col_btn = st.columns([4, 1])
                
                # ข้อมูลที่จะใส่ใน Email
                plate = row['ป้ายทะเบียนรถ']
                status = row['สถานะการแจ้งเตือน']
                to_addr = row['Email_To'] if pd.notna(row['Email_To']) else ""
                cc_addr = row['Email_CC'] if pd.notna(row['Email_CC']) else ""
                
                col_info.write(f"🚗 **ทะเบียน:** {plate} | **สถานะ:** {status} | **ส่งถึง:** {to_addr}")
                
                if col_btn.button(f"Preview เมลคันนี้ ({plate})", key=f"btn_{index}"):
                    if to_addr == "":
                        st.warning(f"คัน {plate} ไม่มีข้อมูล Email ในไฟล์เงื่อนไข")
                    else:
                        subject = f"แจ้งเตือนกำหนดซ่อมบำรุงรถยนต์ - ทะเบียน {plate}"
                        body = f"""
                        <html>
                            <body>
                                <h3>เรียน คุณ {row['Name']}</h3>
                                <p>ขอแจ้งเตือนสถานะการเข้าซ่อมบำรุงรถยนต์ทะเบียน <b>{plate}</b></p>
                                <p>สถานะปัจจุบัน: <span style="color:red;">{status}</span></p>
                                <p>เลขไมล์ล่าสุด: {row['ไมล์ที่เข้าล่าสุด']:,} กม.</p>
                                <p><b>กรุณานำรถเข้าศูนย์บริการตามกำหนด</b></p>
                                <hr>
                                <p>ระบบ Smart Alert (CMS Service Audit)</p>
                            </body>
                        </html>
                        """
                        preview_outlook_email(to_addr, cc_addr, subject, body)
                        st.success(f"เปิด Outlook พรีวิวสำหรับทะเบียน {plate} แล้ว")
        else:
            st.info("ไม่มีรถที่ต้องส่งแจ้งเตือนในขณะนี้")
