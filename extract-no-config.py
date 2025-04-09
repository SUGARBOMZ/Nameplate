import os
import base64
import pandas as pd
import streamlit as st
import requests
import json
from PIL import Image
import io

# ฟังก์ชันสำหรับแปลงรูปภาพเป็น base64
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

# ฟังก์ชันสำหรับเรียกใช้ Gemini API
def extract_data_from_image(api_key, image_data, prompt_text):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt_text},
                    {
                        "inline_data": {
                            "mime_type": "image/jpeg", 
                            "data": image_data
                        }
                    }
                ]
            }
        ],
        "generation_config": {
            "temperature": 0.2,
            "top_p": 0.82,
            "max_output_tokens": 4000
        }
    }
    
    headers = {
        "Content-Type": "application/json"
    }
    
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    
    if response.status_code == 200:
        response_data = response.json()
        # ดึงข้อมูลจาก response
        if 'candidates' in response_data and len(response_data['candidates']) > 0:
            text_content = response_data['candidates'][0]['content']['parts'][0]['text']
            return text_content
        else:
            return "ไม่พบข้อมูลในการตอบกลับ"
    else:
        return f"เกิดข้อผิดพลาด: {response.status_code} - {response.text}"

# สร้าง UI ด้วย Streamlit
st.title("ระบบสกัดข้อมูลจากรูปภาพ")

# รับ API key
api_key = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

# กำหนดค่า prompt โดยตรง
prompt_text = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ โดยเอาเฉพาะ field ดังต่อไปนี้
1: Manufacturer
2: Serial Number
3: Standard (such as IEC, IEEE, ANSI etc.)
4: Rated Capacity (kVA)
5: Rated Voltage (kV) High Side
6: Rated Voltage (kV) Low Side
7: Rated Current (kV) High Side
8: Rated Current (kV) Low Side
9: Impedance Voltage (%)
10: Impedance Voltage (%)
11: Vector Group"""
output_format = "JSON"

# อัปโหลดหลายไฟล์
uploaded_files = st.file_uploader("อัปโหลดรูปภาพ (เลือกได้หลายรูป โดยการกด ctrl + คลิกทีละรูป)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# เมื่อผู้ใช้กดปุ่มประมวลผล
if st.button("ประมวลผล") and api_key and uploaded_files:
    if not uploaded_files:
        st.error("กรุณาอัปโหลดไฟล์รูปภาพ")
    else:
        # แสดง progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        results = []
        
        for i, file in enumerate(uploaded_files):
            # อัปเดต progress
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            
            # แสดงสถานะการทำงาน
            status_text.text(f"กำลังประมวลผล {i+1}/{len(uploaded_files)}: {file.name}")
            
            # แปลงรูปภาพเป็น base64
            image_data = encode_image(file)
            
            # เรียกใช้ Gemini API
            response = extract_data_from_image(api_key, image_data, prompt_text)
            
            # พยายามแปลง response เป็น JSON
            extracted_data = {}
            try:
                # ค้นหาส่วนที่เป็น JSON ในข้อความตอบกลับ
                json_start = response.find('{')
                json_end = response.rfind('}') + 1
                
                if json_start >= 0 and json_end > json_start:
                    json_str = response[json_start:json_end]
                    extracted_data = json.loads(json_str)
                else:
                    extracted_data = {"raw_text": response}
            except json.JSONDecodeError:
                extracted_data = {"raw_text": response}
            
            # เพิ่มข้อมูลลงในผลลัพธ์
            result_entry = {
                "file_name": file.name,
                "extracted_data": extracted_data if isinstance(extracted_data, dict) else {"raw_text": response}
            }
            
            results.append(result_entry)
        
        # เมื่อประมวลผลเสร็จสิ้น
        status_text.text("ประมวลผลเสร็จสิ้น!")
        
        # จัดเตรียมข้อมูลสำหรับ Excel
        excel_data = []
        
        for result in results:
            if isinstance(result["extracted_data"], dict) and "raw_text" not in result["extracted_data"]:
                # นำข้อมูลจาก JSON มาขยายออก
                row_data = {"file_name": result["file_name"]}
                row_data.update(result["extracted_data"])
                excel_data.append(row_data)
            else:
                # ใช้ข้อความทั้งหมด
                excel_data.append({
                    "file_name": result["file_name"],
                    "extracted_text": result["extracted_data"].get("raw_text", str(result["extracted_data"]))
                })
        
        # สร้าง DataFrame
        df = pd.DataFrame(excel_data)
        
        # แสดงตัวอย่างข้อมูล
        st.subheader("ตัวอย่างข้อมูลที่สกัดได้")
        st.dataframe(df)
        
        # สร้างไฟล์ Excel
        output_file = "extracted_data.xlsx"
        df.to_excel(output_file, index=False)
        
        # ให้ผู้ใช้ดาวน์โหลดไฟล์
        with open(output_file, "rb") as file:
            st.download_button(
                label="ดาวน์โหลดไฟล์ Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
