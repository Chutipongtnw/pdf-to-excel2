import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import unicodedata

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกทันที
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. Normalization ขั้นสูงสุด
    text = unicodedata.normalize('NFKC', text)
    
    # 3. Unicode Mapping และซ่อมคำเฉพาะหน้า
    text = text.replace('ค', 'ค้น').replace('คว', 'คว้า')
    text = text.replace('ด', 'ด้')
    
    # ซ่อมคำเฉพาะวิชาโดยตรง
    text = text.replace('สราง', 'สร้าง')
    text = text.replace('รู', 'รู้')
    text = text.replace('รู', 'รู้')
    text = text.replace('หนา', 'หน้า')
    text = text.replace('ตน', 'ต้น')
    text = text.replace('ป่ญญา', 'ปัญญา')
    text = text.replace('ป่จจุบัน', 'ปัจจุบัน')
    
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 4. ล้างรหัส Unicode ขยะและช่องว่างทั้งหมด
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    text = "".join(text.split())

    # 5. แก้ไขพยัญชนะเบิ้ลและสระกระโดด
    text = text.replace('ศลิป', 'ศิลป์') 
    text = text.replace('ฟิสกิ', 'ฟิสิก') 
    text = text.replace('ต่ออ', 'ต่อ')
    text = text.replace('ค้นน', 'ค้น')
    text = text.replace('ข้ออ', 'ข้อ')
    text = text.replace('แผ่นน', 'แผ่น')
    text = text.replace('นำา', 'นำ')
    text = text.replace('ฟ่ง', 'ฟัง')
    text = text.replace('อ่านาน', 'อ่าน') 
    
    # 6. ยุบสระที่เบิ้ล
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    # 7. คลังซ่อมคำมาตรฐาน
    if 'พิ่มเติม' in text and 'เพิ่ม' not in text:
        text = text.replace('พิ่มเติม', 'เพิ่มเติม')

    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยาศาสตร์',
        'นาฏศิลป1': 'นาฏศิลป์ 1',
        'นาฏศิลป2': 'นาฏศิลป์ 2',
        'ทัศนศิลป1': 'ทัศนศิลป์ 1',
        'ทัศนศิลป2': 'ทัศนศิลป์ 2',
        'ฟิสิกส': 'ฟิสิกส์',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'ผลติ': 'ผลิต',
        'คาสตร์': 'ศาสตร์',
        'วดีโอ': 'วิดีโอ',
        'เ์': '์',
        'ป่ญญา': 'ปัญญา',
        'ป่จจุบัน': 'ปัจจุบัน',
        'องค์ความรู': 'องค์ความรู้',
        'ความรู': 'ความรู้',
        'รู': 'รู้'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 8. ยุบวรรณยุกต์ซ้ำ
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 9. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    # 10. ตัดตัวเลขลำดับหน้าชื่อวิชาออก
    text = re.sub(r'^\s*\d+\s*\.\s*', '', text)

    return text.strip()

def clean_invisible_and_spaces(text):
    """ฟังก์ชันสำหรับลบช่องว่างและอักขระที่มองไม่เห็นทั้งหมด"""
    if text is None: return ""
    text = str(text).replace('\n', '')
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0\s]', '', text)
    return text.strip()

st.set_page_config(page_title="ระบบแปลงไฟล์ PDF to Excel เฉพาะไฟล์ผลการเรียนบกพร่อง รายวิชา ธนว เท่านั้น", layout="wide")
st.title("📂 ระบบแปลงไฟล์ PDF to Excel เฉพาะไฟล์ผลการเรียนบกพร่อง รายวิชา ธนว เท่านั้น")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF เพื่อแปลงเป็น Excel", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. ดึงรหัสครู
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: 
                teacher_id = t_match.group(1)

            # 2. ดึงรหัสวิชา
            subject_code = "N/A"
            code_match = re.search(r'รหัสวิชา\s*([^\s]+)', raw_text)
            if code_match:
                subject_code = clean_invisible_and_spaces(code_match.group(1))

            # 3. ดึงชื่อวิชา
            subject_name = "N/A"
            name_match = re.search(r'ชื่อวิชา\s*(.+)', raw_text)
            if name_match:
                line_text = name_match.group(1)
                for kw in ["รหัสวิชา", "ภาคเรียน", "ปีการศึกษา", "ชั้น", "ระดับชั้น"]:
                    if kw in line_text:
                        line_text = line_text.split(kw)[0]
                subject_name = universal_thai_cleaner(line_text)

            # 4. ดึงข้อมูลตารางคะแนน
            table = page.extract_table()
            if table:
                # ค่าดัชนีเริ่มต้น
                col_student_id = 1
                col_remark = 4
                col_grade = 7
                
                # --- ระบบ Super Header: รวมข้อความ 4 บรรทัดแรกเผื่อหัวตารางถูกปัดบรรทัด ---
                super_header = [""] * 20
                for r_idx in range(min(5, len(table))):
                    if not table[r_idx]: continue
                    for c_idx, cell in enumerate(table[r_idx]):
                        if cell and c_idx < 20:
                            super_header[c_idx] += clean_invisible_and_spaces(str(cell))
                
                # ล็อกตำแหน่งคอลัมน์จาก Super Header
                for c_idx, text in enumerate(super_header):
                    if "เลขประจำตัว" in text or "รหัสนักเรียน" in text:
                        col_student_id = c_idx
                    elif "หมายเหตุ" in text or "ระดับชั้น" in text:
                        col_remark = c_idx
                    elif any(kw in text for kw in ["เกรด", "ผลการเรียน", "ปกติ", "ประเมิน"]):
                        col_grade = c_idx

                # วนลูปอ่านข้อมูลนักเรียนรายแถว
                for row in table:
                    if not row or len(row) <= col_student_id:
                        continue
                        
                    s_id = clean_invisible_and_spaces(str(row[col_student_id]))
                    
                    if s_id.isdigit() and len(s_id) >= 3:
                        remark_val = clean_invisible_and_spaces(str(row[col_remark])) if len(row) > col_remark else ""
                        grade_val = clean_invisible_and_spaces(str(row[col_grade])) if len(row) > col_grade else ""
                        
                        # --- ระบบ Failsafe: หากช่องเกรดหลักว่างเปล่า ให้กวาดหาตัวเลขเกรดในช่องใกล้เคียง ---
                        if not grade_val:
                            # รายชื่อเกรดที่เป็นไปได้ทั้งหมด
                            valid_grades = ["4", "4.0", "3.5", "3", "3.0", "2.5", "2", "2.0", "1.5", "1", "1.0", "0", "0.0", "ร", "มส", "ผ", "มผ", "ขร", "ผ่าน", "ไม่ผ่าน"]
                            # กวาดจากคอลัมน์หลังสุดย้อนมาด้านหน้า
                            for idx in range(len(row)-1, col_student_id, -1):
                                val = clean_invisible_and_spaces(str(row[idx]))
                                if val in valid_grades:
                                    grade_val = val
                                    break # ถ้าเจอเกรดที่ใช่ ให้ดึงมาใช้แล้วหยุดกวาดทันที
                        
                        all_data.append({
                            "เลขประจำตัวนักเรียน": s_id,
                            "รหัสวิชา": subject_code,
                            "ชื่อวิชา": subject_name,
                            "ระดับชั้น": remark_val,  
                            "เกรดปกติ": grade_val,    
                            "รหัสครู": teacher_id
                        })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        
        st.success(f"⚡ ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        
        # แสดงผลตารางโดยไม่มีคอลัมน์ "ที่"
        st.dataframe(df, use_container_width=True)
        
        # เขียนข้อมูลลงไฟล์ Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel",
            data=output.getvalue(),
            file_name="student_report_v58.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
