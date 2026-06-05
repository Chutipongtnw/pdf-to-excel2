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
    
    # ซ่อมคำเฉพาะ 5 วิชาโดยตรง ไม่กระทบส่วนอื่น
    text = text.replace('สราง', 'สร้าง')
    text = text.replace('รู', 'รู้')
    text = text.replace('หนา', 'หน้า')
    text = text.replace('ตน', 'ต้น')
    text = text.replace('ป่ญญา', 'ปัญญา')
    
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
    
    # 6. ยุบสระที่เบิ้ล (เเ, แแ, าา)
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    # 7. คลังซ่อมคำมาตรฐาน
    if 'พิ่มเติม' in text and 'เพิ่ม' not in text:
        text = text.replace('พิ่มเติม', 'เพิ่มเติม')

    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยา官方',
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
        'เ์': '์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    # 8. ยุบวรรณยุกต์ซ้ำ
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 9. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

def clean_invisible_and_spaces(text):
    """ฟังก์ชันสำหรับลบช่องว่างและอักขระที่มองไม่เห็นทั้งหมด"""
    if not text: return ""
    text = str(text).replace('\n', '')
    # ลบ Unicode Control Characters ขยะ และช่องว่างทุกประเภทออกทั้งหมด
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0\s]', '', text)
    return text.strip()

st.set_page_config(page_title="ระบบดึงข้อมูลอัจฉริยะ v55", layout="wide")
st.title("📂 ระบบดึงข้อมูล PDF เป็น Excel v55 (ปรับตามโครงสร้างใหม่)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF เพื่อรัน v55", type="pdf")

if uploaded_file is not None:
    all_data = []
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            raw_text = page.extract_text() or ""
            
            # 1. ดึงรหัสครู (หาเฉพาะตัวเลขที่อยู่ในวงเล็บท้ายหน้า)
            teacher_id = "N/A"
            t_match = re.search(r'\((\d+)\)', raw_text)
            if t_match: 
                teacher_id = t_match.group(1)

            # 2. ดึงรหัสวิชา (หาข้อความถัดจากคำว่า 'รหัสวิชา' และตัดสิ่งที่ไม่เกี่ยวข้องออก)
            subject_code = "N/A"
            code_match = re.search(r'รหัสวิชา\s*([^\s]+)', raw_text)
            if code_match:
                subject_code = clean_invisible_and_spaces(code_match.group(1))

            # 3. ดึงชื่อวิชา (หาข้อความถัดจากคำว่า 'ชื่อวิชา')
            subject_name = "N/A"
            name_match = re.search(r'ชื่อวิชา\s*(.+)', raw_text)
            if name_match:
                line_text = name_match.group(1)
                # ตัดคำคีย์เวิร์ดอื่นๆ ที่อาจพ่วงมาในบรรทัดเดียวกันออกไปก่อน
                for kw in ["รหัสวิชา", "ภาคเรียน", "ปีการศึกษา", "ชั้น", "ระดับชั้น"]:
                    if kw in line_text:
                        line_text = line_text.split(kw)[0]
                subject_name = universal_thai_cleaner(line_text)

            # 4. ดึงข้อมูลตารางคะแนน/เกรดนักเรียน
            table = page.extract_table()
            if table:
                # กำหนดค่าดัชนีคอลัมน์เริ่มต้น (กรณีหาหัวตารางไม่เจอจะใช้ค่านี้เป็น Default)
                col_student_id = 1  # เลขประจำตัว
                col_remark = 4      # หมายเหตุ (ระดับชั้น)
                col_grade = 7       # เกรดปกติ
                
                # ตรวจสอบหาตำแหน่งคอลัมน์ที่แท้จริงแบบไดนามิกจากหัวตารางเพื่อความแม่นยำสูงสุด
                for r_idx in range(min(5, len(table))):
                    row_cleaned = [str(cell).replace('\n', '').replace(' ', '') if cell else "" for cell in table[r_idx]]
                    if any("เลขประจำตัว" in cell for cell in row_cleaned) or any("เกรด" in cell for cell in row_cleaned):
                        for c_idx, cell in enumerate(row_cleaned):
                            if "เลขประจำตัว" in cell:
                                col_student_id = c_idx
                            elif "หมายเหตุ" in cell:
                                col_remark = c_idx
                            elif "เกรด" in cell:
                                col_grade = c_idx
                        break

                # วนลูปอ่านข้อมูลนักเรียนรายแถว
                for row in table:
                    if not row or len(row) <= max(col_student_id, col_remark, col_grade):
                        continue
                        
                    s_id = clean_invisible_and_spaces(row[col_student_id])
                    
                    # คัดกรองเอาเฉพาะแถวที่เป็นรหัสนักเรียนจริงๆ (เป็นตัวเลขและไม่เว้นว่าง)
                    if s_id.isdigit() and len(s_id) >= 3:
                        remark_val = clean_invisible_and_spaces(row[col_remark])
                        grade_val = clean_invisible_and_spaces(row[col_grade])
                        
                        all_data.append({
                            "เลขประจำตัวนักเรียน": s_id,
                            "รหัสวิชา": subject_code,
                            "ชื่อวิชา": subject_name,
                            "ระดับชั้น": remark_val,  # นำข้อมูลจากคอลัมน์ "หมายเหตุ" มาใส่
                            "เกรดปกติ": grade_val,    # นำข้อมูลจากคอลัมน์ "เกรดปกติ" มาใส่
                            "รหัสครู": teacher_id
                        })
            progress_bar.progress((i + 1) / len(pdf.pages))

    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        
        # แทรกคอลัมน์ "ที่" เรียงลำดับ 1, 2, 3... ไว้หน้าสุดของตาราง
        df.insert(0, "ที่", range(1, len(df) + 1))
        
        st.success(f"⚡ ดึงข้อมูลสำเร็จ! พบข้อมูลทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        # เขียนข้อมูลลงไฟล์ Excel ด้วย xlsxwriter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (v55)",
            data=output.getvalue(),
            file_name="student_report_v55.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
