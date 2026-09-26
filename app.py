import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="User Summary Report", layout="wide")

st.title("📊 ระบบสรุปยอดผู้ใช้งานรายวัน (พร้อมผูกสูตร Excel)")
st.markdown("อัปโหลดไฟล์ที่มีข้อมูล **วันที่ทำข้อมูล**, **บัญชีสมาชิก**, และ **NEW/OLD** ระบบจะเตรียม RawData และผูกสูตรหน้า Summary ให้ตอนดาวน์โหลด")

uploaded_file = st.file_uploader("เลือกไฟล์ CSV หรือ Excel", type=['csv', 'xlsx', 'xls'])

@st.cache_data
def process_data(df):
    required_cols = ['วันที่ทำข้อมูล', 'บัญชีสมาชิก', 'NEW/OLD']
    for col in required_cols:
        if col not in df.columns:
            st.error(f"❌ ไม่พบคอลัมน์ '{col}' ในไฟล์")
            return None, None

    # แปลงวันที่เป็น Text รูปแบบ YYYY-MM-DD เพื่อให้สูตร SUMIFS ใน Excel จับคู่ได้แม่นยำที่สุด
    df['Date'] = pd.to_datetime(df['วันที่ทำข้อมูล'], dayfirst=True, errors='coerce').dt.strftime('%Y-%m-%d')
    df.dropna(subset=['Date'], inplace=True)
    
    df['User'] = df['บัญชีสมาชิก'].astype(str).str.strip()
    df['Status'] = df['NEW/OLD'].astype(str).str.strip()

    # 1. เรียงลำดับข้อมูล โดยให้ความสำคัญกับ "ใหม่" มากกว่า "เก่า" (เพื่อจัดการกรณี 1 วันมี 2 สถานะ)
    df['Status_Rank'] = df['Status'].map({'ใหม่': 1, 'เก่า': 2}).fillna(3)
    df = df.sort_values(by=['Date', 'User', 'Status_Rank'])

    # 2. สร้าง Helper Columns สำหรับผูกสูตรใน Excel
    # นับว่านี่เป็นการปรากฏตัวครั้งแรกของวันหรือไม่ (1 = ใช่, 0 = ไม่)
    df['นับรายวัน (ซ้ำ)'] = (~df.duplicated(subset=['Date', 'User'])).astype(int)
    
    # นับว่านี่เป็นการปรากฏตัวครั้งแรกสุดในระบบหรือไม่
    is_global_first = ~df.duplicated(subset=['User'])
    df['นับครั้งแรก (ใหม่)'] = (is_global_first & (df['Status'] == 'ใหม่')).astype(int)
    df['นับครั้งแรก (เก่า)'] = (is_global_first & (df['Status'] == 'เก่า')).astype(int)

    # เตรียม Raw Data สำหรับ Export
    raw_output = df[['Date', 'User', 'Status', 'นับรายวัน (ซ้ำ)', 'นับครั้งแรก (ใหม่)', 'นับครั้งแรก (เก่า)']].copy()
    raw_output.rename(columns={'Date': 'วันที่', 'User': 'บัญชีสมาชิก', 'Status': 'NEW/OLD'}, inplace=True)

    # 3. เตรียมตาราง Summary สำหรับพรีวิวบนเว็บ (คำนวณสดด้วย Pandas)
    daily_total = raw_output.groupby('วันที่')['นับรายวัน (ซ้ำ)'].sum().astype(int)
    new_new = raw_output.groupby('วันที่')['นับครั้งแรก (ใหม่)'].sum().astype(int)
    new_old = raw_output.groupby('วันที่')['นับครั้งแรก (เก่า)'].sum().astype(int)
    
    summary = pd.DataFrame({
        'คนเดิมพัน(ซ้ำ)': daily_total,
        'คนเดิมพัน(ไม่ซ้ำ) - ใหม่': new_new,
        'คนเดิมพัน(ไม่ซ้ำ) - เก่า': new_old
    }).reset_index()
    
    summary['คนเดิมพัน(ไม่ซ้ำ)'] = summary['คนเดิมพัน(ไม่ซ้ำ) - ใหม่'] + summary['คนเดิมพัน(ไม่ซ้ำ) - เก่า']
    
    # จัดเรียงคอลัมน์ให้ตรงตามต้องการ
    summary = summary[['วันที่', 'คนเดิมพัน(ซ้ำ)', 'คนเดิมพัน(ไม่ซ้ำ)', 'คนเดิมพัน(ไม่ซ้ำ) - ใหม่', 'คนเดิมพัน(ไม่ซ้ำ) - เก่า']]
    
    return raw_output, summary

def to_excel_with_formulas(raw_df, summary_df):
    output = io.BytesIO()
    # ใช้ xlsxwriter เพื่อให้สามารถเขียนสูตร Excel (write_formula) ได้
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # นำข้อมูลลงชีต
        raw_df.to_excel(writer, index=False, sheet_name='RawData')
        summary_df.to_excel(writer, index=False, sheet_name='Summary')
        
        workbook = writer.book
        ws_summary = writer.sheets['Summary']
        
        # จัดความกว้างคอลัมน์ให้สวยงาม
        ws_summary.set_column('A:A', 15)
        ws_summary.set_column('B:E', 25)
        
        # วนลูปเพื่อเขียนสูตร Excel ลงไปทับตัวเลขในหน้า Summary
        # แถวใน Excel เริ่มต้นที่ 0 (แถว 0 คือ Header, แถว 1 คือบรรทัดแรกของข้อมูล)
        for row_idx in range(len(summary_df)):
            excel_row = row_idx + 2  # หมายเลขแถวจริงใน Excel (เริ่มที่แถว 2)
            
            # คอลัมน์ B: คนเดิมพัน(ซ้ำ) = SUMIFS ของคอลัมน์ D ใน RawData
            formula_b = f'=SUMIFS(RawData!D:D, RawData!A:A, A{excel_row})'
            ws_summary.write_formula(row_idx + 1, 1, formula_b)
            
            # คอลัมน์ D: คนเดิมพัน(ไม่ซ้ำ) - ใหม่ = SUMIFS ของคอลัมน์ E ใน RawData
            formula_d = f'=SUMIFS(RawData!E:E, RawData!A:A, A{excel_row})'
            ws_summary.write_formula(row_idx + 1, 3, formula_d)
            
            # คอลัมน์ E: คนเดิมพัน(ไม่ซ้ำ) - เก่า = SUMIFS ของคอลัมน์ F ใน RawData
            formula_e = f'=SUMIFS(RawData!F:F, RawData!A:A, A{excel_row})'
            ws_summary.write_formula(row_idx + 1, 4, formula_e)
            
            # คอลัมน์ C: คนเดิมพัน(ไม่ซ้ำ) (รวม) = ใหม่ + เก่า
            formula_c = f'=D{excel_row}+E{excel_row}'
            ws_summary.write_formula(row_idx + 1, 2, formula_c)

    output.seek(0)
    return output

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("✅ โหลดไฟล์สำเร็จ! กำลังประมวลผล...")
        
        raw_output, result_summary = process_data(df)
        
        if result_summary is not None:
            st.markdown("### 📈 ผลลัพธ์การสรุปข้อมูล (Preview)")
            st.dataframe(result_summary, use_container_width=True)
            
            # สร้างไฟล์ Excel พร้อมผูกสูตร
            excel_data = to_excel_with_formulas(raw_output, result_summary)
            
            st.download_button(
                label="📥 ดาวน์โหลดผลลัพธ์เป็น Excel (RawData + สูตร Summary)",
                data=excel_data,
                file_name="user_summary_with_formulas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")