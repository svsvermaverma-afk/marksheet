import streamlit as st
import pandas as pd
import io
import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Page Configuration
st.set_page_config(page_title="ABIC Marksheet Generator", layout="wide")

# Register Hindi Font (Make sure font files are in the same folder)
FONT_REG = "NotoSansDevanagari-Regular.ttf"
FONT_BOLD = "NotoSansDevanagari-Bold.ttf"

if os.path.exists(FONT_REG):
    pdfmetrics.registerFont(TTFont("Hindi", FONT_REG))
    pdfmetrics.registerFont(TTFont("HindiBold", FONT_BOLD if os.path.exists(FONT_BOLD) else FONT_REG))
    HINDI_FONT = "Hindi"
    HINDI_BOLD = "HindiBold"
else:
    HINDI_FONT = "Helvetica"
    HINDI_BOLD = "Helvetica-Bold"
    st.warning("⚠️ 'NotoSansDevanagari-Regular.ttf' nahi mila. Hindi text PDF mein sahi render hone ke liye font file GitHub repo mein upload karein.")

# Helper: Calculate Grade from Percentage
def get_grade(marks, max_marks):
    if max_marks == 0:
        return ""
    pct = (marks / max_marks) * 100
    if pct >= 91: return "A1"
    elif pct >= 81: return "A2"
    elif pct >= 71: return "B1"
    elif pct >= 61: return "B2"
    elif pct >= 51: return "C1"
    elif pct >= 41: return "C2"
    elif pct >= 33: return "D"
    else: return "E"

# Generate Sample Excel Template
def get_sample_excel():
    data = [{
        "Roll_No": "44", "Reg_No": "38940", "Name": "Anandi Kumari Singh",
        "Father_Name": "Shriram Singh", "Mother_Name": "Lalita Devi",
        "DOB": "26-11-2009", "Class_Sec": "11 'B'", "House": "Arun",
        "Mobile": "9450320524", "Aadhaar": "982931653622", "PEN_No": "21473874271",
        "Address": "Jhopadi Near Bus Stand, Renukoot",
        # Half Yearly Marks (HY)
        "Hindi_HY": 80, "English_HY": 65, "Math_Bio_HY": 76,
        "Physics_HY_Th": 56, "Physics_HY_Pr": 29,
        "Chem_HY_Th": 38, "Chem_HY_Pr": 28,
        "Moral_Sports_Grade": "A",
        # Annual Marks (AN) - Leave 0 if generating only Half Yearly
        "Hindi_AN": 0, "English_AN": 0, "Math_Bio_AN": 0,
        "Physics_AN_Th": 0, "Physics_AN_Pr": 0,
        "Chem_AN_Th": 0, "Chem_AN_Pr": 0,
        # Co-Scholastic Grades (1 to 15)
        "CS_1": "A", "CS_2": "A", "CS_3": "B", "CS_4": "A", "CS_5": "A",
        "CS_6": "A", "CS_7": "A", "CS_8": "A", "CS_9": "A", "CS_10": "A",
        "CS_11": "A", "CS_12": "B", "CS_13": "A", "CS_14": "B", "CS_15": "A",
        "Attendance": "110/120", "Result": "Promoted", "Session": "2025 - 2026"
    }]
    df = pd.DataFrame(data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Students")
    return buffer.getvalue()

# PDF Drawing Function
def generate_marksheet_pdf(df):
    buffer = io.BytesIO()
    width, height = landscape(A4)  # 841.89 x 595.27 points
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    
    cream_bg = colors.Color(0.98, 0.96, 0.88)
    maroon = colors.Color(0.55, 0.15, 0.15)
    
    for _, row in df.iterrows():
        # ================= PAGE 1: OUTER COVER (BACK LEFT + FRONT RIGHT) =================
        c.setFillColor(cream_bg)
        c.rect(0, 0, width, height, fill=1, stroke=0)
        
        # Center Fold Line
        c.setStrokeColor(colors.grey)
        c.setDash(3, 3)
        c.line(width/2, 10, width/2, height-10)
        c.setDash([])
        
        # --- LEFT HALF: GRADING SYSTEM & INSTRUCTIONS ---
        c.setStrokeColor(maroon)
        c.setLineWidth(1.5)
        c.rect(20, 20, width/2 - 30, height - 40)
        
        c.setFillColor(maroon)
        c.setFont(HINDI_BOLD, 11)
        c.drawCentredString(width*0.25, height - 60, "शैक्षिक मूल्यांकन ग्रेडिंग सिस्टम (प्रतिशत में)")
        
        # Academic Grading Table
        grades_range = ["91-100", "81-90", "71-80", "61-70", "51-60", "41-50", "33-40", "0-32"]
        grades_letter = ["A1", "A2", "B1", "B2", "C1", "C2", "D", "E"]
        gx, gy, gw, gh = 40, height - 125, (width/2 - 80)/8, 25
        for i in range(8):
            c.rect(gx + i*gw, gy + gh, gw, gh)
            c.rect(gx + i*gw, gy, gw, gh)
            c.setFont("Helvetica", 8)
            c.drawCentredString(gx + i*gw + gw/2, gy + gh + 8, grades_range[i])
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(gx + i*gw + gw/2, gy + 8, grades_letter[i])
            
        # Co-Scholastic Grading Table
        c.setFont(HINDI_BOLD, 10)
        c.drawCentredString(width*0.25, height - 175, "जीवन कौशल व्यवहार एवं जीवन मूल्य तथा अभिरुचि मूल्यांकन")
        cs_grades = ["A", "B", "C", "D", "E"]
        cs_desc = ["उत्कृष्ट", "अति उत्तम", "उत्तम", "संतोषजनक", "सुधार अपेक्षित"]
        cgw = (width/2 - 80)/5
        cgy = height - 245
        for i in range(5):
            c.rect(gx + i*cgw, cgy + 25, cgw, 22)
            c.rect(gx + i*cgw, cgy, cgw, 25)
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(gx + i*cgw + cgw/2, cgy + 32, cs_grades[i])
            c.setFont(HINDI_FONT, 8)
            c.drawCentredString(gx + i*cgw + cgw/2, cgy + 9, cs_desc[i])
            
        # General Instructions Box
        c.setFillColor(maroon)
        c.rect(width*0.25 - 60, height - 295, 120, 20, fill=1)
        c.setFillColor(colors.white)
        c.setFont(HINDI_BOLD, 10)
        c.drawCentredString(width*0.25, height - 289, "-: सामान्य निर्देश :-")
        
        c.setFillColor(maroon)
        c.setFont(HINDI_FONT, 8.5)
        instructions = [
            "1. संरक्षक अपने पाल्य/पाल्या के अर्द्धवार्षिक प्राप्तांक देखकर, इस प्रगति पत्र पर",
            "   हस्ताक्षर करके तीन दिन के अंदर विद्यालय में कक्षा अध्यापक के पास जमा कर दें।",
            "2. इस प्रगति पत्र के खो जाने पर दूसरी प्रतिलिपि के लिए ........... रुपये शुल्क देना पड़ेगा।",
            "3. प्रगति पत्र पर अंकित प्राप्तांक के साथ किसी भी तरह की छेड़छाड़ न की जाए।"
        ]
        iy = height - 325
        for line in instructions:
            c.drawString(40, iy, line)
            iy -= 20

        # --- RIGHT HALF: FRONT COVER (STUDENT PROFILE) ---
        rx = width/2 + 10
        rw = width/2 - 30
        c.setStrokeColor(maroon)
        c.setLineWidth(2)
        c.rect(rx, 20, rw, height - 40)
        
        c.setFont(HINDI_BOLD, 8)
        c.drawString(rx + 15, height - 40, "विद्यालय की वेबसाइट: www.adityabirlaschools.co.in/abic-renukoot")
        c.drawRightString(rx + rw - 15, height - 40, "विद्यालय कोड: 1023")
        
        c.setFont(HINDI_BOLD, 16)
        c.drawCentredString(rx + rw/2, height - 95, "आदित्य बिड़ला इण्टरमीडिएट कॉलेज")
        c.setFont(HINDI_BOLD, 12)
        c.drawCentredString(rx + rw/2, height - 115, "रेणुकूट, सोनभद्र (उ०प्र०)")
        
        # Badge
        c.setFillColor(maroon)
        c.roundRect(rx + rw/2 - 65, height - 150, 130, 24, 5, fill=1)
        c.setFillColor(colors.white)
        c.setFont(HINDI_BOLD, 13)
        c.drawCentredString(rx + rw/2, height - 143, "प्रगति-पत्र")
        
        c.setFillColor(maroon)
        c.setFont(HINDI_BOLD, 11)
        c.drawCentredString(rx + rw/2, height - 172, f"शैक्षणिक सत्र {row.get('Session', '2025 - 2026')}")
        
        # Student Details Fields
        fields = [
            (f"छात्र / छात्रा का नाम :  {row['Name']}", ""),
            (f"अनुक्रमांक :  {row['Roll_No']}", f"निलय :  {row['House']}"),
            (f"पिता का नाम :  {row['Father_Name']}", ""),
            (f"माता का नाम :  {row['Mother_Name']}", ""),
            (f"जन्मतिथि :  {row['DOB']}", f"कक्षा एवं वर्ग :  {row['Class_Sec']}"),
            (f"मोबाइल नंबर :  {row['Mobile']}", ""),
            (f"छात्र पंजीकरण संख्या :  {row['Reg_No']}", ""),
            (f"आधार संख्या :  {row['Aadhaar']}", ""),
            (f"पेन संख्या (PEN No.) :  {row['PEN_No']}", ""),
            (f"पता :  {row['Address']}", "")
        ]
        fy = height - 210
        for left_txt, right_txt in fields:
            c.setFillColor(maroon)
            c.setFont(HINDI_BOLD, 10)
            c.drawString(rx + 20, fy, left_txt)
            if right_txt:
                c.drawString(rx + rw/2 + 10, fy, right_txt)
            c.setStrokeColor(maroon)
            c.setLineWidth(0.5)
            c.line(rx + 20, fy - 4, rx + rw - 20, fy - 4)
            fy -= 32
            
        c.showPage()
        
        # ================= PAGE 2: INSIDE MARKSHEET (ACADEMIC + CO-SCHOLASTIC) =================
        c.setFillColor(cream_bg)
        c.rect(0, 0, width, height, fill=1, stroke=0)
        
        # Header
        c.setFillColor(maroon)
        c.setFont(HINDI_BOLD, 15)
        c.drawCentredString(width/2, height - 35, "आदित्य बिड़ला इण्टरमीडिएट कॉलेज, रेणुकूट, सोनभद्र")
        c.setFont(HINDI_BOLD, 11)
        c.drawCentredString(width/2, height - 52, f"प्रगति पत्र विवरण सत्र {row.get('Session', '2025 - 2026')}")
        
        c.setFont(HINDI_BOLD, 9.5)
        c.drawString(30, height - 72, f"छात्र/छात्रा का नाम: {row['Name']}        छात्र पंजिका संख्या: {row['Reg_No']}")
        c.drawString(30, height - 88, f"अनुक्रमांक: {row['Roll_No']}                कक्षा एवं वर्ग: {row['Class_Sec']}")
        
        # --- LEFT SIDE: ACADEMIC EVALUATION TABLE ---
        c.setFont(HINDI_BOLD, 11)
        c.drawCentredString(260, height - 108, "शैक्षिक मूल्यांकन")
        
        tx, ty = 25, height - 120
        col_widths = [22, 85, 22, 22, 22, 24, 28, 28, 30, 28, 28, 30, 28, 28, 32, 30]
        c.setStrokeColor(maroon)
        c.setLineWidth(0.8)
        
        # Convert numeric values safely
        def to_num(val):
            try:
                return float(val) if val != "" else 0
            except:
                return 0

        h_hy, e_hy, mb_hy = to_num(row['Hindi_HY']), to_num(row['English_HY']), to_num(row['Math_Bio_HY'])
        p_hy_th, p_hy_pr = to_num(row['Physics_HY_Th']), to_num(row['Physics_HY_Pr'])
        c_hy_th, c_hy_pr = to_num(row['Chem_HY_Th']), to_num(row['Chem_HY_Pr'])

        h_an, e_an, mb_an = to_num(row['Hindi_AN']), to_num(row['English_AN']), to_num(row['Math_Bio_AN'])
        p_an_th, p_an_pr = to_num(row['Physics_AN_Th']), to_num(row['Physics_AN_Pr'])
        c_an_th, c_an_pr = to_num(row['Chem_AN_Th']), to_num(row['Chem_AN_Pr'])

        subjects = [
            ("1", "हिन्दी", int(h_hy), "", int(h_hy), int(h_an), "", int(h_an)),
            ("2", "अंग्रेजी", int(e_hy), "", int(e_hy), int(e_an), "", int(e_an)),
            ("3", "गणित/जीव विज्ञान", int(mb_hy), "", int(mb_hy), int(mb_an), "", int(mb_an)),
            ("4", "भौतिक विज्ञान", int(p_hy_th), int(p_hy_pr), int(p_hy_th + p_hy_pr),
             int(p_an_th), int(p_an_pr), int(p_an_th + p_an_pr)),
            ("5", "रसायन विज्ञान", int(c_hy_th), int(c_hy_pr), int(c_hy_th + c_hy_pr),
             int(c_an_th), int(c_an_pr), int(c_an_th + c_an_pr)),
            ("6", "नैतिक, खेल एवं शारीरिक शिक्षा", "", "", str(row.get('Moral_Sports_Grade', 'A')), "", "", "")
        ]
        
        total_tw = sum(col_widths)
        c.rect(tx, ty - 55, total_tw, 55)
        c.setFont(HINDI_BOLD, 7.5)
        c.drawString(tx + 35, ty - 30, "विषय")
        c.drawString(tx + 115, ty - 20, "मासिक परीक्षा")
        c.drawString(tx + 205, ty - 20, "अर्द्धवार्षिक परीक्षा")
        c.drawString(tx + 292, ty - 20, "वार्षिक परीक्षा")
        c.drawString(tx + 380, ty - 20, "सम्पपूर्ण परीक्षा")
        
        sub_headers = ["क्र.", "विषय", "प्रथम", "द्वि.", "तृ.", "योग", "सैद्ध.", "प्रायो.", "योग", "सैद्ध.", "प्रायो.", "योग", "सैद्ध.", "प्रायो.", "योग", "ग्रेड"]
        curr_x = tx
        for idx, cw in enumerate(col_widths):
            c.line(curr_x, ty, curr_x, ty - 55)
            if idx >= 2:
                c.setFont(HINDI_FONT, 6.5)
                c.drawCentredString(curr_x + cw/2, ty - 45, sub_headers[idx])
            curr_x += cw
        c.line(tx + 107, ty - 30, tx + total_tw - 30, ty - 30)
        
        ry = ty - 55
        row_h = 26
        hy_grand = 0
        an_grand = 0
        
        for sub in subjects:
            c.rect(tx, ry - row_h, total_tw, row_h)
            cx = tx
            
            hy_tot = sub[4] if isinstance(sub[4], (int, float)) else 0
            an_tot = sub[7] if isinstance(sub[7], (int, float)) else 0
            hy_grand += hy_tot
            an_grand += an_tot
            
            grand_sub = hy_tot + an_tot if an_tot > 0 else hy_tot
            max_sub = 200 if an_tot > 0 else 100
            grade = get_grade(grand_sub, max_sub) if isinstance(sub[4], (int, float)) else sub[4]
            
            vals = [
                sub[0], sub[1], "", "", "", "",
                str(sub[2]) if sub[2] != "" else "",
                str(sub[3]) if sub[3] != "" else "",
                str(sub[4]) if sub[4] != "" else "",
                str(sub[5]) if sub[5] not in ["", 0] else "",
                str(sub[6]) if sub[6] not in ["", 0] else "",
                str(sub[7]) if sub[7] not in ["", 0] else "",
                "", "", str(grand_sub) if (an_tot > 0 and isinstance(sub[4], (int, float))) else "",
                grade
            ]
            for idx, cw in enumerate(col_widths):
                c.line(cx, ry, cx, ry - row_h)
                c.setFont(HINDI_BOLD if idx in [1, 8, 11, 15] else "Helvetica", 7.5 if idx==1 else 8.5)
                if idx == 1:
                    c.drawString(cx + 3, ry - 16, vals[idx])
                else:
                    c.drawCentredString(cx + cw/2, ry - 16, vals[idx])
                cx += cw
            ry -= row_h
            
        # Total Row
        c.rect(tx, ry - 24, total_tw, 24)
        c.setFont(HINDI_BOLD, 8.5)
        c.drawString(tx + 35, ry - 16, "योग (Total)")
        c.drawCentredString(tx + sum(col_widths[:8]) + col_widths[8]/2, ry - 16, f"{int(hy_grand)}/500")
        if an_grand > 0:
            c.drawCentredString(tx + sum(col_widths[:11]) + col_widths[11]/2, ry - 16, f"{int(an_grand)}/500")
            c.drawCentredString(tx + sum(col_widths[:14]) + col_widths[14]/2, ry - 16, f"{int(hy_grand+an_grand)}/1000")
        
        # Footer Signatures Box
        ry -= 24
        c.rect(tx, ry - 115, total_tw, 115)
        c.setFont(HINDI_BOLD, 8)
        c.drawString(tx + 8, ry - 16, f"उपस्थिति : {row.get('Attendance', '')}")
        c.drawString(tx + 8, ry - 34, f"परीक्षाफल : {row.get('Result', '')}")
        c.drawString(tx + 8, ry - 55, "ह० कक्षाध्यापक")
        c.drawString(tx + 8, ry - 73, "ह० निरीक्षक")
        c.drawString(tx + 8, ry - 91, "ह० प्रधानाचार्य")
        c.drawString(tx + 8, ry - 109, "ह० अभिभावक")
        
        c.drawCentredString(tx + total_tw/2 + 20, ry - 16, "विशेष विवरण")
        c.drawRightString(tx + total_tw - 15, ry - 55, "ह० कक्षाध्यापक")
        c.drawRightString(tx + total_tw - 15, ry - 73, "ह० निरीक्षक")
        c.drawRightString(tx + total_tw - 15, ry - 91, "ह० प्रधानाचार्य")
        c.drawRightString(tx + total_tw - 15, ry - 109, "ह० अभिभावक")
        
        # --- RIGHT SIDE: CO-SCHOLASTIC EVALUATION ---
        csx = tx + total_tw + 18
        csw = width - csx - 25
        csy = height - 85
        
        c.setFillColor(maroon)
        c.roundRect(csx + csw/2 - 75, csy, 150, 20, 4, fill=1)
        c.setFillColor(colors.white)
        c.setFont(HINDI_BOLD, 10)
        c.drawCentredString(csx + csw/2, csy + 6, "सह शैक्षणिक मूल्यांकन")
        
        c.setFillColor(maroon)
        csy -= 28
        c.rect(csx, csy, csw, 22)
        c.setFont(HINDI_BOLD, 8.5)
        c.drawString(csx + 5, csy + 7, "क्रम")
        c.drawCentredString(csx + csw/2, csy + 7, "मूल्यांकन का क्षेत्र")
        c.drawRightString(csx + csw - 10, csy + 7, "ग्रेड")
        
        co_items = [
            ("HEADER", "जीवन कौशल"),
            ("1", "सृजनात्मकता"),
            ("2", "प्रभावी सम्प्रेषण"),
            ("3", "निर्णय लेने की क्षमता एवं नेतृत्व क्षमता"),
            ("4", "जिज्ञासा, तार्किक चिंतन एवं समस्या का समाधान"),
            ("5", "तनाव प्रबन्धन"),
            ("HEADER", "व्यवहार एवं जीवन मूल्य"),
            ("6", "शिक्षकों से व्यवहार"),
            ("7", "सहपाठियों से व्यवहार"),
            ("8", "शिष्टाचार, अनुशासन, समयबद्धता"),
            ("9", "व्यक्तिगत एवं विद्यालयी स्वच्छता"),
            ("10", "पर्यावरण के प्रति जागरूकता"),
            ("11", "सहयोग, सामूहिकता, सहानुभूति"),
            ("HEADER", "अभिरुचि"),
            ("12", "खेल एवं योग"),
            ("13", "पाठ्य सहगामी गतिविधियाँ"),
            ("14", "कला / संगीत / साहित्यिक गतिविधियाँ"),
            ("15", "वैज्ञानिक अभिरुचि")
        ]
        
        item_h = 21.5
        for code, label in co_items:
            csy -= item_h
            c.rect(csx, csy, csw, item_h)
            if code == "HEADER":
                c.setFont(HINDI_BOLD, 8.5)
                c.drawCentredString(csx + csw/2, csy + 6, label)
            else:
                c.line(csx + 28, csy, csx + 28, csy + item_h)
                c.line(csx + csw - 38, csy, csx + csw - 38, csy + item_h)
                c.setFont("Helvetica", 8)
                c.drawCentredString(csx + 14, csy + 6, code)
                c.setFont(HINDI_FONT, 8)
                c.drawString(csx + 34, csy + 6, label)
                c.setFont("Helvetica-Bold", 8.5)
                grade_val = str(row.get(f"CS_{code}", "A"))
                c.drawCentredString(csx + csw - 19, csy + 6, grade_val)
                
        c.showPage()
        
    c.save()
    return buffer.getvalue()

# ================= STREAMLIT UI =================
st.title("🎓 ABIC Renukoot - Pragati Patra (Marksheet) Generator")
st.markdown("Aditya Birla Intermediate College ke official format mein **Half-Yearly / Annual Marksheet PDF** banayein.")

tab1, tab2 = st.tabs(["📂 Bulk Excel Upload (Poori Class ke liye)", "✍️ Single Student Entry"])

with tab1:
    st.subheader("Step 1: Sample Excel Template Download Karein")
    st.download_button(
        label="📥 Download Sample Excel Template",
        data=get_sample_excel(),
        file_name="ABIC_Class_Marksheet_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.subheader("Step 2: Bhari hui Excel File Upload Karein")
    uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df_uploaded = pd.read_excel(uploaded_file).fillna("")
        st.dataframe(df_uploaded.head())
        if st.button("🚀 Generate All Marksheets PDF"):
            pdf_bytes = generate_marksheet_pdf(df_uploaded)
            st.success(f"✅ {len(df_uploaded)} students ki marksheet taiyar hai!")
            st.download_button(
                label="📄 Download Combined Print-Ready PDF",
                data=pdf_bytes,
                file_name="ABIC_Marksheets_Print.pdf",
                mime="application/pdf"
            )

with tab2:
    st.subheader("Single Student Quick Entry")
    with st.form("single_student_form"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("Student Name", "Anandi Kumari Singh")
        roll = c2.text_input("Roll No", "44")
        reg = c3.text_input("Reg No", "38940")
        
        c4, c5, c6 = st.columns(3)
        fname = c4.text_input("Father's Name", "Shriram Singh")
        mname = c5.text_input("Mother's Name", "Lalita Devi")
        dob = c6.text_input("DOB", "26-11-2009")
        
        c7, c8, c9 = st.columns(3)
        cls_sec = c7.text_input("Class & Sec", "11 'B'")
        house = c8.text_input("House (Nilay)", "Arun")
        mob = c9.text_input("Mobile", "9450320524")
        
        c10, c11, c12 = st.columns(3)
        aadhaar = c10.text_input("Aadhaar No", "982931653622")
        pen = c11.text_input("PEN No", "21473874271")
        addr = c12.text_input("Address", "Jhopadi Near Bus Stand, Renukoot")
        
        st.markdown("#### Half-Yearly Marks")
        m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
        h_hy = m1.number_input("Hindi (100)", 0, 100, 80)
        e_hy = m2.number_input("English (100)", 0, 100, 65)
        mb_hy = m3.number_input("Math/Bio (100)", 0, 100, 76)
        p_th = m4.number_input("Phy Th (70)", 0, 70, 56)
        p_pr = m5.number_input("Phy Pr (30)", 0, 30, 29)
        c_th = m6.number_input("Chem Th (70)", 0, 70, 38)
        c_pr = m7.number_input("Chem Pr (30)", 0, 30, 28)
        
        submitted = st.form_submit_button("Generate Marksheet PDF")

    if submitted:
        single_df = pd.DataFrame([{
            "Roll_No": roll, "Reg_No": reg, "Name": name,
            "Father_Name": fname, "Mother_Name": mname, "DOB": dob,
            "Class_Sec": cls_sec, "House": house, "Mobile": mob,
            "Aadhaar": aadhaar, "PEN_No": pen, "Address": addr,
            "Hindi_HY": h_hy, "English_HY": e_hy, "Math_Bio_HY": mb_hy,
            "Physics_HY_Th": p_th, "Physics_HY_Pr": p_pr,
            "Chem_HY_Th": c_th, "Chem_HY_Pr": c_pr,
            "Moral_Sports_Grade": "A",
            "Hindi_AN": 0, "English_AN": 0, "Math_Bio_AN": 0,
            "Physics_AN_Th": 0, "Physics_AN_Pr": 0,
            "Chem_AN_Th": 0, "Chem_AN_Pr": 0,
            **{f"CS_{i}": "A" for i in range(1, 16)},
            "Attendance": "", "Result": "", "Session": "2025 - 2026"
        }])
        pdf_data = generate_marksheet_pdf(single_df)
        st.success("✅ Marksheet PDF taiyar hai! Neeche button se download karein.")
        st.download_button(
            label="📄 Download PDF",
            data=pdf_data,
            file_name=f"{name}_Marksheet.pdf",
            mime="application/pdf"
        )
