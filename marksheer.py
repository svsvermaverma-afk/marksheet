import streamlit as st
import pandas as pd
import base64
import os
import streamlit.components.v1 as components

st.set_page_config(page_title="पूर्ण प्रगति-पत्र - ABIC रेणुकूट", layout="wide")

# GitHub repository mein rakhi logo.png ko Base64 mein convert karne ka function
def get_logo_base64(image_path="logo.png"):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode("utf-8")
    return ""

logo_b64 = get_logo_base64("logo.png")
logo_html = (
    f'<img src="data:image/png;base64,{logo_b64}" class="school-logo" alt="ABIC Logo">'
    if logo_b64
    else '<div style="font-weight:bold; color:#7c2d12; margin-bottom:5px;">[logo.png अपलोड करें]</div>'
)

st.title("📄 पूर्ण प्रगति-पत्र (2-Page Full Marksheet) जनरेटर")

# Sidebar Settings
with st.sidebar:
    st.header("1. डेटा अपलोड करें")
    uploaded_file = st.file_uploader("Excel (.xlsx) या CSV फाइल चुनें", type=["xlsx", "xls", "csv"])
    session_year = st.text_input("शैक्षणिक सत्र", value="2026 - 2027")
    stream_type = st.selectbox("वर्ग / संकाय चुनें", ["विज्ञान वर्ग (Science - 11/12)", "वाणिज्य वर्ग (Commerce - 11/12)", "कला वर्ग (Arts - 11/12)", "हाईस्कूल / जूनियर (6-10)"])

# Default Sample Data
student_data = {
    "name": "ABHIMANYU YADAV",
    "roll": "1",
    "house": "MARTAND",
    "father": "SRILAL YADAV",
    "mother": "SONA DEVI",
    "dob": "2008-11-19",
    "class_sec": "12 B",
    "mobile": "9889853760",
    "reg_no": "38940",
    "aadhaar": "982931653622",
    "pen": "21473874271",
    "address": "Renukoot, Sonbhadra (U.P.)"
}

# Default Subjects based on Stream
if "विज्ञान" in stream_type:
    default_subjects = ["सामान्य हिन्दी", "अंग्रेजी (English)", "भौतिक विज्ञान (Physics)", "रसायन विज्ञान (Chemistry)", "गणित / जीव विज्ञान", "खेल एवं शारीरिक शिक्षा"]
elif "वाणिज्य" in stream_type:
    default_subjects = ["सामान्य हिन्दी", "अंग्रेजी (English)", "लेखाशास्त्र (Accountancy)", "व्यवसाय अध्ययन (Business Studies)", "अर्थशास्त्र (Economics)", "खेल एवं शारीरिक शिक्षा"]
elif "कला" in stream_type:
    default_subjects = ["सामान्य हिन्दी", "अंग्रेजी (English)", "इतिहास (History)", "भूगोल (Geography)", "नागरिक शास्त्र (Civics)", "खेल एवं शारीरिक शिक्षा"]
else:
    default_subjects = ["हिन्दी", "अंग्रेजी (English)", "गणित (Mathematics)", "विज्ञान (Science)", "सामाजिक विज्ञान", "कला / कम्प्यूटर"]

marks_rows = []
for sub in default_subjects:
    marks_rows.append({
        "subject": sub,
        "q_max": "20", "q_obt": "",
        "h_max": "100", "h_obt": "",
        "a_max": "100", "a_obt": "",
        "total_max": "220", "total_obt": "",
        "grade": ""
    })

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.str.strip()
        name_col = next((c for c in df.columns if "name" in c.lower() or "नाम" in c), df.columns[0])
        selected_student = st.sidebar.selectbox("विद्यार्थी चुनें:", df[name_col].astype(str).tolist())
        row = df[df[name_col].astype(str) == selected_student].iloc[0]

        def get_val(keywords, default=""):
            for col in df.columns:
                if any(k in col.lower() for k in keywords):
                    val = str(row[col])
                    if val.lower() == "nan":
                        return ""
                    return val.replace(" 00:00:00", "").strip()
            return default

        student_data["name"] = get_val(["student", "name", "छात्र"], "")
        student_data["roll"] = get_val(["roll", "अनुक्रमांक"], "")
        student_data["house"] = get_val(["house", "nilay", "निलय"], "")
        student_data["father"] = get_val(["father", "पिता"], "")
        student_data["mother"] = get_val(["mother", "माता"], "")
        student_data["dob"] = get_val(["dob", "birth", "जन्म"], "")
        student_data["class_sec"] = get_val(["class", "कक्षा"], "")
        student_data["mobile"] = get_val(["mobile", "phone", "मोबाइल"], "")
        student_data["reg_no"] = get_val(["reg", "पंजीकरण"], "")
        student_data["aadhaar"] = get_val(["aadhar", "aadhaar", "आधार"], "")
        student_data["pen"] = get_val(["pen", "पेन"], "")
        student_data["address"] = get_val(["address", "पता"], "")

        # Agar Excel mein subject columns hain toh unke marks bhi fetch karein
        for item in marks_rows:
            sub_key = item["subject"].split("(")[0].strip().lower()
            for col in df.columns:
                if sub_key in col.lower() or any(w in col.lower() for w in item["subject"].lower().replace("(", "").replace(")", "").split() if len(w) > 3):
                    val = str(row[col]).replace(" 00:00:00", "").strip()
                    if val.lower() != "nan":
                        item["h_obt"] = val
                        break

    except Exception as e:
        st.error(f"फाइल पढ़ने में त्रुटि: {e}")

# Build Marks Table HTML rows
marks_table_rows_html = ""
for idx, m in enumerate(marks_rows, 1):
    marks_table_rows_html += f"""
    <tr>
      <td>{idx}</td>
      <td class="sub-name" contenteditable="true">{m['subject']}</td>
      <td contenteditable="true">{m['q_max']}</td>
      <td class="blue-text" contenteditable="true">{m['q_obt']}</td>
      <td contenteditable="true">{m['h_max']}</td>
      <td class="blue-text" contenteditable="true">{m['h_obt']}</td>
      <td contenteditable="true">{m['a_max']}</td>
      <td class="blue-text" contenteditable="true">{m['a_obt']}</td>
      <td contenteditable="true">{m['total_max']}</td>
      <td class="blue-text" contenteditable="true">{m['total_obt']}</td>
      <td class="blue-text" contenteditable="true">{m['grade']}</td>
    </tr>
    """

# Complete 2-Page A4 Landscape HTML + CSS
html_template = f"""
<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <style>
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: 'Mangal', 'Nirmala UI', 'Noto Sans Devanagari', Arial, sans-serif;
      background-color: transparent;
      display: flex;
      flex-direction: column;
      align-items: center;
      color: #2b1d14;
      gap: 20px;
    }}

    .print-bar {{
      width: 100%;
      text-align: center;
      padding: 8px;
      background: #fff8eb;
      border: 1px solid #d6c3a5;
      border-radius: 8px;
    }}

    .btn-print {{
      background-color: #7c2d12;
      color: #ffffff;
      border: none;
      padding: 10px 28px;
      font-size: 16px;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }}

    .btn-print:hover {{
      background-color: #9a3412;
    }}

    .tip-text {{
      font-size: 12.5px;
      color: #5c3a21;
      margin-top: 5px;
    }}

    /* A4 Landscape Sheet */
    .page-sheet {{
      width: 297mm;
      height: 210mm;
      background-color: #fdf6e2;
      padding: 6mm;
      border: 1px solid #bba88a;
      display: flex;
      justify-content: center;
      align-items: center;
      page-break-after: always;
    }}

    .page-sheet:last-child {{
      page-break-after: auto;
    }}

    .outer-border {{
      width: 100%;
      height: 100%;
      border: 2.5px solid #5c3a21;
      padding: 3px;
      display: flex;
    }}

    .inner-border {{
      width: 100%;
      height: 100%;
      border: 1px solid #5c3a21;
      display: flex;
      flex-direction: column;
    }}

    /* PAGE 1: Two-Column Layout */
    .two-col-grid {{
      width: 100%;
      height: 100%;
      display: grid;
      grid-template-columns: 1fr 1fr;
    }}

    .column {{
      padding: 9mm 8.5mm;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }}

    .left-col {{
      border-right: 1.5px dashed #5c3a21;
    }}

    .section-title {{
      font-size: 13.5px;
      font-weight: bold;
      color: #5c2818;
      margin-bottom: 6px;
    }}

    table.grade-table {{
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 16px;
      text-align: center;
      font-size: 12px;
    }}

    table.grade-table th,
    table.grade-table td {{
      border: 1px solid #5c3a21;
      padding: 6px 3px;
    }}

    table.grade-table th {{
      font-weight: normal;
      background-color: rgba(92, 58, 33, 0.06);
    }}

    table.grade-table td {{
      font-weight: bold;
    }}

    .instructions-header {{
      text-align: center;
      margin-bottom: 10px;
    }}

    .badge-maroon {{
      display: inline-block;
      background-color: #8c3b2b;
      color: #ffffff;
      font-weight: bold;
      font-size: 14px;
      padding: 3px 18px;
      border-radius: 3px;
    }}

    .instructions-list {{
      list-style: none;
      font-size: 12.8px;
      line-height: 1.75;
    }}

    .instructions-list li {{
      margin-bottom: 8px;
      text-align: justify;
    }}

    .top-meta {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      font-size: 10.5px;
      text-align: center;
    }}

    .meta-box {{
      border: 1px solid #5c3a21;
      padding: 2px 7px;
      margin-top: 2px;
      font-weight: bold;
      background-color: #fffdf7;
    }}

    .school-header {{
      text-align: center;
      margin-top: 2px;
    }}

    .school-logo {{
      width: 66px;
      height: auto;
      display: block;
      margin: 0 auto 4px auto;
    }}

    .school-name {{
      font-size: 20px;
      font-weight: bold;
      color: #5c2818;
    }}

    .school-address {{
      font-size: 13.5px;
      font-weight: bold;
      color: #5c2818;
    }}

    .report-badge {{
      display: inline-block;
      background-color: #8c3b2b;
      color: #ffffff;
      font-size: 16px;
      font-weight: bold;
      padding: 3px 22px;
      border-radius: 4px;
      margin: 5px 0 3px 0;
    }}

    .session-text {{
      font-size: 14.5px;
      font-weight: bold;
      color: #3b2314;
    }}

    .student-details {{
      margin-top: 6px;
      font-size: 13px;
      line-height: 1.95;
    }}

    .field-row {{
      display: flex;
      align-items: baseline;
      width: 100%;
    }}

    .field-label {{
      white-space: nowrap;
      font-weight: bold;
      color: #4a2c18;
    }}

    .field-value {{
      flex: 1;
      border-bottom: 1.5px dotted #5c3a21;
      padding-left: 8px;
      font-weight: bold;
      color: #1e3a8a;
      min-height: 19px;
    }}

    .field-half {{
      display: flex;
      align-items: baseline;
      width: 50%;
    }}

    /* PAGE 2: Inside Marks Sheet Styling */
    .inside-page {{
      padding: 7mm 9mm;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      height: 100%;
    }}

    .inside-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 1.5px solid #5c3a21;
      padding-bottom: 5px;
    }}

    .inside-title-box {{
      text-align: center;
    }}

    .inside-school {{
      font-size: 18px;
      font-weight: bold;
      color: #5c2818;
    }}

    .inside-sub {{
      font-size: 13px;
      font-weight: bold;
      color: #8c3b2b;
    }}

    .mini-student-bar {{
      display: flex;
      justify-content: space-between;
      background-color: rgba(92, 58, 33, 0.07);
      border: 1px solid #5c3a21;
      padding: 5px 10px;
      font-size: 12.5px;
      font-weight: bold;
      margin: 6px 0;
    }}

    .blue-text {{
      color: #1e3a8a;
      font-weight: bold;
    }}

    table.marks-table {{
      width: 100%;
      border-collapse: collapse;
      text-align: center;
      font-size: 12.5px;
    }}

    table.marks-table th,
    table.marks-table td {{
      border: 1px solid #5c3a21;
      padding: 5px 4px;
    }}

    table.marks-table th {{
      background-color: rgba(140, 59, 43, 0.12);
      color: #4a2c18;
      font-weight: bold;
    }}

    table.marks-table td.sub-name {{
      text-align: left;
      padding-left: 8px;
      font-weight: bold;
    }}

    .total-row {{
      background-color: rgba(92, 58, 33, 0.08);
      font-weight: bold;
    }}

    .bottom-grid {{
      display: grid;
      grid-template-columns: 1.2fr 1fr;
      gap: 12px;
      margin-top: 6px;
    }}

    table.small-info-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      text-align: center;
    }}

    table.small-info-table th,
    table.small-info-table td {{
      border: 1px solid #5c3a21;
      padding: 4px 6px;
    }}

    table.small-info-table th {{
      background-color: rgba(92, 58, 33, 0.08);
    }}

    .result-box {{
      border: 1px solid #5c3a21;
      padding: 6px 10px;
      font-size: 12.5px;
      display: flex;
      flex-direction: column;
      justify-content: space-around;
      background-color: #fffdf7;
    }}

    .sign-row {{
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      margin-top: 24px;
      padding: 0 15px;
      font-size: 13px;
      font-weight: bold;
      color: #4a2c18;
    }}

    .sign-box {{
      text-align: center;
      border-top: 1px solid #5c3a21;
      padding-top: 4px;
      min-width: 160px;
    }}

    /* Exact Full A4 Landscape Print Rules */
    @media print {{
      @page {{
        size: A4 landscape;
        margin: 0;
      }}

      html, body {{
        width: 297mm;
        margin: 0 !important;
        padding: 0 !important;
        background: #fdf6e2 !important;
        gap: 0 !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }}

      .print-bar {{
        display: none !important;
      }}

      .page-sheet {{
        width: 297mm !important;
        height: 210mm !important;
        border: none !important;
        padding: 5mm !important;
        margin: 0 !important;
        page-break-inside: avoid;
      }}
    }}
  </style>
</head>
<body>

  <div class="print-bar">
    <button class="btn-print" onclick="window.print()">🖨️ पूरी मार्कशीट प्रिंट करें (2 Pages - Full A4 Landscape)</button>
    <div class="tip-text">💡 टिप: आप मार्कशीट के किसी भी नंबर या नाम पर क्लिक करके सीधे एडिट कर सकते हैं। प्रिंट करते समय <b>Two-sided (Flip on short edge)</b> चुनें।</div>
  </div>

  <!-- ================= PAGE 1: FRONT & BACK COVER ================= -->
  <div class="page-sheet">
    <div class="outer-border">
      <div class="inner-border">
        <div class="two-col-grid">

          <!-- LEFT COLUMN -->
          <div class="column left-col">
            <div>
              <div class="section-title">शैक्षिक मूल्यांकन ग्रेडिंग सिस्टम (प्रतिशत में)</div>
              <table class="grade-table">
                <tr>
                  <th>91-100</th><th>81-90</th><th>71-80</th><th>61-70</th>
                  <th>51-60</th><th>41-50</th><th>33-40</th><th>0-32</th>
                </tr>
                <tr>
                  <td>A1</td><td>A2</td><td>B1</td><td>B2</td>
                  <td>C1</td><td>C2</td><td>D</td><td>E</td>
                </tr>
              </table>

              <div class="section-title">जीवन कौशल व्यवहार एवं जीवन मूल्य तथा अभिरुचि मूल्यांकन<br>(ग्रेडिंग सिस्टम)</div>
              <table class="grade-table">
                <tr>
                  <th>A</th><th>B</th><th>C</th><th>D</th><th>E</th>
                </tr>
                <tr>
                  <td>उत्कृष्ट</td><td>अति<br>उत्तम</td><td>उत्तम</td><td>संतोषजनक</td><td>सुधार<br>अपेक्षित</td>
                </tr>
              </table>
            </div>

            <div>
              <div class="instructions-header">
                <span class="badge-maroon">-: सामान्य निर्देश :-</span>
              </div>
              <ul class="instructions-list">
                <li>1. संरक्षक अपने पाल्य/पाल्या के अर्द्धवार्षिक प्राप्तांक देखकर, इस प्रगति पत्र पर हस्ताक्षर करके तीन दिन के अंदर विद्यालय में कक्षा अध्यापक के पास जमा कर दें।</li>
                <li>2. इस प्रगति पत्र के खो जाने पर दूसरी प्रतिलिपि के लिए ...................... रुपये का शुल्क देना पड़ेगा।</li>
                <li>3. प्रगति पत्र पर अंकित प्राप्तांक के साथ किसी भी तरह की छेड़छाड़ न की जाए।</li>
              </ul>
            </div>
          </div>

          <!-- RIGHT COLUMN -->
          <div class="column right-col">
            <div class="top-meta">
              <div>
                <div>विद्यालय की वेबसाइट</div>
                <div class="meta-box">www.adityabirlaschools.co.in/abic-renukoot</div>
              </div>
              <div>
                <div>विद्यालय का कोड</div>
                <div class="meta-box">1023</div>
              </div>
            </div>

            <div class="school-header">
              {logo_html}
              <div class="school-name">आदित्य बिड़ला इण्टरमीडिएट कॉलेज</div>
              <div class="school-address">रेणुकूट, सोनभद्र (उ०प्र०)</div>
              <div class="report-badge">प्रगति-पत्र</div>
              <div class="session-text">शैक्षणिक सत्र {session_year}</div>
            </div>

            <div class="student-details">
              <div class="field-row">
                <span class="field-label">छात्र / छात्रा का नाम :</span>
                <span class="field-value" contenteditable="true">{student_data['name']}</span>
              </div>

              <div class="field-row">
                <div class="field-half">
                  <span class="field-label">अनुक्रमांक :</span>
                  <span class="field-value" contenteditable="true">{student_data['roll']}</span>
                </div>
                <div class="field-half" style="padding-left: 10px;">
                  <span class="field-label">निलय :</span>
                  <span class="field-value" contenteditable="true">{student_data['house']}</span>
                </div>
              </div>

              <div class="field-row">
                <span class="field-label">पिता का नाम :</span>
                <span class="field-value" contenteditable="true">{student_data['father']}</span>
              </div>

              <div class="field-row">
                <span class="field-label">माता का नाम :</span>
                <span class="field-value" contenteditable="true">{student_data['mother']}</span>
              </div>

              <div class="field-row">
                <div class="field-half">
                  <span class="field-label">जन्मतिथि :</span>
                  <span class="field-value" contenteditable="true">{student_data['dob']}</span>
                </div>
                <div class="field-half" style="padding-left: 10px;">
                  <span class="field-label">कक्षा एवं वर्ग :</span>
                  <span class="field-value" contenteditable="true">{student_data['class_sec']}</span>
                </div>
              </div>

              <div class="field-row">
                <span class="field-label">मोबाइल नंबर :</span>
                <span class="field-value" contenteditable="true">{student_data['mobile']}</span>
              </div>

              <div class="field-row">
                <span class="field-label">छात्र पंजीकरण संख्या :</span>
                <span class="field-value" contenteditable="true">{student_data['reg_no']}</span>
              </div>

              <div class="field-row">
                <span class="field-label">आधार संख्या :</span>
                <span class="field-value" contenteditable="true">{student_data['aadhaar']}</span>
              </div>

              <div class="field-row">
                <span class="field-label">पेन संख्या (PEN No.) :</span>
                <span class="field-value" contenteditable="true">{student_data['pen']}</span>
              </div>

              <div class="field-row">
                <span class="field-label">पता :</span>
                <span class="field-value" contenteditable="true">{student_data['address']}</span>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>
  </div>

  <!-- ================= PAGE 2: INSIDE MARKS SHEET ================= -->
  <div class="page-sheet">
    <div class="outer-border">
      <div class="inner-border">
        <div class="inside-page">

          <!-- Header -->
          <div class="inside-header">
            <div style="font-size: 12px; font-weight: bold;">विद्यालय कोड : 1023</div>
            <div class="inside-title-box">
              <div class="inside-school">आदित्य बिड़ला इण्टरमीडिएट कॉलेज, रेणुकूट, सोनभद्र (उ०प्र०)</div>
              <div class="inside-sub">शैक्षिक मूल्यांकन विवरण-पत्र (सत्र {session_year})</div>
            </div>
            <div style="font-size: 12px; font-weight: bold;">पंजीकरण सं० : <span class="blue-text">{student_data['reg_no']}</span></div>
          </div>

          <!-- Student Info Summary Bar -->
          <div class="mini-student-bar">
            <span>छात्र/छात्रा : <span class="blue-text">{student_data['name']}</span></span>
            <span>कक्षा एवं वर्ग : <span class="blue-text">{student_data['class_sec']}</span></span>
            <span>अनुक्रमांक : <span class="blue-text">{student_data['roll']}</span></span>
            <span>निलय : <span class="blue-text">{student_data['house']}</span></span>
          </div>

          <!-- Subject-wise Marks Table -->
          <table class="marks-table">
            <thead>
              <tr>
                <th rowspan="2" style="width: 40px;">क्र०सं०</th>
                <th rowspan="2" style="width: 240px;">विषय (Subjects)</th>
                <th colspan="2">त्रैमासिक / मासिक परीक्षा</th>
                <th colspan="2">अर्द्धवार्षिक परीक्षा</th>
                <th colspan="2">वार्षिक परीक्षा</th>
                <th colspan="2">महायोग (Grand Total)</th>
                <th rowspan="2" style="width: 65px;">ग्रेड</th>
              </tr>
              <tr>
                <th>पूर्णांक</th>
                <th>प्राप्तांक</th>
                <th>पूर्णांक</th>
                <th>प्राप्तांक</th>
                <th>पूर्णांक</th>
                <th>प्राप्तांक</th>
                <th>पूर्णांक</th>
                <th>प्राप्तांक</th>
              </tr>
            </thead>
            <tbody>
              {marks_table_rows_html}
              <tr class="total-row">
                <td colspan="2">महायोग (Grand Total)</td>
                <td contenteditable="true">100</td>
                <td class="blue-text" contenteditable="true"></td>
                <td contenteditable="true">500</td>
                <td class="blue-text" contenteditable="true"></td>
                <td contenteditable="true">500</td>
                <td class="blue-text" contenteditable="true"></td>
                <td contenteditable="true">1100</td>
                <td class="blue-text" contenteditable="true"></td>
                <td class="blue-text" contenteditable="true"></td>
              </tr>
            </tbody>
          </table>

          <!-- Attendance, Co-scholastic & Result Summary -->
          <div class="bottom-grid">
            <table class="small-info-table">
              <thead>
                <tr>
                  <th>मूल्यांकन क्षेत्र</th>
                  <th>अर्द्धवार्षिक</th>
                  <th>वार्षिक</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="text-align:left; font-weight:bold;">जीवन कौशल व्यवहार एवं मूल्य (ग्रेड)</td>
                  <td class="blue-text" contenteditable="true">A</td>
                  <td class="blue-text" contenteditable="true">A</td>
                </tr>
                <tr>
                  <td style="text-align:left; font-weight:bold;">अभिरुचि मूल्यांकन (ग्रेड)</td>
                  <td class="blue-text" contenteditable="true">A</td>
                  <td class="blue-text" contenteditable="true">A</td>
                </tr>
                <tr>
                  <td style="text-align:left; font-weight:bold;">उपस्थिति (कुल कार्यदिवस / उपस्थित)</td>
                  <td class="blue-text" contenteditable="true">98 / 110</td>
                  <td class="blue-text" contenteditable="true">195 / 215</td>
                </tr>
              </tbody>
            </table>

            <div class="result-box">
              <div><b>कुल प्राप्तांक प्रतिशत (%) :</b> <span class="blue-text" contenteditable="true">________ %</span> &nbsp;&nbsp; <b>कक्षा में स्थान :</b> <span class="blue-text" contenteditable="true">______</span></div>
              <div><b>परीक्षाफल (Result) :</b> <span class="blue-text" contenteditable="true">उत्तीर्ण (Passed) / प्रोन्नत (Promoted)</span></div>
              <div><b>विशेष टिप्पणी (Remarks) :</b> <span class="blue-text" contenteditable="true">उत्तम प्रदर्शन (Good Performance)</span></div>
            </div>
          </div>

          <!-- Signatures Footer -->
          <div class="sign-row">
            <div class="sign-box">कक्षा अध्यापक के हस्ताक्षर</div>
            <div class="sign-box">अभिभावक के हस्ताक्षर</div>
            <div class="sign-box">प्रधानाचार्य के हस्ताक्षर एवं मुहर</div>
          </div>

        </div>
      </div>
    </div>
  </div>

</body>
</html>
"""

# Render full 2-page marksheet inside Streamlit
components.html(html_template, height=1680, scrolling=True)
