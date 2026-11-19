import streamlit as st
import pandas as pd
import base64
import os
import streamlit.components.v1 as components

st.set_page_config(page_title="प्रगति-पत्र जनरेटर - ABIC रेणुकूट", layout="wide")

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

st.title("📄 प्रगति-पत्र (Marksheet) प्रिंट सिस्टम")

# Sidebar mein Excel/CSV upload ya Manual Entry ka option
with st.sidebar:
    st.header("1. डेटा अपलोड करें")
    uploaded_file = st.file_uploader("Excel (.xlsx) या CSV फाइल चुनें", type=["xlsx", "xls", "csv"])
    session_year = st.text_input("शैक्षणिक सत्र", value="2026 - 2027")

# Default sample data (jab tak Excel upload na ho)
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

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.str.strip()
        
        # Pehle column (ya Name column) se student chunein
        name_col = next((c for c in df.columns if "name" in c.lower() or "नाम" in c), df.columns[0])
        selected_student = st.sidebar.selectbox("विद्यार्थी चुनें:", df[name_col].astype(str).tolist())
        row = df[df[name_col].astype(str) == selected_student].iloc[0]

        def get_val(keywords, default=""):
            for col in df.columns:
                if any(k in col.lower() for k in keywords):
                    val = str(row[col])
                    if val.lower() == "nan":
                        return ""
                    # Date mein se 00:00:00 hatane ke liye
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

    except Exception as e:
        st.error(f"फाइल पढ़ने में त्रुटि: {e}")

# HTML + CSS Template (Exact Two-Column Full A4 Landscape)
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
    }}

    .print-bar {{
      width: 100%;
      text-align: center;
      margin-bottom: 12px;
    }}

    .btn-print {{
      background-color: #7c2d12;
      color: #ffffff;
      border: none;
      padding: 10px 26px;
      font-size: 16px;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }}

    .btn-print:hover {{
      background-color: #9a3412;
    }}

    /* A4 Landscape Exact Dimensions */
    .page-sheet {{
      width: 297mm;
      height: 210mm;
      background-color: #fdf6e2;
      padding: 7mm;
      border: 1px solid #ccc;
      display: flex;
      justify-content: center;
      align-items: center;
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
      display: grid;
      grid-template-columns: 1fr 1fr;
    }}

    .column {{
      padding: 10mm 9mm;
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
      margin-bottom: 18px;
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
      margin-bottom: 12px;
    }}

    .instructions-badge {{
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
      width: 68px;
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
      margin-top: 1px;
    }}

    .report-badge {{
      display: inline-block;
      background-color: #8c3b2b;
      color: #ffffff;
      font-size: 16px;
      font-weight: bold;
      padding: 3px 22px;
      border-radius: 4px;
      margin: 6px 0 4px 0;
    }}

    .session-text {{
      font-size: 14.5px;
      font-weight: bold;
      color: #3b2314;
    }}

    .student-details {{
      margin-top: 8px;
      font-size: 13px;
      line-height: 2.0;
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

    /* Full Page A4 Landscape Print Rules */
    @media print {{
      @page {{
        size: A4 landscape;
        margin: 0;
      }}

      html, body {{
        width: 297mm;
        height: 210mm;
        margin: 0 !important;
        padding: 0 !important;
        background: #fdf6e2 !important;
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
        padding: 6mm !important;
        margin: 0 !important;
        page-break-after: avoid;
        page-break-inside: avoid;
      }}
    }}
  </style>
</head>
<body>

  <div class="print-bar">
    <button class="btn-print" onclick="window.print()">🖨️ पूरा पेज प्रिंट करें (Full A4 Landscape)</button>
  </div>

  <div class="page-sheet">
    <div class="outer-border">
      <div class="inner-border">

        <!-- LEFT COLUMN -->
        <div class="column left-col">
          <div>
            <div class="section-title">शैक्षिक मूल्यांकन ग्रेडिंग सिस्टम (प्रतिशत में)</div>
            <table class="grade-table">
              <tr>
                <th>91-100</th>
                <th>81-90</th>
                <th>71-80</th>
                <th>61-70</th>
                <th>51-60</th>
                <th>41-50</th>
                <th>33-40</th>
                <th>0-32</th>
              </tr>
              <tr>
                <td>A1</td>
                <td>A2</td>
                <td>B1</td>
                <td>B2</td>
                <td>C1</td>
                <td>C2</td>
                <td>D</td>
                <td>E</td>
              </tr>
            </table>

            <div class="section-title">जीवन कौशल व्यवहार एवं जीवन मूल्य तथा अभिरुचि मूल्यांकन<br>(ग्रेडिंग सिस्टम)</div>
            <table class="grade-table">
              <tr>
                <th>A</th>
                <th>B</th>
                <th>C</th>
                <th>D</th>
                <th>E</th>
              </tr>
              <tr>
                <td>उत्कृष्ट</td>
                <td>अति<br>उत्तम</td>
                <td>उत्तम</td>
                <td>संतोषजनक</td>
                <td>सुधार<br>अपेक्षित</td>
              </tr>
            </table>
          </div>

          <div>
            <div class="instructions-header">
              <span class="instructions-badge">-: सामान्य निर्देश :-</span>
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

</body>
</html>
"""

# Streamlit mein isolated iframe ke andar render karein taaki print bilkul clean aur full-page aaye
components.html(html_template, height=860, scrolling=True)
