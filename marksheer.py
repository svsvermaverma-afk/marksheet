import streamlit as st
import pandas as pd
import io
from weasyprint import HTML

st.set_page_config(page_title="ABIC Marksheet Generator", layout="wide")

# Grade Calculation Helper
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

def to_num(val):
    try:
        return int(float(val)) if str(val).strip() != "" else 0
    except:
        return 0

# Sample Excel Generator
def get_sample_excel():
    data = [{
        "Roll_No": "44", "Reg_No": "38940", "Name": "Anandi Kumari Singh",
        "Father_Name": "Shriram Singh", "Mother_Name": "Lalita Devi",
        "DOB": "26-11-2009", "Class_Sec": "11 'B'", "House": "Arun",
        "Mobile": "9450320524", "Aadhaar": "982931653622", "PEN_No": "21473874271",
        "Address": "Jhopadi Near Bus Stand, Renukoot",
        "Hindi_HY": 80, "English_HY": 65, "Math_Bio_HY": 76,
        "Physics_HY_Th": 56, "Physics_HY_Pr": 29,
        "Chem_HY_Th": 38, "Chem_HY_Pr": 28,
        "Moral_Sports_Grade": "A",
        "Hindi_AN": 0, "English_AN": 0, "Math_Bio_AN": 0,
        "Physics_AN_Th": 0, "Physics_AN_Pr": 0,
        "Chem_AN_Th": 0, "Chem_AN_Pr": 0,
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

# HTML + CSS Template for Exact Marksheet Design
def generate_marksheet_pdf(df):
    pages_html = ""
    
    for _, row in df.iterrows():
        h_hy, e_hy, mb_hy = to_num(row.get('Hindi_HY')), to_num(row.get('English_HY')), to_num(row.get('Math_Bio_HY'))
        p_hy_th, p_hy_pr = to_num(row.get('Physics_HY_Th')), to_num(row.get('Physics_HY_Pr'))
        c_hy_th, c_hy_pr = to_num(row.get('Chem_HY_Th')), to_num(row.get('Chem_HY_Pr'))
        
        p_hy_tot = p_hy_th + p_hy_pr
        c_hy_tot = c_hy_th + c_hy_pr
        hy_grand = h_hy + e_hy + mb_hy + p_hy_tot + c_hy_tot
        
        h_an, e_an, mb_an = to_num(row.get('Hindi_AN')), to_num(row.get('English_AN')), to_num(row.get('Math_Bio_AN'))
        p_an_th, p_an_pr = to_num(row.get('Physics_AN_Th')), to_num(row.get('Physics_AN_Pr'))
        c_an_th, c_an_pr = to_num(row.get('Chem_AN_Th')), to_num(row.get('Chem_AN_Pr'))
        
        p_an_tot = p_an_th + p_an_pr
        c_an_tot = c_an_th + c_an_pr
        an_grand = h_an + e_an + mb_an + p_an_tot + c_an_tot
        has_an = an_grand > 0
        
        def fmt(v): return str(v) if v > 0 else ""
        def g_calc(hy, an):
            tot = hy + an if has_an else hy
            mx = 200 if has_an else 100
            return get_grade(tot, mx)

        pages_html += f"""
        <!-- PAGE 1: OUTER COVER -->
        <div class="sheet">
            <div class="half-left">
                <div class="box-border">
                    <div class="sec-title">शैक्षिक मूल्यांकन ग्रेडिंग सिस्टम (प्रतिशत में)</div>
                    <table class="grid-table">
                        <tr><td>91-100</td><td>81-90</td><td>71-80</td><td>61-70</td><td>51-60</td><td>41-50</td><td>33-40</td><td>0-32</td></tr>
                        <tr class="bold"><td>A1</td><td>A2</td><td>B1</td><td>B2</td><td>C1</td><td>C2</td><td>D</td><td>E</td></tr>
                    </table>

                    <div class="sec-title" style="margin-top:18px;">जीवन कौशल व्यवहार एवं जीवन मूल्य तथा अभिरुचि मूल्यांकन (ग्रेडिंग सिस्टम)</div>
                    <table class="grid-table">
                        <tr class="bold"><td>A</td><td>B</td><td>C</td><td>D</td><td>E</td></tr>
                        <tr><td>उत्कृष्ट</td><td>अति उत्तम</td><td>उत्तम</td><td>संतोषजनक</td><td>सुधार अपेक्षित</td></tr>
                    </table>

                    <div class="badge-center">-: सामान्य निर्देश :-</div>
                    <ol class="inst-list">
                        <li>संरक्षक अपने पाल्य/पाल्या के अर्द्धवार्षिक प्राप्तांक देखकर, इस प्रगति पत्र पर हस्ताक्षर करके तीन दिन के अंदर विद्यालय में कक्षा अध्यापक के पास जमा कर दें।</li>
                        <li>इस प्रगति पत्र के खो जाने पर दूसरी प्रतिलिपि के लिए ................... रुपये का शुल्क देना पड़ेगा।</li>
                        <li>प्रगति पत्र पर अंकित प्राप्तांक के साथ किसी भी तरह की छेड़छाड़ न की जाए।</li>
                    </ol>
                </div>
            </div>

            <div class="half-right">
                <div class="box-border front-cover">
                    <div class="top-meta">
                        <span>विद्यालय की वेबसाइट: <b>www.adityabirlaschools.co.in/abic-renukoot</b></span>
                        <span>विद्यालय का कोड: <b>1023</b></span>
                    </div>
                    <div class="school-header">
                        <h1>आदित्य बिड़ला इण्टरमीडिएट कॉलेज</h1>
                        <h3>रेणुकूट, सोनभद्र (उ०प्र०)</h3>
                        <div class="main-badge">प्रगति-पत्र</div>
                        <div class="session-txt">शैक्षणिक सत्र {row.get('Session', '2025 - 2026')}</div>
                    </div>

                    <div class="student-fields">
                        <div class="field-row">छात्र / छात्रा का नाम : <span class="val">{row.get('Name', '')}</span></div>
                        <div class="field-row">
                            <span>अनुक्रमांक : <span class="val">{row.get('Roll_No', '')}</span></span>
                            <span>निलय : <span class="val">{row.get('House', '')}</span></span>
                        </div>
                        <div class="field-row">पिता का नाम : <span class="val">{row.get('Father_Name', '')}</span></div>
                        <div class="field-row">माता का नाम : <span class="val">{row.get('Mother_Name', '')}</span></div>
                        <div class="field-row">
                            <span>जन्मतिथि : <span class="val">{row.get('DOB', '')}</span></span>
                            <span>कक्षा एवं वर्ग : <span class="val">{row.get('Class_Sec', '')}</span></span>
                        </div>
                        <div class="field-row">मोबाइल नंबर : <span class="val">{row.get('Mobile', '')}</span></div>
                        <div class="field-row">छात्र पंजीकरण संख्या : <span class="val">{row.get('Reg_No', '')}</span></div>
                        <div class="field-row">आधार संख्या : <span class="val">{row.get('Aadhaar', '')}</span></div>
                        <div class="field-row">पेन संख्या (PEN No.) : <span class="val">{row.get('PEN_No', '')}</span></div>
                        <div class="field-row">पता : <span class="val">{row.get('Address', '')}</span></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- PAGE 2: INSIDE MARKSHEET -->
        <div class="sheet inside-page">
            <div class="inside-header">
                <h2>आदित्य बिड़ला इण्टरमीडिएट कॉलेज, रेणुकूट, सोनभद्र</h2>
                <div class="sub-h">प्रगति पत्र विवरण सत्र {row.get('Session', '2025 - 2026')}</div>
                <div class="stu-strip">
                    <span>छात्र/छात्रा का नाम: <b>{row.get('Name', '')}</b></span>
                    <span>छात्र पंजिका संख्या: <b>{row.get('Reg_No', '')}</b></span>
                    <span>अनुक्रमांक: <b>{row.get('Roll_No', '')}</b></span>
                    <span>कक्षा एवं वर्ग: <b>{row.get('Class_Sec', '')}</b></span>
                </div>
            </div>

            <div class="inside-body">
                <!-- ACADEMIC TABLE -->
                <div class="acad-sec">
                    <div class="table-caption">शैक्षिक मूल्यांकन</div>
                    <table class="marks-table">
                        <thead>
                            <tr>
                                <th rowspan="2">क्र.</th>
                                <th rowspan="2" style="width:95px;">विषय</th>
                                <th colspan="4">मासिक परीक्षा</th>
                                <th colspan="3">अर्द्धवार्षिक परीक्षा</th>
                                <th colspan="3">वार्षिक परीक्षा</th>
                                <th colspan="3">सम्पूर्ण परीक्षा</th>
                                <th rowspan="2">ग्रेड</th>
                            </tr>
                            <tr>
                                <th>प्र.</th><th>द्वि.</th><th>तृ.</th><th>योग</th>
                                <th>सैद्ध.</th><th>प्रायो.</th><th>योग</th>
                                <th>सैद्ध.</th><th>प्रायो.</th><th>योग</th>
                                <th>सैद्ध.</th><th>प्रायो.</th><th>योग</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td><td class="left">हिन्दी</td>
                                <td></td><td></td><td></td><td></td>
                                <td></td><td></td><td class="bold">{h_hy}</td>
                                <td></td><td></td><td class="bold">{fmt(h_an)}</td>
                                <td></td><td></td><td class="bold">{fmt(h_hy+h_an) if has_an else ''}</td>
                                <td class="bold">{g_calc(h_hy, h_an)}</td>
                            </tr>
                            <tr>
                                <td>2</td><td class="left">अंग्रेजी</td>
                                <td></td><td></td><td></td><td></td>
                                <td></td><td></td><td class="bold">{e_hy}</td>
                                <td></td><td></td><td class="bold">{fmt(e_an)}</td>
                                <td></td><td></td><td class="bold">{fmt(e_hy+e_an) if has_an else ''}</td>
                                <td class="bold">{g_calc(e_hy, e_an)}</td>
                            </tr>
                            <tr>
                                <td>3</td><td class="left">गणित/जीव विज्ञान</td>
                                <td></td><td></td><td></td><td></td>
                                <td></td><td></td><td class="bold">{mb_hy}</td>
                                <td></td><td></td><td class="bold">{fmt(mb_an)}</td>
                                <td></td><td></td><td class="bold">{fmt(mb_hy+mb_an) if has_an else ''}</td>
                                <td class="bold">{g_calc(mb_hy, mb_an)}</td>
                            </tr>
                            <tr>
                                <td>4</td><td class="left">भौतिक विज्ञान</td>
                                <td></td><td></td><td></td><td></td>
                                <td>{p_hy_th}</td><td>{p_hy_pr}</td><td class="bold">{p_hy_tot}</td>
                                <td>{fmt(p_an_th)}</td><td>{fmt(p_an_pr)}</td><td class="bold">{fmt(p_an_tot)}</td>
                                <td></td><td></td><td class="bold">{fmt(p_hy_tot+p_an_tot) if has_an else ''}</td>
                                <td class="bold">{g_calc(p_hy_tot, p_an_tot)}</td>
                            </tr>
                            <tr>
                                <td>5</td><td class="left">रसायन विज्ञान</td>
                                <td></td><td></td><td></td><td></td>
                                <td>{c_hy_th}</td><td>{c_hy_pr}</td><td class="bold">{c_hy_tot}</td>
                                <td>{fmt(c_an_th)}</td><td>{fmt(c_an_pr)}</td><td class="bold">{fmt(c_an_tot)}</td>
                                <td></td><td></td><td class="bold">{fmt(c_hy_tot+c_an_tot) if has_an else ''}</td>
                                <td class="bold">{g_calc(c_hy_tot, c_an_tot)}</td>
                            </tr>
                            <tr>
                                <td>6</td><td class="left">नैतिक, खेल एवं शारीरिक शिक्षा</td>
                                <td></td><td></td><td></td><td></td>
                                <td></td><td></td><td class="bold">{row.get('Moral_Sports_Grade', 'A')}</td>
                                <td></td><td></td><td></td>
                                <td></td><td></td><td></td>
                                <td class="bold">{row.get('Moral_Sports_Grade', 'A')}</td>
                            </tr>
                            <tr class="bold">
                                <td colspan="2">योग (Total)</td>
                                <td></td><td></td><td></td><td></td>
                                <td colspan="3">{hy_grand} / 500</td>
                                <td colspan="3">{f"{an_grand} / 500" if has_an else ""}</td>
                                <td colspan="3">{f"{hy_grand+an_grand} / 1000" if has_an else ""}</td>
                                <td>{get_grade(hy_grand+an_grand if has_an else hy_grand, 1000 if has_an else 500)}</td>
                            </tr>
                        </tbody>
                    </table>

                    <table class="sign-table">
                        <tr>
                            <td style="width:32%; text-align:left;">
                                <div>उपस्थिति : <b>{row.get('Attendance', '')}</b></div>
                                <div>परीक्षाफल : <b>{row.get('Result', '')}</b></div>
                                <div class="sig-line">ह० कक्षाध्यापक</div>
                                <div class="sig-line">ह० निरीक्षक</div>
                                <div class="sig-line">ह० प्रधानाचार्य</div>
                                <div class="sig-line">ह० अभिभावक</div>
                            </td>
                            <td style="width:38%; vertical-align:top;">विशेष विवरण</td>
                            <td style="width:30%; text-align:left; vertical-align:bottom;">
                                <div class="sig-line">ह० कक्षाध्यापक</div>
                                <div class="sig-line">ह० निरीक्षक</div>
                                <div class="sig-line">ह० प्रधानाचार्य</div>
                                <div class="sig-line">ह० अभिभावक</div>
                            </td>
                        </tr>
                    </table>
                </div>

                <!-- CO-SCHOLASTIC TABLE -->
                <div class="coschol-sec">
                    <div class="co-badge">सह शैक्षणिक मूल्यांकन</div>
                    <table class="co-table">
                        <thead>
                            <tr><th style="width:28px;">क्रम</th><th>मूल्यांकन का क्षेत्र</th><th style="width:38px;">ग्रेड</th></tr>
                        </thead>
                        <tbody>
                            <tr class="cat-row"><td colspan="3">जीवन कौशल</td></tr>
                            <tr><td>1</td><td class="left">सृजनात्मकता</td><td class="bold">{row.get('CS_1','A')}</td></tr>
                            <tr><td>2</td><td class="left">प्रभावी सम्प्रेषण</td><td class="bold">{row.get('CS_2','A')}</td></tr>
                            <tr><td>3</td><td class="left">निर्णय लेने की क्षमता एवं नेतृत्व क्षमता</td><td class="bold">{row.get('CS_3','A')}</td></tr>
                            <tr><td>4</td><td class="left">जिज्ञासा, तार्किक चिंतन एवं समस्या का समाधान</td><td class="bold">{row.get('CS_4','A')}</td></tr>
                            <tr><td>5</td><td class="left">तनाव प्रबन्धन</td><td class="bold">{row.get('CS_5','A')}</td></tr>
                            <tr class="cat-row"><td colspan="3">व्यवहार एवं जीवन मूल्य</td></tr>
                            <tr><td>6</td><td class="left">शिक्षकों से व्यवहार</td><td class="bold">{row.get('CS_6','A')}</td></tr>
                            <tr><td>7</td><td class="left">सहपाठियों से व्यवहार</td><td class="bold">{row.get('CS_7','A')}</td></tr>
                            <tr><td>8</td><td class="left">शिष्टाचार, अनुशासन, समयबद्धता</td><td class="bold">{row.get('CS_8','A')}</td></tr>
                            <tr><td>9</td><td class="left">व्यक्तिगत एवं विद्यालयी स्वच्छता</td><td class="bold">{row.get('CS_9','A')}</td></tr>
                            <tr><td>10</td><td class="left">पर्यावरण के प्रति जागरूकता</td><td class="bold">{row.get('CS_10','A')}</td></tr>
                            <tr><td>11</td><td class="left">सहयोग, सामूहिकता, सहानुभूति</td><td class="bold">{row.get('CS_11','A')}</td></tr>
                            <tr class="cat-row"><td colspan="3">अभिरुचि</td></tr>
                            <tr><td>12</td><td class="left">खेल एवं योग</td><td class="bold">{row.get('CS_12','A')}</td></tr>
                            <tr><td>13</td><td class="left">पाठ्य सहगामी गतिविधियाँ</td><td class="bold">{row.get('CS_13','A')}</td></tr>
                            <tr><td>14</td><td class="left">कला / संगीत / साहित्यिक गतिविधियाँ</td><td class="bold">{row.get('CS_14','A')}</td></tr>
                            <tr><td>15</td><td class="left">वैज्ञानिक अभिरुचि</td><td class="bold">{row.get('CS_15','A')}</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        @page {{
            size: A4 landscape;
            margin: 0;
        }}
        body {{
            margin: 0;
            padding: 0;
            font-family: 'Noto Sans Devanagari', sans-serif;
            color: #6b1d1d;
            -webkit-print-color-adjust: exact;
        }}
        .sheet {{
            width: 297mm;
            height: 210mm;
            background-color: #fbf5df;
            box-sizing: border-box;
            padding: 8mm;
            display: flex;
            page-break-after: always;
        }}
        .half-left, .half-right {{
            width: 50%;
            height: 100%;
            box-sizing: border-box;
        }}
        .half-left {{ padding-right: 5mm; border-right: 1px dashed #999; }}
        .half-right {{ padding-left: 5mm; }}
        .box-border {{
            border: 2px solid #7a2020;
            height: 100%;
            padding: 6mm;
            box-sizing: border-box;
        }}
        .sec-title {{
            text-align: center;
            font-weight: 700;
            font-size: 11pt;
            margin-bottom: 6px;
        }}
        .grid-table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
            font-size: 9.5pt;
        }}
        .grid-table td {{
            border: 1px solid #7a2020;
            padding: 5px 2px;
        }}
        .badge-center {{
            background: #7a2020;
            color: #fff;
            font-weight: 700;
            text-align: center;
            width: 150px;
            margin: 22px auto 12px auto;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 10.5pt;
        }}
        .inst-list {{
            font-size: 9.5pt;
            line-height: 1.7;
            padding-left: 18px;
        }}
        .top-meta {{
            display: flex;
            justify-content: space-between;
            font-size: 7.5pt;
            border-bottom: 1px solid #7a2020;
            padding-bottom: 4px;
        }}
        .school-header {{
            text-align: center;
            margin-top: 10px;
        }}
        .school-header h1 {{
            margin: 4px 0 2px 0;
            font-size: 17pt;
            font-weight: 700;
        }}
        .school-header h3 {{
            margin: 0 0 8px 0;
            font-size: 12pt;
        }}
        .main-badge {{
            display: inline-block;
            background: #7a2020;
            color: #fff;
            font-size: 13pt;
            font-weight: 700;
            padding: 3px 24px;
            border-radius: 6px;
        }}
        .session-txt {{
            font-weight: 700;
            font-size: 11.5pt;
            margin-top: 6px;
        }}
        .student-fields {{
            margin-top: 12px;
            font-size: 10pt;
            font-weight: 600;
        }}
        .field-row {{
            border-bottom: 1px dotted #7a2020;
            padding: 5px 0 2px 0;
            display: flex;
            justify-content: space-between;
        }}
        .val {{
            color: #1a237e;
            font-weight: 700;
            padding-left: 6px;
        }}
        /* INSIDE PAGE */
        .inside-page {{
            flex-direction: column;
            padding: 6mm 8mm;
        }}
        .inside-header {{
            text-align: center;
        }}
        .inside-header h2 {{
            margin: 0;
            font-size: 15.5pt;
        }}
        .sub-h {{
            font-weight: 700;
            font-size: 10.5pt;
        }}
        .stu-strip {{
            display: flex;
            justify-content: space-between;
            font-size: 9.5pt;
            margin-top: 4px;
            border-bottom: 1px solid #7a2020;
            padding-bottom: 3px;
        }}
        .inside-body {{
            display: flex;
            gap: 6mm;
            margin-top: 5px;
            flex: 1;
        }}
        .acad-sec {{ width: 61%; }}
        .coschol-sec {{ width: 39%; }}
        .table-caption {{
            text-align: center;
            font-weight: 700;
            font-size: 11pt;
            margin-bottom: 3px;
        }}
        .marks-table, .co-table, .sign-table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
            font-size: 8.5pt;
        }}
        .marks-table th, .marks-table td, .co-table th, .co-table td, .sign-table td {{
            border: 1px solid #7a2020;
            padding: 3.5px 2px;
        }}
        .marks-table th {{ background-color: #f3e5be; font-weight: 700; }}
        .left {{ text-align: left; padding-left: 5px !important; }}
        .bold {{ font-weight: 700; }}
        .sign-table {{ margin-top: 6px; font-size: 8.5pt; }}
        .sig-line {{ margin-top: 11px; font-weight: 600; }}
        .co-badge {{
            background: #7a2020;
            color: #fff;
            text-align: center;
            font-weight: 700;
            font-size: 10pt;
            padding: 2px 12px;
            border-radius: 4px;
            width: 160px;
            margin: 0 auto 4px auto;
        }}
        .co-table td {{ padding: 2.3px 4px; font-size: 8pt; }}
        .cat-row td {{ background-color: #f3e5be; font-weight: 700; }}
    </style>
    </head>
    <body>{pages_html}</body>
    </html>
    """
    
    pdf_bytes = HTML(string=full_html).write_pdf()
    return pdf_bytes

# ================= STREAMLIT UI =================
st.title("🎓 ABIC Renukoot - Pragati Patra (Marksheet) Generator")
st.markdown("Aditya Birla Intermediate College ke official format mein **100% Shuddh Hindi Font** ke saath Marksheet PDF banayein.")

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
            with st.spinner("PDF taiyar ho rahi hai..."):
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
        with st.spinner("PDF taiyar ho rahi hai..."):
            pdf_data = generate_marksheet_pdf(single_df)
        st.success("✅ Marksheet PDF taiyar hai! Neeche button se download karein.")
        st.download_button(
            label="📄 Download PDF",
            data=pdf_data,
            file_name=f"{name}_Marksheet.pdf",
            mime="application/pdf"
        )
