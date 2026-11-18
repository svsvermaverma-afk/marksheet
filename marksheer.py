import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
from weasyprint import HTML

st.set_page_config(page_title="ABIC Complete Marksheet System", layout="wide")

# ================= STREAM & SUBJECT CONFIGURATION =================
STREAM_CONFIG = {
    "Class 11-12 Science (PCM / PCB)": [
        ("Sub1", "हिन्दी", False, 100, 0),
        ("Sub2", "अंग्रेजी", False, 100, 0),
        ("Sub3", "गणित / जीव विज्ञान", False, 100, 0),
        ("Sub4", "भौतिक विज्ञान", True, 70, 30),
        ("Sub5", "रसायन विज्ञान", True, 70, 30),
    ],
    "Class 11-12 Commerce": [
        ("Sub1", "हिन्दी", False, 100, 0),
        ("Sub2", "अंग्रेजी", False, 100, 0),
        ("Sub3", "व्यवसाय अध्ययन (BST)", False, 100, 0),
        ("Sub4", "लेखाशास्त्र (Accountancy)", False, 100, 0),
        ("Sub5", "कम्प्यूटर", True, 70, 30),
    ],
    "Class 11-12 Arts": [
        ("Sub1", "हिन्दी", False, 100, 0),
        ("Sub2", "अंग्रेजी", False, 100, 0),
        ("Sub3", "इतिहास (History)", False, 100, 0),
        ("Sub4", "समाजशास्त्र (Sociology)", False, 100, 0),
        ("Sub5", "कम्प्यूटर", True, 70, 30),
    ],
    "Class 11-12 Home Science": [
        ("Sub1", "हिन्दी", False, 100, 0),
        ("Sub2", "अंग्रेजी", False, 100, 0),
        ("Sub3", "गृह विज्ञान (Home Science)", True, 70, 30),
        ("Sub4", "कम्प्यूटर", True, 70, 30),
        ("Sub5", "समाजशास्त्र (Sociology)", False, 100, 0),
    ]
}

# ================= FORMULA HELPERS =================
def get_grade(marks, max_marks):
    if max_marks <= 0:
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
        if pd.isna(val) or str(val).strip() == "":
            return 0
        return int(round(float(val)))
    except:
        return 0

def fmt(val):
    return str(val) if val > 0 else ""

# ================= FORMULA-BASED EXCEL TEMPLATE GENERATOR =================
def get_stream_excel_template(stream_name):
    subjects = STREAM_CONFIG[stream_name]
    sample = {
        "Roll_No": "44", "Reg_No": "38940", "Name": "Anandi Kumari Singh",
        "Father_Name": "Shriram Singh", "Mother_Name": "Lalita Devi",
        "DOB": "26-11-2009", "Class_Sec": "11 'B'", "House": "Arun",
        "Mobile": "9450320524", "Aadhaar": "982931653622", "PEN_No": "21473874271",
        "Address": "Renukoot, Sonbhadra",
    }
    for code, name, has_pr, max_th, max_pr in subjects:
        sample[f"{code}_M1"] = 0
        sample[f"{code}_M2"] = 0
        sample[f"{code}_M3"] = 0
        sample[f"{code}_HY_Th"] = 56 if has_pr else 78
        sample[f"{code}_HY_Pr"] = 28 if has_pr else 0
        sample[f"{code}_AN_Th"] = 0
        sample[f"{code}_AN_Pr"] = 0

    sample["Moral_Sports_Grade"] = "A"
    for i in range(1, 16):
        sample[f"CS_{i}"] = "A"
    sample["Attendance"] = "110/120"
    sample["Session"] = "2025 - 2026"

    df = pd.DataFrame([sample])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Marks_Entry")
        legend_data = [{"Column_Prefix": s[0], "Subject_Name": s[1], "Max_Theory": s[3], "Max_Practical": s[4]} for s in subjects]
        pd.DataFrame(legend_data).to_excel(writer, index=False, sheet_name="Subject_Guide")
    return buffer.getvalue()

# ================= HTML BUILDER =================
def build_marksheet_html(df, stream_name, is_preview=False):
    subjects = STREAM_CONFIG[stream_name]
    pages_html = ""
    target_df = df.iloc[:1] if is_preview else df
    
    for _, row in target_df.iterrows():
        rows_html = ""
        total_m_all = 0
        total_hy_all = 0
        total_an_all = 0
        total_grand_th = 0
        total_grand_pr = 0
        total_grand_all = 0
        
        has_monthly = False
        has_annual = False
        failed_subjects = 0

        for code, _, _, _, _ in subjects:
            if (to_num(row.get(f"{code}_M1")) + to_num(row.get(f"{code}_M2")) + to_num(row.get(f"{code}_M3"))) > 0:
                has_monthly = True
            if (to_num(row.get(f"{code}_AN_Th")) + to_num(row.get(f"{code}_AN_Pr"))) > 0:
                has_annual = True

        for idx, (code, sub_name, has_pr, max_th, max_pr) in enumerate(subjects, start=1):
            m1 = to_num(row.get(f"{code}_M1"))
            m2 = to_num(row.get(f"{code}_M2"))
            m3 = to_num(row.get(f"{code}_M3"))
            m_tot = m1 + m2 + m3
            
            hy_th = to_num(row.get(f"{code}_HY_Th"))
            hy_pr = to_num(row.get(f"{code}_HY_Pr")) if has_pr else 0
            hy_tot = hy_th + hy_pr
            
            an_th = to_num(row.get(f"{code}_AN_Th"))
            an_pr = to_num(row.get(f"{code}_AN_Pr")) if has_pr else 0
            an_tot = an_th + an_pr
            
            sub_grand_th = m_tot + hy_th + an_th
            sub_grand_pr = hy_pr + an_pr
            sub_grand_tot = sub_grand_th + sub_grand_pr
            
            if has_annual:
                max_possible = 200 + (30 if has_monthly else 0)
                eval_marks = sub_grand_tot
            else:
                max_possible = 100 + (30 if has_monthly else 0)
                eval_marks = m_tot + hy_tot

            grade = get_grade(eval_marks, max_possible)
            if grade == "E":
                failed_subjects += 1

            total_m_all += m_tot
            total_hy_all += hy_tot
            total_an_all += an_tot
            total_grand_th += sub_grand_th
            total_grand_pr += sub_grand_pr
            total_grand_all += sub_grand_tot

            rows_html += f"""
            <tr>
                <td>{idx}</td>
                <td class="left">{sub_name}</td>
                <td>{fmt(m1)}</td><td>{fmt(m2)}</td><td>{fmt(m3)}</td><td class="bold">{fmt(m_tot)}</td>
                <td>{fmt(hy_th) if has_pr else ''}</td>
                <td>{fmt(hy_pr) if has_pr else ''}</td>
                <td class="bold">{fmt(hy_tot)}</td>
                <td>{fmt(an_th) if has_pr else ''}</td>
                <td>{fmt(an_pr) if has_pr else ''}</td>
                <td class="bold">{fmt(an_tot)}</td>
                <td>{fmt(sub_grand_th) if (has_annual and has_pr) else ''}</td>
                <td>{fmt(sub_grand_pr) if (has_annual and has_pr) else ''}</td>
                <td class="bold">{fmt(sub_grand_tot) if has_annual else ''}</td>
                <td class="bold">{grade}</td>
            </tr>
            """

        moral_grade = str(row.get('Moral_Sports_Grade', 'A')).strip() or 'A'
        rows_html += f"""
        <tr>
            <td>6</td>
            <td class="left">नैतिक, खेल एवं शारीरिक शिक्षा</td>
            <td></td><td></td><td></td><td></td>
            <td></td><td></td><td class="bold">{moral_grade}</td>
            <td></td><td></td><td class="bold">{moral_grade if has_annual else ''}</td>
            <td></td><td></td><td></td>
            <td class="bold">{moral_grade}</td>
        </tr>
        """

        num_subs = len(subjects)
        if has_annual:
            grand_max = num_subs * (200 + (30 if has_monthly else 0))
            final_obtained = total_grand_all
        else:
            grand_max = num_subs * (100 + (30 if has_monthly else 0))
            final_obtained = total_m_all + total_hy_all

        overall_pct = round((final_obtained / grand_max) * 100, 2) if grand_max > 0 else 0
        overall_grade = get_grade(final_obtained, grand_max)
        
        if str(row.get('Result', '')).strip() != "":
            result_status = str(row.get('Result'))
        else:
            if failed_subjects == 0:
                result_status = f"उत्तीर्ण ({overall_pct}%)"
            elif failed_subjects == 1:
                result_status = f"पूरक / सप्लीमेंट्री ({overall_pct}%)"
            else:
                result_status = f"अनुत्तीर्ण ({overall_pct}%)"

        pages_html += f"""
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
                <div class="acad-sec">
                    <div class="table-caption">शैक्षिक मूल्यांकन</div>
                    <table class="marks-table">
                        <thead>
                            <tr>
                                <th rowspan="2">क्र.</th>
                                <th rowspan="2" style="width:105px;">विषय</th>
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
                            {rows_html}
                            <tr class="bold" style="background-color:#f9efd2;">
                                <td colspan="2">योग (Total)</td>
                                <td colspan="4">{fmt(total_m_all)}</td>
                                <td colspan="3">{total_hy_all} / {num_subs*100}</td>
                                <td colspan="3">{f"{total_an_all} / {num_subs*100}" if has_annual else ""}</td>
                                <td colspan="3">{f"{total_grand_all} / {grand_max}" if has_annual else f"{final_obtained} / {grand_max}"}</td>
                                <td>{overall_grade}</td>
                            </tr>
                        </tbody>
                    </table>

                    <table class="sign-table">
                        <tr>
                            <td style="width:36%; text-align:left;">
                                <div>उपस्थिति : <b>{row.get('Attendance', '')}</b></div>
                                <div>परीक्षाफल : <b>{result_status}</b></div>
                                <div class="sig-line">ह० कक्षाध्यापक</div>
                                <div class="sig-line">ह० निरीक्षक</div>
                                <div class="sig-line">ह० प्रधानाचार्य</div>
                                <div class="sig-line">ह० अभिभावक</div>
                            </td>
                            <td style="width:34%; vertical-align:top;">
                                <b>विशेष विवरण</b>
                                <div style="margin-top:8px; font-size:8pt;">प्रतिशत: <b>{overall_pct}%</b> | समग्र ग्रेड: <b>{overall_grade}</b></div>
                            </td>
                            <td style="width:30%; text-align:left; vertical-align:bottom;">
                                <div class="sig-line">ह० कक्षाध्यापक</div>
                                <div class="sig-line">ह० निरीक्षक</div>
                                <div class="sig-line">ह० प्रधानाचार्य</div>
                                <div class="sig-line">ह० अभिभावक</div>
                            </td>
                        </tr>
                    </table>
                </div>

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

    preview_css = """
        body { background-color: #2b2b2b; padding: 15px; display: flex; flex-direction: column; align-items: center; gap: 20px; }
        .sheet { box-shadow: 0 8px 24px rgba(0,0,0,0.45); transform-origin: top center; }
    """ if is_preview else ""

    html_doc = f"""
    <!DOCTYPE html>
    <html lang="hi">
    <head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        @page {{ size: A4 landscape; margin: 0; }}
        body {{
            margin: 0; padding: 0;
            font-family: 'Noto Sans Devanagari', sans-serif;
            color: #6b1d1d;
            -webkit-print-color-adjust: exact;
        }}
        .sheet {{
            width: 297mm; height: 210mm;
            background-color: #fbf5df;
            box-sizing: border-box;
            padding: 8mm;
            display: flex;
            page-break-after: always;
        }}
        .half-left, .half-right {{ width: 50%; height: 100%; box-sizing: border-box; }}
        .half-left {{ padding-right: 5mm; border-right: 1px dashed #999; }}
        .half-right {{ padding-left: 5mm; }}
        .box-border {{ border: 2px solid #7a2020; height: 100%; padding: 6mm; box-sizing: border-box; }}
        .sec-title {{ text-align: center; font-weight: 700; font-size: 11pt; margin-bottom: 6px; }}
        .grid-table {{ width: 100%; border-collapse: collapse; text-align: center; font-size: 9.5pt; }}
        .grid-table td {{ border: 1px solid #7a2020; padding: 5px 2px; }}
        .badge-center {{
            background: #7a2020; color: #fff; font-weight: 700; text-align: center;
            width: 150px; margin: 22px auto 12px auto; padding: 4px 10px; border-radius: 4px; font-size: 10.5pt;
        }}
        .inst-list {{ font-size: 9.5pt; line-height: 1.7; padding-left: 18px; }}
        .top-meta {{ display: flex; justify-content: space-between; font-size: 7.5pt; border-bottom: 1px solid #7a2020; padding-bottom: 4px; }}
        .school-header {{ text-align: center; margin-top: 10px; }}
        .school-header h1 {{ margin: 4px 0 2px 0; font-size: 17pt; font-weight: 700; }}
        .school-header h3 {{ margin: 0 0 8px 0; font-size: 12pt; }}
        .main-badge {{ display: inline-block; background: #7a2020; color: #fff; font-size: 13pt; font-weight: 700; padding: 3px 24px; border-radius: 6px; }}
        .session-txt {{ font-weight: 700; font-size: 11.5pt; margin-top: 6px; }}
        .student-fields {{ margin-top: 12px; font-size: 10pt; font-weight: 600; }}
        .field-row {{ border-bottom: 1px dotted #7a2020; padding: 5px 0 2px 0; display: flex; justify-content: space-between; }}
        .val {{ color: #1a237e; font-weight: 700; padding-left: 6px; }}
        .inside-page {{ flex-direction: column; padding: 6mm 8mm; }}
        .inside-header {{ text-align: center; }}
        .inside-header h2 {{ margin: 0; font-size: 15.5pt; }}
        .sub-h {{ font-weight: 700; font-size: 10.5pt; }}
        .stu-strip {{ display: flex; justify-content: space-between; font-size: 9.5pt; margin-top: 4px; border-bottom: 1px solid #7a2020; padding-bottom: 3px; }}
        .inside-body {{ display: flex; gap: 6mm; margin-top: 5px; flex: 1; }}
        .acad-sec {{ width: 62%; }}
        .coschol-sec {{ width: 38%; }}
        .table-caption {{ text-align: center; font-weight: 700; font-size: 11pt; margin-bottom: 3px; }}
        .marks-table, .co-table, .sign-table {{ width: 100%; border-collapse: collapse; text-align: center; font-size: 8.2pt; }}
        .marks-table th, .marks-table td, .co-table th, .co-table td, .sign-table td {{ border: 1px solid #7a2020; padding: 3.5px 2px; }}
        .marks-table th {{ background-color: #f3e5be; font-weight: 700; }}
        .left {{ text-align: left; padding-left: 5px !important; }}
        .bold {{ font-weight: 700; }}
        .sign-table {{ margin-top: 6px; font-size: 8.5pt; }}
        .sig-line {{ margin-top: 11px; font-weight: 600; }}
        .co-badge {{ background: #7a2020; color: #fff; text-align: center; font-weight: 700; font-size: 10pt; padding: 2px 12px; border-radius: 4px; width: 160px; margin: 0 auto 4px auto; }}
        .co-table td {{ padding: 2.3px 4px; font-size: 8pt; }}
        .cat-row td {{ background-color: #f3e5be; font-weight: 700; }}
        {preview_css}
    </style>
    </head>
    <body>{pages_html}</body>
    </html>
    """
    return html_doc

# ================= STREAMLIT UI =================
st.title("🎓 ABIC Renukoot - Marksheet Generator")

selected_stream = st.selectbox(
    "📌 Step 1: Stream Chunein (Subject Formula Set Karne Ke Liye)",
    list(STREAM_CONFIG.keys())
)

active_subs = STREAM_CONFIG[selected_stream]
cols_info = st.columns(len(active_subs))
for idx, (code, sname, has_pr, m_th, m_pr) in enumerate(active_subs):
    cols_info[idx].info(f"**{sname}**\n\n{'Theory: '+str(m_th)+' + Prac: '+str(m_pr) if has_pr else 'Max Marks: '+str(m_th)}")

st.markdown("### 📂 Step 2: Excel File Upload Karein")
col_t1, col_t2 = st.columns([1, 2])
with col_t1:
    st.download_button(
        label=f"📥 Download Sample Excel ({selected_stream})",
        data=get_stream_excel_template(selected_stream),
        file_name=f"ABIC_{selected_stream.replace(' ', '_')}_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with col_t2:
    uploaded_file = st.file_uploader("Marks Bhari Hui Excel (.xlsx) Upload Karein", type=["xlsx"])

if uploaded_file:
    df_uploaded = pd.read_excel(uploaded_file).fillna("")
    df_uploaded["Roll_No_Str"] = df_uploaded["Roll_No"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_uploaded["Reg_No_Str"] = df_uploaded["Reg_No"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_uploaded["Class_Sec_Str"] = df_uploaded["Class_Sec"].astype(str).str.strip()

    tab1, tab2 = st.tabs(["📚 Bulk Preview & All Download", "🔍 Single Student (By Class & Roll No / SR No)"])

    with tab1:
        st.success(f"✅ Total **{len(df_uploaded)} students** ka data load ho gaya hai.")
        
        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            if st.button("🚀 Generate All Marksheets PDF (Bulk)", type="primary"):
                with st.spinner(f"Sabhi {len(df_uploaded)} students ki marksheet PDF taiyar ho rahi hai..."):
                    full_html = build_marksheet_html(df_uploaded, selected_stream, is_preview=False)
                    st.session_state["bulk_pdf"] = HTML(string=full_html).write_pdf()
                    st.session_state["bulk_count"] = len(df_uploaded)

        with col_b2:
            if "bulk_pdf" in st.session_state:
                st.download_button(
                    label=f"📥 Download All {st.session_state['bulk_count']} Marksheets (Single PDF)",
                    data=st.session_state["bulk_pdf"],
                    file_name=f"ABIC_{selected_stream.replace(' ', '_')}_All_Marksheets.pdf",
                    mime="application/pdf"
                )

        st.markdown("---")
        st.subheader(f"👁️ Sample Marksheet Preview (First Student: {df_uploaded.iloc[0]['Name']})")
        preview_html = build_marksheet_html(df_uploaded.iloc[[0]], selected_stream, is_preview=True)
        components.html(preview_html, height=1680, scrolling=True)

    with tab2:
        st.subheader("🔍 Class aur Roll No / SR No (Reg No) se Student Khojein")
        available_classes = sorted(df_uploaded["Class_Sec_Str"].unique().tolist())
        
        sc1, sc2, sc3 = st.columns(3)
        selected_class = sc1.selectbox("1. Class & Section Chunein", available_classes)
        search_type = sc2.radio("2. Kis se khojna hai?", ["Roll Number", "SR Number (Reg No)"], horizontal=True)
        search_val = sc3.text_input("3. Roll No ya SR No likhein", placeholder="Jaise: 44 ya 38940").strip()

        class_df = df_uploaded[df_uploaded["Class_Sec_Str"] == selected_class]

        if search_val:
            if search_type == "Roll Number":
                matched_df = class_df[class_df["Roll_No_Str"] == search_val]
            else:
                matched_df = class_df[class_df["Reg_No_Str"] == search_val]

            if not matched_df.empty:
                stu = matched_df.iloc[0]
                st.success(f"🎯 Student Mila: **{stu['Name']}** | Class: **{stu['Class_Sec_Str']}** | Roll No: **{stu['Roll_No_Str']}** | SR No: **{stu['Reg_No_Str']}**")
                
                single_full_html = build_marksheet_html(matched_df.iloc[[0]], selected_stream, is_preview=False)
                single_preview_html = build_marksheet_html(matched_df.iloc[[0]], selected_stream, is_preview=True)
                
                with st.spinner("Single Marksheet PDF taiyar ho rahi hai..."):
                    single_pdf_bytes = HTML(string=single_full_html).write_pdf()
                
                st.download_button(
                    label=f"📄 Download {stu['Name']} Marksheet PDF",
                    data=single_pdf_bytes,
                    file_name=f"{stu['Roll_No_Str']}_{stu['Name']}_Marksheet.pdf",
                    mime="application/pdf"
                )
                
                st.markdown("---")
                st.subheader(f"👁️ Marksheet Preview: {stu['Name']}")
                components.html(single_preview_html, height=1680, scrolling=True)
            else:
                st.error(f"❌ Class **{selected_class}** mein {search_type} **`{search_val}`** ka koi student nahi mila. Kripya number check karein.")
        else:
            st.info(f"ℹ️ Class **{selected_class}** mein kul **{len(class_df)} students** hain. Upar box mein **Roll No** ya **SR No** daalte hi marksheet dikh jayegi.")
