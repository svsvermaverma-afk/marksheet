<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>प्रगति-पत्र - आदित्य बिड़ला इण्टरमीडिएट कॉलेज</title>
  <style>
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: 'Mangal', 'Nirmala UI', 'Noto Sans Devanagari', Arial, sans-serif;
      background-color: #e5e5e5;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 15px;
      color: #2b1d14;
    }

    /* Screen par Print Button (Print mein hide ho jayega) */
    .no-print {
      margin-bottom: 15px;
      display: flex;
      gap: 12px;
    }

    .btn-print {
      background-color: #7c2d12;
      color: #ffffff;
      border: none;
      padding: 10px 24px;
      font-size: 16px;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }

    .btn-print:hover {
      background-color: #9a3412;
    }

    /* A4 Landscape Full Page Container */
    .page-sheet {
      width: 297mm;
      height: 210mm;
      background-color: #fdf6e2; /* Light cream/yellowish tint photo jaisa */
      padding: 8mm;
      box-shadow: 0 0 12px rgba(0, 0, 0, 0.25);
      display: flex;
      justify-content: center;
      align-items: center;
    }

    .outer-border {
      width: 100%;
      height: 100%;
      border: 2.5px solid #5c3a21;
      padding: 3px;
      display: flex;
    }

    .inner-border {
      width: 100%;
      height: 100%;
      border: 1px solid #5c3a21;
      display: grid;
      grid-template-columns: 1fr 1fr;
    }

    /* Left & Right Columns */
    .column {
      padding: 12mm 10mm;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }

    .left-col {
      border-right: 1.5px solid #5c3a21;
    }

    /* Left Column Styling */
    .section-title {
      font-size: 14px;
      font-weight: bold;
      color: #5c2818;
      margin-bottom: 6px;
    }

    table.grade-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 18px;
      text-align: center;
      font-size: 12.5px;
    }

    table.grade-table th,
    table.grade-table td {
      border: 1px solid #5c3a21;
      padding: 6px 4px;
    }

    table.grade-table th {
      font-weight: normal;
      background-color: rgba(92, 58, 33, 0.05);
    }

    table.grade-table td {
      font-weight: bold;
    }

    .instructions-box {
      margin-top: 5px;
    }

    .instructions-header {
      text-align: center;
      margin-bottom: 12px;
    }

    .instructions-badge {
      display: inline-block;
      background-color: #b4533c;
      color: #ffffff;
      font-weight: bold;
      font-size: 14px;
      padding: 3px 16px;
      border-radius: 3px;
    }

    .instructions-list {
      list-style: none;
      font-size: 13px;
      line-height: 1.75;
    }

    .instructions-list li {
      margin-bottom: 8px;
      text-align: justify;
    }

    /* Right Column Styling */
    .top-meta {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      font-size: 11px;
      text-align: center;
    }

    .meta-box {
      border: 1px solid #5c3a21;
      padding: 2px 8px;
      margin-top: 2px;
      font-weight: bold;
      background-color: #fffdf7;
    }

    .school-header {
      text-align: center;
      margin-top: 4px;
    }

    .school-logo {
      width: 72px;
      height: auto;
      display: block;
      margin: 0 auto 6px auto;
    }

    .school-name {
      font-size: 21px;
      font-weight: bold;
      color: #5c2818;
      letter-spacing: 0.3px;
    }

    .school-address {
      font-size: 14px;
      font-weight: bold;
      color: #5c2818;
      margin-top: 2px;
    }

    .report-badge {
      display: inline-block;
      background-color: #b4533c;
      color: #ffffff;
      font-size: 17px;
      font-weight: bold;
      padding: 3px 22px;
      border-radius: 4px;
      margin: 8px 0 6px 0;
    }

    .session-text {
      font-size: 15px;
      font-weight: bold;
      color: #3b2314;
    }

    /* Student Details Form Lines */
    .student-details {
      margin-top: 10px;
      font-size: 13.5px;
      line-height: 2.05;
    }

    .field-row {
      display: flex;
      align-items: baseline;
      width: 100%;
    }

    .field-label {
      white-space: nowrap;
      font-weight: bold;
      color: #4a2c18;
    }

    .field-value {
      flex: 1;
      border-bottom: 1.5px dotted #5c3a21;
      padding-left: 8px;
      font-weight: bold;
      color: #1e3a8a;
      min-height: 20px;
    }

    .field-half {
      display: flex;
      align-items: baseline;
      width: 50%;
    }

    /* Exact Full A4 Landscape Print Rules */
    @media print {
      @page {
        size: A4 landscape;
        margin: 0;
      }

      body {
        background: none;
        padding: 0;
        margin: 0;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }

      .no-print {
        display: none !important;
      }

      .page-sheet {
        width: 297mm;
        height: 210mm;
        box-shadow: none;
        margin: 0;
        padding: 6mm;
        page-break-after: avoid;
        page-break-inside: avoid;
      }
    }
  </style>
</head>
<body>

  <!-- Print Button -->
  <div class="no-print">
    <button class="btn-print" onclick="window.print()">🖨️ पूरा पेज प्रिंट करें (A4 Landscape)</button>
  </div>

  <!-- Marksheet Outer A4 Sheet -->
  <div class="page-sheet">
    <div class="outer-border">
      <div class="inner-border">

        <!-- LEFT COLUMN: Grading System & General Instructions -->
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

          <div class="instructions-box">
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

        <!-- RIGHT COLUMN: School Header, Logo & Student Details -->
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
            <!-- GitHub Repository mein rakhi logo.png file -->
            <img src="./logo.png" alt="ABIC Renukoot Logo" class="school-logo">
            <div class="school-name">आदित्य बिड़ला इण्टरमीडिएट कॉलेज</div>
            <div class="school-address">रेणुकूट, सोनभद्र (उ०प्र०)</div>
            <div class="report-badge">प्रगति-पत्र</div>
            <div class="session-text">शैक्षणिक सत्र 2025 - 2026</div>
          </div>

          <div class="student-details">
            <div class="field-row">
              <span class="field-label">छात्र / छात्रा का नाम</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <div class="field-half">
                <span class="field-label">अनुक्रमांक</span>
                <span class="field-value" contenteditable="true"></span>
              </div>
              <div class="field-half" style="padding-left: 10px;">
                <span class="field-label">निलय</span>
                <span class="field-value" contenteditable="true"></span>
              </div>
            </div>

            <div class="field-row">
              <span class="field-label">पिता का नाम</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <span class="field-label">माता का नाम</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <div class="field-half">
                <span class="field-label">जन्मतिथि</span>
                <span class="field-value" contenteditable="true"></span>
              </div>
              <div class="field-half" style="padding-left: 10px;">
                <span class="field-label">कक्षा एवं वर्ग</span>
                <span class="field-value" contenteditable="true"></span>
              </div>
            </div>

            <div class="field-row">
              <span class="field-label">मोबाइल नंबर</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <span class="field-label">छात्र पंजीकरण संख्या</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <span class="field-label">आधार संख्या</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <span class="field-label">पेन संख्या (PEN No.)</span>
              <span class="field-value" contenteditable="true"></span>
            </div>

            <div class="field-row">
              <span class="field-label">पता</span>
              <span class="field-value" contenteditable="true"></span>
            </div>
          </div>
        </div>

      </div>
    </div>
  </div>

</body>
</html>
