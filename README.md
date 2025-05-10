📚 Student Report Card Generator with Charts, Email & Custom UI

Automatically generates personalized student report cards from Google Sheets, inserts performance charts, and emails them as PDFs. Now includes a user-friendly sidebar interface!




✨ Features

📝 Google Sheets Integration – Uses student data directly from a spreadsheet.

📊 Dynamic Charts – Automatically embeds performance charts.

🧠 Rubrics – Converts numeric scores into performance levels (EXCEEDING, MEETING, etc).

🖼️ Branding – School logo and academic details added to each report.

📄 PDF Conversion – Generates downloadable report cards as PDFs.

📧 Email Delivery – Automatically sends report cards to student emails.

🗂️ Organized Drive Folder – Stores all reports in a unique folder.

🖥️ Sidebar UI – Launch and monitor report generation with a clean, embedded sidebar.





🧩 Setup Guide

1. 📊 Prepare the Spreadsheet

Ensure your sheet has columns in this format:

ADM | NAME | ENG_CAT1 | ENG_CAT2 | ... | CRE_CAT1 | CRE_CAT2 | EMAIL

In row 1 (or a metadata sheet), insert:

W1: Term

X1: Year

Y1: Grade



---

2. 🧾 Create a Google Docs Template

Include placeholders like:

{{NAME}}, {{ADM}}, {{ENG}}, {{SCI_CAT2_RUBRIC}}

{{SCHOOL_LOGO}}, {{GRADE}}, {{TERM}}, {{YEAR}}

{{CHART}} for inserting the performance chart



---

3. 🧑‍💻 Add the Script + Sidebar UI

Open Extensions > Apps Script

Paste the Google Apps Script code and the HTML below

Add the sidebar using:


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📄 Report Cards")
    .addItem("Open Sidebar", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Report Card Generator");
  SpreadsheetApp.getUi().showSidebar(html);
}


---

4. 🧾 Sidebar HTML (Sidebar.html)

<!-- Sidebar.html -->
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body {
        font-family: "Roboto", sans-serif;
        padding: 10px;
      }
      button {
        background-color: #1a73e8;
        color: white;
        border: none;
        padding: 10px 20px;
        font-size: 14px;
        border-radius: 5px;
        cursor: pointer;
      }
      button:hover {
        background-color: #155ab6;
      }
    </style>
  </head>
  <body>
    <h3>Generate Report Cards</h3>
    <p>This will create report cards for all learners, with charts and email delivery.</p>
    <button onclick="generate()">Generate Now</button>
    <div id="status" style="margin-top: 10px;"></div>
    <script>
      function generate() {
        document.getElementById("status").innerText = "Processing...";
        google.script.run
          .withSuccessHandler(function() {
            document.getElementById("status").innerText = "✅ Report Cards Generated Successfully!";
          })
          .withFailureHandler(function(error) {
            document.getElementById("status").innerText = "❌ Error: " + error.message;
          })
          .generateReportCardsWithCharts();
      }
    </script>
  </body>
</html>


---

🧠 Rubric Logic (Sample)


---

📤 Sample Email Output

> Subject: [Student Name] - Report Card Term [X]

Body: Please find attached the report card for Term X, Year Y.
Includes downloadable PDF & chart.




---

🛠 Technologies Used

Google Apps Script

Google Docs API

Google Drive API

Google Gmail API

HTML/CSS/JavaScript (Sidebar UI)



Here is my google doc template:

{{SCHOOL_LOGO}}

# Student Report Card for Grade 8

**Student Name:** {{NAME}}  
**Admission Number:** {{ADM}}

---

### Subjects and Scores

| Learning Areas | Cat1 Scores | Cat1 Rubrics | Cat2 Scores | Cat2 Rubrics | Average Scores | Rubric Levels |
|----------------|-------------|--------------|-------------|--------------|----------------|----------------|
| English        | {{ENG_CAT1}} | {{ENG_CAT1_RUBRIC}} | {{ENG_CAT2}} | {{ENG_CAT2_RUBRIC}} | {{ENG}} | {{ENG_RUBRIC}} |
| Kiswahili      | {{KIS_CAT1}} | {{KIS_CAT1_RUBRIC}} | {{KIS_CAT2}} | {{KIS_CAT2_RUBRIC}} | {{KIS}} | {{KIS_RUBRIC}} |
| Maths          | {{MATH_CAT1}} | {{MATH_CAT1_RUBRIC}} | {{MATH_CAT2}} | {{MATH_CAT2_RUBRIC}} | {{MATH}} | {{MATH_RUBRIC}} |
| Int Scie       | {{SCI_CAT1}} | {{SCI_CAT1_RUBRIC}} | {{SCI_CAT2}} | {{SCI_CAT2_RUBRIC}} | {{SCI}} | {{SCI_RUBRIC}} |
| SST            | {{SST_CAT1}} | {{SST_CAT1_RUBRIC}} | {{SST_CAT2}} | {{SST_CAT2_RUBRIC}} | {{SST}} | {{SST_RUBRIC}} |
| C.Arts & S     | {{ART_CAT1}} | {{ART_CAT1_RUBRIC}} | {{ART_CAT2}} | {{ART_CAT2_RUBRIC}} | {{ART}} | {{ART_RUBRIC}} |
| Pretech        | {{PRETECH_CAT1}} | {{PRETECH_CAT1_RUBRIC}} | {{PRETECH_CAT2}} | {{PRETECH_CAT2_RUBRIC}} | {{PRETECH}} | {{PRETECH_RUBRIC}} |
| Agrics         | {{AGRIC_CAT1}} | {{AGRIC_CAT1_RUBRIC}} | {{AGRIC_CAT2}} | {{AGRIC_CAT2_RUBRIC}} | {{AGRIC}} | {{AGRIC_RUBRIC}} |
| C.R.E          | {{CRE_CAT1}} | {{CRE_CAT1_RUBRIC}} | {{CRE_CAT2}} | {{CRE_CAT2_RUBRIC}} | {{CRE}} | {{CRE_RUBRIC}} |

---

### Totals and Averages

| Category        | Total         | Mean           | Rubric Level         |
|-----------------|---------------|----------------|-----------------------|
| CAT 1 SCORES    | {{TOTAL_CAT1}} | {{MEAN_CAT1}} | {{MEAN_CAT1_RUBRIC}} |
| CAT 2 SCORES    | {{TOTAL_CAT2}} | {{MEAN_CAT2}} | {{MEAN_CAT2_RUBRIC}} |
| AVERAGE SCORES  | {{TOTAL_AVG}}  | {{MEAN_AVG}}  | {{MEAN_AVG_RUBRIC}}  |

---

**Class Teacher's Comment:** {{AVG_COMMENT}}  

**H.O.I's Comment:** ………………………………………………………………………………………………………………..

**H.O.I's Signature:** ____________________  
**Date:** ………………………………………………

**Closing Date:** 4th of April, 2025  
**Opening Date:** 28th of April, 2025


📬 Contact

Created by Tonny Wanjala Mulati
📧 tonnymulati79@gmail.com
📹 YouTube Channel
💼 LinkedIn




