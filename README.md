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




📬 Contact

Created by Tonny Wanjala Mulati
📧 tonnymulati79@gmail.com
📹 YouTube Channel
💼 LinkedIn




