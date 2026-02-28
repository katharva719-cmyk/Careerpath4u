// server.js
const express = require("express");
const bodyParser = require("body-parser");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(bodyParser.json());
app.use(express.static(__dirname)); // serve HTML

const excelFile = path.join(__dirname, "empdata.xlsx");

app.post("/submit-feedback", (req, res) => {
  try {
    const { name, email, rating, comments } = req.body;
    console.log("📩 New Feedback Received:", req.body);

    let workbook;
    if (fs.existsSync(excelFile)) {
      workbook = xlsx.readFile(excelFile);
    } else {
      workbook = xlsx.utils.book_new();
    }

    const sheetName = "FeedbackData";
    let worksheet = workbook.Sheets[sheetName];
    let data = worksheet ? xlsx.utils.sheet_to_json(worksheet) : [];

    // Add new row
    data.push({
      Name: name,
      Email: email,
      Rating: rating,
      Comments: comments,
    });

    // Always include headers
    worksheet = xlsx.utils.json_to_sheet(data, { header: ["Name", "Email", "Rating", "Comments"] });
    workbook.Sheets[sheetName] = worksheet;

    if (!workbook.SheetNames.includes(sheetName)) {
      workbook.SheetNames.push(sheetName);
    }

    xlsx.writeFile(workbook, excelFile);

    console.log("✅ Feedback saved to Excel!");
    res.json({ success: true });
  } catch (err) {
    console.error("❌ Error saving feedback:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(3000, () => {
  console.log("🚀 Server running at http://localhost:3000");
});
