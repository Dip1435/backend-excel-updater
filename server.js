const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(cors());
app.use(express.static("public")); // Serve frontend

// Upload Route
app.post("/upload", upload.single("excel"), async (req, res) => {
  try {
    const filePath = req.file.path;

    // Read Excel File
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    let data = XLSX.utils.sheet_to_json(sheet);

    // Add new columns with formulas
    data = data.map((row) => {
      let bmi = "";
      if (row.WT && row.HT) {
        bmi = (row.WT / row.HT / row.HT) * 10000;
      }
      return {
        ...row,
        BMI: bmi?.toFixed(2),
        BPS: "",
        Status: "",
        Pulses: "",
      };
    });

    // Create new sheet
    const newSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Updated Data");

    // Save updated Excel file (temporary location)
    const outputFilePath = path.join(__dirname, "updated_data.xlsx");
    XLSX.writeFile(newWorkbook, outputFilePath);

    res.download(outputFilePath, "updated_data.xlsx", () => {
      fs.unlinkSync(filePath); // Cleanup uploaded file
      fs.unlinkSync(outputFilePath); // Cleanup generated file
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: "Internal Server Error" });
  }
});

// Start Server
const PORT = process.env.PORT || 5000; // Use dynamic port for Render
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
