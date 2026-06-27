const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json());

// Upload folder
const upload = multer({
    dest: "uploads/"
});

// Upload Excel and Convert to JSON
app.post("/upload", upload.single("excel"), (req, res) => {
    try {

        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: "No file uploaded"
            });
        }

        // Read Excel
        const workbook = XLSX.readFile(req.file.path);

        const result = {};

        // Read all sheets
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];

            // result[sheetName] = XLSX.utils.sheet_to_json(
            //     worksheet,
            //     {
            //         defval: "",       // Empty cell => ""
            //         raw: false
            //     }
            // );
            // Convert sheet to 2D array
            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: "",
                raw: false
            });

            // First row = title (ignore)
            // Second row = headers
            const headers = rows[1];

            // Third row onwards = data
            const data = rows.slice(2);

            result[sheetName] = data.map(row => {
                const obj = {};

                headers.forEach((header, index) => {
                    obj[header] = row[index] ?? "";
                });

                return obj;
            });

        });

        // Delete uploaded file
        fs.unlinkSync(req.file.path);

        console.log(result);
        res.json({
            success: true,
            sheets: workbook.SheetNames,
            data: result
        });

    } catch (err) {

        console.log(err);

        res.status(500).json({
            success: false,
            message: err.message
        });

    }
});

app.listen(5000, () => {
    console.log("Server running on port 5000");
});