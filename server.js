const express = require("express");
const fileUpload = require("express-fileupload");
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");

const app = express();

// macOS-specific LibreOffice path
const LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice";

app.use(fileUpload());
app.use(express.static("public"));

app.post("/convert", async (req, res) => {
  if (!req.files?.ppt) {
    return res.status(400).send("No file uploaded");
  }

  const pptFile = req.files.ppt;
  const uploadPath = path.join(__dirname, "uploads", pptFile.name);
  const outputDir = path.join(__dirname, "converted");

  try {
    await pptFile.mv(uploadPath);

    // Modified conversion command for macOS
    const command = `"${LIBREOFFICE_PATH}" --headless --convert-to pdf --outdir "${outputDir}" "${uploadPath}"`;

    exec(command, (error, stdout, stderr) => {
      if (error) {
        console.error("Conversion Error:", error);
        console.error("stderr:", stderr);
        return res.status(500).send("Conversion failed");
      }

      const pdfFilename =
        path.basename(uploadPath, path.extname(uploadPath)) + ".pdf";
      const pdfPath = path.join(outputDir, pdfFilename);

      res.download(pdfPath, (err) => {
        if (err) console.error("Download Error:", err);
        // Cleanup with error handling
        try {
          fs.unlinkSync(uploadPath);
          fs.unlinkSync(pdfPath);
        } catch (cleanupErr) {
          console.error("Cleanup Error:", cleanupErr);
        }
      });
    });
  } catch (err) {
    console.error("Upload Error:", err);
    res.status(500).send(err.message);
  }
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
