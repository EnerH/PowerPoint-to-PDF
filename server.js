const express = require("express");
const fileUpload = require("express-fileupload");
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const archiver = require("archiver");

const app = express();

// macOS-specific LibreOffice path
const LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice";

app.use(fileUpload());
app.use(express.static("public"));

app.post("/convert", async (req, res) => {
  if (!req.files?.ppt) {
    return res.status(400).send("No files uploaded");
  }

  const pptFiles = Array.isArray(req.files.ppt)
    ? req.files.ppt
    : [req.files.ppt];
  const uploadDir = path.join(__dirname, "uploads");
  const outputDir = path.join(__dirname, "converted");
  let uploadedFilesInfo = [];

  try {
    // Create directories if they don't exist
    fs.mkdirSync(uploadDir, { recursive: true });
    fs.mkdirSync(outputDir, { recursive: true });

    // Move all files to upload directory with unique names
    for (const pptFile of pptFiles) {
      const uniqueName =
        Date.now() +
        "-" +
        Math.random().toString(36).substr(2, 9) +
        "-" +
        pptFile.name;
      const uploadPath = path.join(uploadDir, uniqueName);
      await pptFile.mv(uploadPath);
      uploadedFilesInfo.push({
        originalName: pptFile.name,
        uploadPath: uploadPath,
        pdfName:
          path.basename(pptFile.name, path.extname(pptFile.name)) + ".pdf",
        uniquePdfName:
          path.basename(uniqueName, path.extname(uniqueName)) + ".pdf",
      });
    }

    // Convert all files
    const command = `"${LIBREOFFICE_PATH}" --headless --convert-to pdf --outdir "${outputDir}" ${uploadedFilesInfo
      .map((f) => `"${f.uploadPath}"`)
      .join(" ")}`;
    await new Promise((resolve, reject) => {
      exec(command, (error, stdout, stderr) => {
        if (error) {
          console.error("Conversion Error:", error);
          console.error("stderr:", stderr);
          reject(new Error("Conversion failed"));
          return;
        }
        resolve();
      });
    });

    // Create zip archive
    const archive = archiver("zip");
    res.attachment("converted.zip");
    archive.pipe(res);

    // Add PDFs to archive with original names
    for (const fileInfo of uploadedFilesInfo) {
      const pdfPath = path.join(outputDir, fileInfo.uniquePdfName);
      if (fs.existsSync(pdfPath)) {
        archive.file(pdfPath, { name: fileInfo.pdfName });
      }
    }

    // Finalize and send
    await archive.finalize();

    // Cleanup after response
    res.on("finish", () => {
      uploadedFilesInfo.forEach((info) => {
        try {
          fs.unlinkSync(info.uploadPath);
          fs.unlinkSync(path.join(outputDir, info.uniquePdfName));
        } catch (err) {
          console.error("Cleanup error:", err);
        }
      });
    });
  } catch (err) {
    // Cleanup uploaded files on error
    uploadedFilesInfo.forEach((info) => {
      try {
        fs.unlinkSync(info.uploadPath);
      } catch (e) {}
    });
    console.error("Error:", err);
    res.status(500).send(err.message);
  }
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
