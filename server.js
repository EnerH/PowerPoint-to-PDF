const express = require("express");
const fileUpload = require("express-fileupload");
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");

const app = express();

// Configuration
const PORT = process.env.PORT || 3000;
const UPLOAD_DIR = path.join(__dirname, "uploads");
const CONVERTED_DIR = path.join(__dirname, "converted");
const LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice";

// Ensure directories exist
[UPLOAD_DIR, CONVERTED_DIR].forEach((dir) => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Security middleware
app.use(helmet());
app.use(
  rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100, // Limit each IP to 100 requests per window
    standardHeaders: true,
    legacyHeaders: false,
  })
);

// File upload configuration
app.use(
  fileUpload({
    limits: { fileSize: 50 * 1024 * 1024 }, // 50MB limit
    abortOnLimit: true,
    useTempFiles: false,
  })
);

// Static files with proper MIME types
app.use(
  express.static(path.join(__dirname, "public"), {
    setHeaders: (res, filePath) => {
      if (filePath.endsWith(".css")) {
        res.setHeader("Content-Type", "text/css");
      }
    },
  })
);

// Conversion endpoint
app.post("/convert", async (req, res) => {
  try {
    if (!req.files?.ppt) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const pptFile = req.files.ppt;
    const fileExt = path.extname(pptFile.name).toLowerCase();

    // Validate file type
    if (![".ppt", ".pptx"].includes(fileExt)) {
      return res.status(400).json({ error: "Invalid file type" });
    }

    // Create unique filenames
    const timestamp = Date.now();
    const uploadFilename = `${timestamp}${fileExt}`;
    const uploadPath = path.join(UPLOAD_DIR, uploadFilename);
    const pdfFilename = `${timestamp}.pdf`;
    const pdfPath = path.join(CONVERTED_DIR, pdfFilename);

    // Save uploaded file
    await pptFile.mv(uploadPath);

    // Silent conversion LibreOffice
    const command = `nohup "${LIBREOFFICE_PATH}" --headless --convert-to pdf --outdir "${outputDir}" "${uploadPath}" > /dev/null 2>&1 &`;

    // Convert using LibreOffice
    await new Promise((resolve, reject) => {
      const command = `"${LIBREOFFICE_PATH}" --headless --convert-to pdf --outdir "${CONVERTED_DIR}" "${uploadPath}"`;

      exec(command, (error, stdout, stderr) => {
        if (error) {
          console.error(`Conversion Error: ${stderr}`);
          return reject(new Error("Conversion failed"));
        }

        if (!fs.existsSync(pdfPath)) {
          return reject(new Error("No output file created"));
        }

        resolve();
      });
    });

    // Stream the converted file
    res.download(pdfPath, pdfFilename, (err) => {
      // Cleanup files regardless of success
      [uploadPath, pdfPath].forEach((file) => {
        try {
          fs.unlinkSync(file);
        } catch (cleanupErr) {
          console.error("Cleanup Error:", cleanupErr);
        }
      });

      if (err) {
        console.error("Download Error:", err);
        if (!res.headersSent) {
          res.status(500).json({ error: "Download failed" });
        }
      }
    });
  } catch (err) {
    console.error("Server Error:", err);
    const statusCode = err.message.includes("Conversion") ? 500 : 400;
    res.status(statusCode).json({ error: err.message });
  }
});

// Handle favicon requests
app.get("/favicon.ico", (req, res) => res.status(204).end());

// Error handling middleware
app.use((err, req, res, next) => {
  console.error("Unhandled Error:", err);
  res.status(500).json({ error: "Internal server error" });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`Uploads directory: ${UPLOAD_DIR}`);
  console.log(`Converted files directory: ${CONVERTED_DIR}`);
});
