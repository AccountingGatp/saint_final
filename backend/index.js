import express from "express";
import cors from "cors";
import multer from "multer";
import { buildWorkbookBuffer, buildImportBuffer } from "./excel.js";

const app = express();
const PORT = process.env.PORT || 4000;

// Keep uploads in memory only — nothing is written to disk.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // 25 MB
});

app.use(cors());

app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded (expected field 'file')." });
    }

    // Optional overrides from the form; fall back to the script defaults.
    const date = req.body.date || "06/26";
    const currency = req.body.currency || "AUD";

    const outBuffer = await buildWorkbookBuffer(req.file.buffer, { date, currency });

    const base = (req.file.originalname || "workbook").replace(/\.[^.]+$/, "");
    const outName = `${base}-output.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    return res.send(outBuffer);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to process file." });
  }
});

// Turn a previously generated output workbook into an accounting import workbook.
app.post("/import", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded (expected field 'file')." });
    }

    const outBuffer = buildImportBuffer(req.file.buffer);

    const base = (req.file.originalname || "workbook")
      .replace(/\.[^.]+$/, "")
      .replace(/-output$/, "");
    const outName = `${base}-import.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    return res.send(outBuffer);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || "Failed to build import file." });
  }
});

app.listen(PORT, () => {
  console.log(`Backend listening on http://localhost:${PORT}`);
});
