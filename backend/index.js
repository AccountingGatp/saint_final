import "dotenv/config";
import express from "express";
import cors from "cors";
import multer from "multer";
import { buildWorkbookBuffer, buildImportBuffer } from "./excel.js";
import { uploadBuffer, downloadFile, getDownloadUrl } from "./storage/b2.js";

const app = express();
const PORT = process.env.PORT || 4000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
});

app.use(cors());
app.use(express.json());

app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res
        .status(400)
        .json({ error: "No file uploaded (expected field 'file')." });
    }

    const date = req.body.date || "06/26";
    const currency = req.body.currency || "AUD";

    const outBuffer = await buildWorkbookBuffer(req.file.buffer, {
      date,
      currency,
    });

    const base = (req.file.originalname || "workbook").replace(/\.[^.]+$/, "");
    const outName = `${base}-output.xlsx`;

    const stored = await uploadBuffer(outBuffer, {
      type: "output",
      fileName: outName,
    });

    return res.json({
      fileId: stored.fileId,
      fileName: stored.fileName,
      downloadUrl: stored.downloadUrl,
    });
  } catch (err) {
    console.error(err);
    return res
      .status(500)
      .json({ error: err.message || "Failed to process file." });
  }
});

app.post("/import", async (req, res) => {
  try {
    const { fileId } = req.body;

    if (!fileId) {
      return res.status(400).json({ error: "fileId is required." });
    }

    const { buffer, fileName } = await downloadFile("output", fileId);
    const outBuffer = buildImportBuffer(buffer);

    const base = fileName.replace(/\.[^.]+$/, "").replace(/-output$/, "");
    const outName = `${base}-import.xlsx`;

    const stored = await uploadBuffer(outBuffer, {
      type: "import",
      fileName: outName,
    });

    return res.json({
      fileId: stored.fileId,
      fileName: stored.fileName,
      downloadUrl: stored.downloadUrl,
    });
  } catch (err) {
    console.error(err);
    return res
      .status(500)
      .json({ error: err.message || "Failed to build import file." });
  }
});

app.get("/files/:fileId/download", async (req, res) => {
  try {
    const { fileId } = req.params;
    const type = req.query.type === "import" ? "import" : "output";
    const downloadUrl = await getDownloadUrl(type, fileId);

    return res.json({ downloadUrl });
  } catch (err) {
    console.error(err);
    return res.status(404).json({ error: "File not found." });
  }
});

app.listen(PORT, () => {
  console.log(`Backend listening on http://localhost:${PORT}`);
});
