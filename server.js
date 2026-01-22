import dotenv from "dotenv";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import os from "os";
import crypto from "crypto";
import OpenAI from "openai";
import multer from "multer";
import XLSX from "xlsx";
import mammoth from "mammoth";
import PptxParser from "node-pptx-parser";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const pdfParse = require("pdf-parse");

// 只在本機存在 .env 時才載入（雲端 Railway 會用環境變數，不會有 .env）
if (fs.existsSync("./.env")) {
  dotenv.config({ path: "./.env" });
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// ✅ 保留 JSON（如果未來還想支援 application/json）
app.use(express.json({ limit: "16mb" }));
app.use(express.static(path.join(__dirname, "public")));

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ✅ 上傳（不落地，直接 memory）
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 15 * 1024 * 1024, // 單檔 15MB（你可調）
    files: 20, // 兩欄各最多 10 檔，所以總數抓 20 比較安全
  },
});

function safe(v, fallback = "") {
  const s = (v ?? "").toString().trim();
  return s.length ? s : fallback;
}

function buildInputBundle(contextDocuments, rubricInput) {
  const ctx = safe(contextDocuments);
  const rb = safe(rubricInput);
  if (!rb) return ctx;

  return [
    ctx,
    "",
    "────────────────────",
    "【補充：評分規則／rubric（輸入資料的一部分）】",
    rb,
  ].join("\n");
}

// 從外部檔案讀取模板（你那一串 prompt 原封不動貼在 template.txt）
function loadTemplate() {
  const templatePath = path.join(__dirname, "template.txt");
  if (!fs.existsSync(templatePath)) {
    throw new Error(
      "找不到 template.txt。請在專案根目錄建立 template.txt，並把你的完整 prompt 原封不動貼進去。"
    );
  }
  return fs.readFileSync(templatePath, "utf-8");
}

function buildPrompt(template, payload) {
  return template
    .replaceAll("{{CONTEXT_DOCUMENTS}}", payload.contextDocuments)
    .replaceAll("{{TRAINING_GOAL}}", payload.trainingGoal)
    .replaceAll("{{TRAINEE_ROLE}}", payload.traineeRole)
    .replaceAll("{{PERSONA_ROLE}}", payload.personaRole)
    .replaceAll("{{SCENARIO_TYPE}}", payload.scenarioType)
    .replaceAll("{{COACH_ROLE}}", payload.coachRole)
    .replaceAll("{{CUSTOMER_NAME}}", payload.customerName);
}

function getExt(filename = "") {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i + 1).toLowerCase() : "";
}

function excelBufferToText(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const lines = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws);
    lines.push(`--- Sheet: ${sheetName} ---\n${csv}`);
  }
  return lines.join("\n\n");
}

// ✅ 避免大檔塞爆 prompt：做簡單截斷（可調）
function clampText(s, maxChars = 80000) {
  const text = (s ?? "").toString();
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars) + `\n\n(…已截斷，原長度 ${text.length} chars)`;
}

function isSupportedTextExt(ext) {
  return ext === "txt" || ext === "csv" || ext === "md" || ext === "json";
}
function isSupportedExcelExt(ext) {
  return ext === "xlsx" || ext === "xls";
}
function isSupportedImageExt(ext) {
  return ext === "png" || ext === "jpg" || ext === "jpeg" || ext === "webp";
}
function isSupportedPdfExt(ext) {
  return ext === "pdf";
}
function isSupportedDocxExt(ext) {
  return ext === "docx";
}
function isSupportedPptxExt(ext) {
  return ext === "pptx";
}

// ✅ PPTX parser 需要 file path：把 buffer 暫存到臨時檔後解析，再刪除
async function pptxBufferToText(buffer) {
  const tmpDir = os.tmpdir();
  const tmpName = `upload-${crypto.randomUUID()}.pptx`;
  const tmpPath = path.join(tmpDir, tmpName);

  await fs.promises.writeFile(tmpPath, buffer);
  try {
    const parser = new PptxParser(tmpPath);
    const slides = await parser.extractText();

    const lines = [];
    slides.forEach((slide) => {
      const id = slide.id ?? "unknown";
      const text = Array.isArray(slide.text) ? slide.text.join("\n") : "";
      lines.push(`--- Slide: ${id} ---\n${text}`);
    });

    return lines.join("\n\n");
  } finally {
    fs.promises.unlink(tmpPath).catch(() => {});
  }
}

function fileListSummary(files = []) {
  return files.map((f) => ({ name: f.originalname, size: f.size }));
}

/**
 * ✅ /api/generate：支援 multipart/form-data（文字 + 兩欄檔案）
 * - text：逐字稿文字（必填其一：文字或檔案）
 * - rubricInput：評分表文字（選填）
 * - contextFiles：逐字稿檔案（選填，可含圖片）
 * - rubricFiles：評分表檔案（選填，不接受圖片）
 */
app.post(
  "/api/generate",
  upload.fields([
    { name: "contextFiles", maxCount: 10 },
    { name: "rubricFiles", maxCount: 10 },
  ]),
  async (req, res) => {
    const body = req.body || {};
    const contextFiles = (req.files?.contextFiles || []);
    const rubricFiles = (req.files?.rubricFiles || []);

    console.log("POST /api/generate", {
      hasContextText: !!safe(body.text || body.contextDocuments),
      hasRubricText: !!safe(body.rubricInput),
      contextFiles: fileListSummary(contextFiles),
      rubricFiles: fileListSummary(rubricFiles),
    });

    try {
      // 逐字稿欄文字（必填其一：文字或檔案）
      const baseContext = safe(body.contextDocuments) || safe(body.text);

      // 評分表欄文字（選填）
      const baseRubric = safe(body.rubricInput);

      // ---- 解析逐字稿欄檔案：文字/表格/文件 → text，圖片 → vision
      const contextTextChunks = [];
      const imageInputs = [];

      for (const f of contextFiles) {
        const ext = getExt(f.originalname);

        // 文字類
        if (isSupportedTextExt(ext)) {
          let content = f.buffer.toString("utf8");
          if (ext === "json") {
            try {
              const obj = JSON.parse(content);
              content = JSON.stringify(obj, null, 2);
            } catch {}
          }
          contextTextChunks.push(`[逐字稿附件文字：${f.originalname}]\n${clampText(content)}`);
          continue;
        }

        // Excel
        if (isSupportedExcelExt(ext)) {
          const csvText = excelBufferToText(f.buffer);
          contextTextChunks.push(`[逐字稿附件Excel：${f.originalname}]\n${clampText(csvText)}`);
          continue;
        }

        // PDF（掃描檔提示）
        if (isSupportedPdfExt(ext)) {
          const data = await pdfParse(f.buffer);
          const pdfText = (data?.text || "").trim();
          contextTextChunks.push(
            `[逐字稿附件PDF：${f.originalname}]\n${
              pdfText ? clampText(pdfText) : "（此 PDF 可能為掃描檔，未能解析出文字）"
            }`
          );
          continue;
        }

        // DOCX
        if (isSupportedDocxExt(ext)) {
          const result = await mammoth.extractRawText({ buffer: f.buffer });
          const docText = (result?.value || "").trim();
          contextTextChunks.push(
            `[逐字稿附件Word：${f.originalname}]\n${
              docText ? clampText(docText) : "（此 Word 檔未能解析出文字）"
            }`
          );
          continue;
        }

        // PPTX
        if (isSupportedPptxExt(ext)) {
          const pptText = (await pptxBufferToText(f.buffer)) || "";
          contextTextChunks.push(
            `[逐字稿附件PPTX：${f.originalname}]\n${
              pptText.trim() ? clampText(pptText) : "（此 PPTX 未能解析出文字）"
            }`
          );
          continue;
        }

        // 圖片（vision）— 只允許逐字稿欄
        if (isSupportedImageExt(ext)) {
          const mime =
            ext === "png" ? "image/png" : ext === "webp" ? "image/webp" : "image/jpeg";
          const base64 = f.buffer.toString("base64");
          const dataUrl = `data:${mime};base64,${base64}`;
          imageInputs.push({ type: "input_image", image_url: dataUrl });
          continue;
        }

        return res.status(400).json({
          ok: false,
          error: `逐字稿欄不支援的檔案格式：${f.originalname}`,
        });
      }

      // ---- 解析評分表欄檔案：全部轉文字（不接受圖片）
      const rubricTextChunks = [];

      for (const f of rubricFiles) {
        const ext = getExt(f.originalname);

        if (isSupportedImageExt(ext)) {
          return res.status(400).json({
            ok: false,
            error: `評分表欄不接受圖片檔：${f.originalname}`,
          });
        }

        if (isSupportedTextExt(ext)) {
          let content = f.buffer.toString("utf8");
          if (ext === "json") {
            try {
              const obj = JSON.parse(content);
              content = JSON.stringify(obj, null, 2);
            } catch {}
          }
          rubricTextChunks.push(`[評分表附件文字：${f.originalname}]\n${clampText(content)}`);
          continue;
        }

        if (isSupportedExcelExt(ext)) {
          const csvText = excelBufferToText(f.buffer);
          rubricTextChunks.push(`[評分表附件Excel：${f.originalname}]\n${clampText(csvText)}`);
          continue;
        }

        if (isSupportedPdfExt(ext)) {
          const data = await pdfParse(f.buffer);
          const pdfText = (data?.text || "").trim();
          rubricTextChunks.push(
            `[評分表附件PDF：${f.originalname}]\n${
              pdfText ? clampText(pdfText) : "（此 PDF 可能為掃描檔，未能解析出文字）"
            }`
          );
          continue;
        }

        if (isSupportedDocxExt(ext)) {
          const result = await mammoth.extractRawText({ buffer: f.buffer });
          const docText = (result?.value || "").trim();
          rubricTextChunks.push(
            `[評分表附件Word：${f.originalname}]\n${
              docText ? clampText(docText) : "（此 Word 檔未能解析出文字）"
            }`
          );
          continue;
        }

        if (isSupportedPptxExt(ext)) {
          const pptText = (await pptxBufferToText(f.buffer)) || "";
          rubricTextChunks.push(
            `[評分表附件PPTX：${f.originalname}]\n${
              pptText.trim() ? clampText(pptText) : "（此 PPTX 未能解析出文字）"
            }`
          );
          continue;
        }

        return res.status(400).json({
          ok: false,
          error: `評分表欄不支援的檔案格式：${f.originalname}`,
        });
      }

      // ---- 合併逐字稿欄：文字 + 檔案文字
      const mergedContext = [
        baseContext,
        contextTextChunks.length
          ? "\n\n【逐字稿欄附件彙整】\n" + contextTextChunks.join("\n\n")
          : "",
      ].join("");

      // ---- 合併評分表欄：文字 + 檔案文字
      const mergedRubric = [
        baseRubric,
        rubricTextChunks.length
          ? "\n\n【評分表欄附件彙整】\n" + rubricTextChunks.join("\n\n")
          : "",
      ].join("");

      // 防呆：逐字稿欄一定要有「文字或檔案或圖片」
      if (!safe(mergedContext) && imageInputs.length === 0) {
        return res.status(400).json({
          ok: false,
          error: "缺少逐字稿內容：請輸入逐字稿文字或上傳逐字稿檔案。",
        });
      }

      // 把 rubric 附加進 context（維持你原本習慣）
      const contextBundle = buildInputBundle(mergedContext, mergedRubric);

      // 你原本的 payload 邏輯（不動）
      const payload = {
        contextDocuments: safe(contextBundle, "（此處貼上逐字稿或檢核點）"),
        trainingGoal: "（未提供：請 AI 自行判斷並在輸出中明確定義訓練目標）",
        traineeRole: "（未提供：請 AI 自行判斷學員角色）",
        personaRole: "（未提供：請 AI 自行判斷對練角色類型）",
        scenarioType: "（未提供：請 AI 自行判斷並在輸出中明確定義情境類型）",
        coachRole: "（未提供）",
        customerName: "（未提供）",
      };

      const template = loadTemplate();
      const prompt = buildPrompt(template, payload);

      const model = process.env.OPENAI_MODEL || "gpt-5-mini";

      const response = await client.responses.create({
        model,
        instructions:
          "你是嚴格遵守格式與禁止事項的情境腳本 Prompt 工程師。只輸出使用者要求的內容，不要多說。",
        input: [
          {
            role: "user",
            content: [{ type: "input_text", text: prompt }, ...imageInputs],
          },
        ],
      });

      res.json({ ok: true, output: response.output_text || "" });
    } catch (err) {
      console.error("500 error:", err);
      res.status(500).json({ ok: false, error: err?.message || "Unknown error" });
    }
  }
);

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`http://localhost:${port}`));
