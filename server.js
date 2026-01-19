import dotenv from "dotenv";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import OpenAI from "openai";

// 只在本機存在 .env 時才載入（雲端 Railway 會用環境變數，不會有 .env）
if (fs.existsSync("./.env")) {
  dotenv.config({ path: "./.env" });
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: "4mb" }));
app.use(express.static(path.join(__dirname, "public")));

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

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

app.post("/api/generate", async (req, res) => {
  try {
    const body = req.body || {};

    // 必填：逐字稿 / 檢核點
    if (!safe(body.contextDocuments)) {
      return res.status(400).json({
        ok: false,
        error: "缺少 contextDocuments（逐字稿／檢核點）。",
      });
    }

    const contextBundle = buildInputBundle(body.contextDocuments, body.rubricInput);

    // 使用者只輸入「逐字稿/檢核點」(必填) + 「評分表」(選填)
    // 其他欄位全部交由 AI 自行判斷
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

    // gpt-5 不支援 temperature，所以不要放 temperature
    const response = await client.responses.create({
      model: process.env.OPENAI_MODEL || "gpt-5",
      instructions:
        "你是嚴格遵守格式與禁止事項的情境腳本 Prompt 工程師。只輸出使用者要求的內容，不要多說。額外硬性要求：輸出的 Persona Prompt 必須包含『主題鎖定（Topic Lock）＋完成條件（Completion Gate）＋衍生追問階梯（Follow-up Ladder）＋禁止跳題（No Jumping）』的規則，確保不會從 A 話題未完成就跳到 B。",
      input: prompt,
    });

    res.json({ ok: true, output: response.output_text || "" });
  } catch (err) {
    res.status(500).json({ ok: false, error: err?.message || "Unknown error" });
  }
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`http://localhost:${port}`));
