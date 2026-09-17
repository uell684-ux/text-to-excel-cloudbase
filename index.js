const path = require("path");
const express = require("express");
const cors = require("cors");
const morgan = require("morgan");
const { init: initDB, Counter } = require("./db");
const { runDifyWorkflow } = require("./dify");
const { generateExcel } = require("./excel");

const logger = morgan("tiny");

const app = express();
app.use(express.urlencoded({ extended: false }));
app.use(express.json());
app.use(cors());
app.use(logger);

// 首页
app.get("/", async (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// 更新计数
app.post("/api/count", async (req, res) => {
  const { action } = req.body;
  if (action === "inc") {
    await Counter.create();
  } else if (action === "clear") {
    await Counter.destroy({
      truncate: true,
    });
  }
  res.send({
    code: 0,
    data: await Counter.count(),
  });
});

// 获取计数
app.get("/api/count", async (req, res) => {
  const result = await Counter.count();
  res.send({
    code: 0,
    data: result,
  });
});

// 小程序调用，获取微信 Open ID
app.get("/api/wx_openid", async (req, res) => {
  if (req.headers["x-wx-source"]) {
    res.send(req.headers["x-wx-openid"]);
  }
});


// Text to Excel - call Dify workflow
app.post("/api/convert", async (req, res) => {
  try {
    const { text } = req.body;

    if (!text || !text.trim()) {
      return res.status(400).send({
        success: false,
        error: "text is required"
      });
    }

    const result = await runDifyWorkflow(text);

    res.send({
      success: true,
      result: result
    });
  } catch (error) {
    console.error("Dify workflow error:", error);

    res.status(500).send({
      success: false,
      error: error.message || "Dify workflow failed"
    });
  }
});


// Generate Excel from unstructured text
app.post("/api/excel", async (req, res) => {
  try {
    const { text } = req.body;

    if (!text || !text.trim()) {
      return res.status(400).send({
        success: false,
        error: "text is required"
      });
    }

    // 1. Send original text to Dify
    const result = await runDifyWorkflow(text);

    // 2. Convert Dify JSON string to JavaScript object
    let data;

    if (typeof result === "string") {
      data = JSON.parse(result);
    } else {
      data = result;
    }

    if (data.success === false) {
      throw new Error(data.error || "Dify processing failed");
    }

    // 3. Generate Excel
    const buffer = await generateExcel(data);

    // 4. Return .xlsx directly
    const filename = "text-to-excel.xlsx";

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${filename}"`
    );

    res.send(Buffer.from(buffer));

  } catch (error) {
    console.error("Excel generation error:", error);

    res.status(500).send({
      success: false,
      error: error.message || "Excel generation failed"
    });
  }
});

const port = process.env.PORT || 80;

async function bootstrap() {
  await initDB();
  app.listen(port, () => {
    console.log("启动成功", port);
  });
}

bootstrap();
