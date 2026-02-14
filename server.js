require('dotenv').config();
const express = require("express");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3889;
const HOST = process.env.HOST || "0.0.0.0";

// Serve static files from dist directory
app.use(express.static(path.join(__dirname, "dist")));

// API endpoint to get configuration
app.get("/api/config", (req, res) => {
  const modelType = process.env.MODEL_TYPE || "cloud";
  
  let config = {
    modelType: modelType,
    marginTop: parseFloat(process.env.MARGIN_TOP) || 2.54,
    marginBottom: parseFloat(process.env.MARGIN_BOTTOM) || 2.54,
    marginLeft: parseFloat(process.env.MARGIN_LEFT) || 3.18,
    marginRight: parseFloat(process.env.MARGIN_RIGHT) || 3.18,
    lineSpacing: parseFloat(process.env.LINE_SPACING) || 28,
    titleFont: process.env.TITLE_FONT || "黑体",
    titleFontSize: parseFloat(process.env.TITLE_FONT_SIZE) || 16,
    h1Font: process.env.H1_FONT || "黑体",
    h1FontSize: parseFloat(process.env.H1_FONT_SIZE) || 16,
    h2Font: process.env.H2_FONT || "黑体",
    h2FontSize: parseFloat(process.env.H2_FONT_SIZE) || 14,
    h3Font: process.env.H3_FONT || "楷体",
    h3FontSize: parseFloat(process.env.H3_FONT_SIZE) || 12,
    bracketFont: process.env.BRACKET_FONT || "仿宋",
    bracketFontSize: parseFloat(process.env.BRACKET_FONT_SIZE) || 12
  };
  
  if (modelType === "cloud") {
    config.apiUrl = process.env.API_URL || "https://api.deepseek.com/chat/completions";
    config.apiKey = process.env.API_KEY || "";
    config.model = process.env.API_MODEL || "deepseek-chat";
  } else if (modelType === "local") {
    config.apiUrl = process.env.LOCAL_API_URL || "http://localhost:11434/api/chat";
    config.apiKey = process.env.LOCAL_API_KEY || "";
    config.model = process.env.LOCAL_MODEL_NAME || "llama2";
    config.localNeedAuth = process.env.LOCAL_API_NEED_AUTH === "true";
  }
  
  res.json(config);
});

// API endpoint to switch model type
app.post("/api/switch-model", express.json(), (req, res) => {
  const { modelType, apiUrl, apiKey, model } = req.body;
  
  if (!modelType || (modelType !== "cloud" && modelType !== "local")) {
    return res.status(400).json({ error: "Invalid model type" });
  }
  
  process.env.MODEL_TYPE = modelType;
  
  if (modelType === "cloud") {
    if (apiUrl) process.env.API_URL = apiUrl;
    if (apiKey) process.env.API_KEY = apiKey;
    if (model) process.env.API_MODEL = model;
  } else if (modelType === "local") {
    if (apiUrl) process.env.LOCAL_API_URL = apiUrl;
    if (apiKey) process.env.LOCAL_API_KEY = apiKey;
    if (model) process.env.LOCAL_MODEL_NAME = model;
  }
  
  res.json({ success: true, message: "Model switched successfully" });
});

app.listen(PORT, HOST, () => {
  const modelType = process.env.MODEL_TYPE || "cloud";
  const modelInfo = modelType === "cloud" 
    ? `Cloud Model: ${process.env.API_MODEL || "deepseek-chat"}`
    : `Local Model: ${process.env.LOCAL_MODEL_NAME || "llama2"} @ ${process.env.LOCAL_API_URL || "localhost:11434"}`;
  
  console.log("");
  console.log("========================================");
  console.log("  WpsAgent Started!");
  console.log("========================================");
  console.log("");
  console.log("  Local: http://localhost:" + PORT);
  console.log("  Network: http://" + HOST + ":" + PORT);
  console.log("");
  console.log("  Model: " + modelInfo);
  console.log("");
  console.log("  Press Ctrl+C to stop");
  console.log("========================================");
  console.log("");
});

process.on("SIGINT", () => {
  console.log("\nShutting down...");
  process.exit(0);
});
