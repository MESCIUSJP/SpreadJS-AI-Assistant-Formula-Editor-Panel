import express from "express";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;
app.use(express.json());
app.use(express.static("public"));

const ai = new GoogleGenAI({ apiKey: process.env.AI_API_KEY });
const defaultModel = process.env.AI_MODEL || "gemini-2.5-flash";

const convertMessages = (messages) => {
  const systemParts = [];
  const contents = [];
  for (const msg of messages) {
    if (msg.role === "system") {
      systemParts.push(msg.content);
    } else {
      contents.push({
        role: msg.role === "assistant" ? "model" : "user",
        parts: [{ text: msg.content }],
      });
    }
  }
  if (systemParts.length > 0 && contents.length > 0 && contents[0].role === "user") {
    contents[0].parts[0].text = systemParts.join("\n\n") + "\n\n" + contents[0].parts[0].text;
  }
  return { contents };
}

app.post("/api/queryAI", async (req, res) => {
  try {
    const body = req.body;
    const model = body.model || defaultModel;
    const { contents } = convertMessages(body.messages || []);
    const params = {
      model,
      contents,
      config: {
        temperature: body.temperature ?? 0.7,
        maxOutputTokens: body.max_tokens ?? 4096,
      },
    };
    if (body.stream) {
      await handleStream(params, res);
    } else {
      await handleNonStream(params, model, res);
    }
  } catch (error) {
    handleError(error, res, req.body.stream);
  }
});

const handleStream = async (params, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  try {
    const response = await ai.models.generateContentStream(params);
    const created = Math.floor(Date.now() / 1000);
    const requestId = "gemini-" + Date.now();
    for await (const chunk of response) {
      const text = chunk.text || "";
      if (!text) continue;
      const openaiChunk = {
        id: requestId,
        object: "chat.completion.chunk",
        created,
        model: params.model,
        choices: [{
          index: 0,
          delta: { content: text },
          finish_reason: null,
        }],
      };
      res.write(`data: ${JSON.stringify(openaiChunk)}\n\n`);
    }
    res.write("data: [DONE]\n\n");
    res.end();
  } catch (error) {
    res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
    res.end();
  }
}

const handleNonStream = async (params, model, res) => {
  const response = await ai.models.generateContent(params);
  const text = response.text || "";
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    id: "gemini-" + Date.now(),
    object: "chat.completion",
    created: Math.floor(Date.now() / 1000),
    model,
    choices: [{
      index: 0,
      message: { role: "assistant", content: text },
      finish_reason: "stop",
    }],
    usage: { prompt_tokens: 0, completion_tokens: 0, total_tokens: 0 },
  });
}

const handleError = (error, res, isStream) => {
  console.error("Error:", error);
  if (isStream) {
    res.write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
    res.end();
  } else {
    res.status(500).json({ error: error.message });
  }
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});