const axios = require("axios");

async function runDifyWorkflow(text) {
  const apiKey = process.env.DIFY_API_KEY;
  const apiUrl = process.env.DIFY_API_URL || "https://api.dify.ai/v1";

  if (!apiKey) {
    throw new Error("DIFY_API_KEY is not configured");
  }

  if (!text || !text.trim()) {
    throw new Error("Input text is empty");
  }

  try {
    const response = await axios.post(
      `${apiUrl}/workflows/run`,
      {
        inputs: {
          text: text
        },
        response_mode: "blocking",
        user: "wechat-user"
      },
      {
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        timeout: 120000
      }
    );

    const data = response.data;

    if (!data || !data.data) {
      throw new Error("Dify response missing data");
    }

    if (data.data.status !== "succeeded") {
      throw new Error(
        data.data.error ||
        `Dify workflow status: ${data.data.status}`
      );
    }

    if (!data.data.outputs) {
      throw new Error("Dify response missing outputs");
    }

    const result = data.data.outputs.result;

    if (result === undefined || result === null) {
      throw new Error("Dify output 'result' not found");
    }

    return result;

  } catch (error) {
    if (error.response) {
      const message =
        error.response.data?.message ||
        error.response.data?.error ||
        `Dify request failed: HTTP ${error.response.status}`;

      throw new Error(message);
    }

    throw error;
  }
}

module.exports = {
  runDifyWorkflow
};
