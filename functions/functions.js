/**
 * Call ChatGPT from Excel
 * @customfunction
 * @param {string} prompt Instruction (e.g., "Translate:")
 * @param {string} input The input text
 * @param {number} [temperature] Optional temperature (default: 0.7)
 * @returns {string}
 */
async function ChatGPT(prompt, input, temperature = 0.7) {
  const fullPrompt = `${prompt} ${input}`;
  const apiKey = await OfficeRuntime.storage.getItem("openai-key");

  if (!apiKey) {
    return "❌ No API key set. Open task pane and enter your OpenAI key.";
  }

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [{ role: "user", content: fullPrompt }],
        max_tokens: 150,
        temperature: temperature
      })
    });

    const data = await response.json();
    return data.choices[0].message.content;
  } catch (error) {
    return "❌ Error: " + error.message;
  }
}
