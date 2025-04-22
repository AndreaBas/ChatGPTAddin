async function saveKey() {
  const key = document.getElementById("apikey").value;
  await OfficeRuntime.storage.setItem("openai-key", key);
  document.getElementById("status").innerText = "✅ Key saved!";
}

async function clearKey() {
  await OfficeRuntime.storage.removeItem("openai-key");
  document.getElementById("apikey").value = "";
  document.getElementById("status").innerText = "🧹 API key cleared.";
}
