/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function sendMessage() {
  const input = document.getElementById("userInput").value;
  if (!input) return;

  appendToChat("You", input);

  // Step 1: Get selected cell content from Excel
  let selectedText = "";
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("text");
      await context.sync();
      selectedText = range.text;
    });
  } catch (error) {
    appendToChat("System", "⚠️ Failed to get selected cell.");
    console.error(error);
  }

  // Step 2: Send message + context to Groq API
  try {
    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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" // Replace this
      },
      body: JSON.stringify({
        model: "llama3-8b-8192",
        messages: [
          { role: "system", content: "You're an Excel assistant. Help the user based on their question and selected cell content." },
          { role: "user", content: `User question: ${input}\nSelected cell: ${selectedText}` }
        ]
      }),
    });

    const result = await response.json();
    const aiReply = result.choices?.[0]?.message?.content || "No response.";
    appendToChat("AI", aiReply);
  } catch (err) {
    appendToChat("System", "❌ Error connecting to Groq API.");
    console.error(err);
  }

  document.getElementById("userInput").value = "";
}

function appendToChat(sender, message) {
  const box = document.getElementById("chatBox");
  const msg = document.createElement("div");
  msg.innerHTML = `<strong>${sender}:</strong> ${message}`;
  box.appendChild(msg);
  box.scrollTop = box.scrollHeight;
}

window.sendMessage = sendMessage;