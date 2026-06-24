fetch("http://localhost:3000/api/plagiarism-check", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({text: "Hello there this is an evaluation. Cấu trúc JSON báo cáo sáng kiến", model: "gemini-3.1-flash-lite"})
}).then(r => r.text()).then(console.log).catch(console.error);
