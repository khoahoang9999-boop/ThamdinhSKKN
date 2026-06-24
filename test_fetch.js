fetch("http://localhost:3000/api/evaluate", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({initiativeTitle: "Test", initiativeText: "Hello there this is an evaluation."})
}).then(r => r.text()).then(console.log).catch(console.error);
