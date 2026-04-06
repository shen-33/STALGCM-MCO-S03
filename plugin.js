$("#check-btn").click(runChecker);

async function runChecker() {
  await Word.run(async (context) => {
    const results = document.getElementById("results");
    results.innerText = "Sinusuri...";

    // 1. Get the document body
    const body = context.document.body;
    body.load("text");
    await context.sync();

    // 2. CLEAN THE TEXT (Fixes Multi-line)
    // We replace all newlines (\r or \n) with a single space
    const fullText = body.text.replace(/[\r\n\t]+/g, " ").trim();

    // 3. ENCODE & SEND
    const url = `http://localhost:8081/?language=tl&text=${encodeURIComponent(fullText)}`;

    console.log("Sending to server:", fullText); // Check your browser console to see this

    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error("Server not responding");

      const xmlText = await response.text();
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlText, "text/xml");
      const errors = xmlDoc.getElementsByTagName("error");

      if (errors.length === 0) {
        results.innerHTML = "<b style='color:green;'>Mahusay!</b> Walang nahanap na mali.";
      } else {
        results.innerHTML = `<b>Nahanap na mga mali (${errors.length}):</b>`;
        for (let i = 0; i < errors.length; i++) {
          const msg = errors[i].getAttribute("msg");
          results.innerHTML += `<div style="border-left:4px solid red; padding:5px; margin:10px 0; background:#fff1f1; font-size:13px;">${msg}</div>`;
        }
      }
    } catch (e) {
      console.error(e);
      results.innerHTML = "<b style='color:red;'>Error:</b> Hindi makakonekta. <br/><small>May Error</small>";
    }
  });
}
