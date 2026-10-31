// Runs the control panel's JavaScript in Node with a tiny fake DOM, against the REAL local server.
// usage: node panel_dom_test.js <port> <token>      (needs Node 18+ for fetch)
const [port, token] = process.argv.slice(2);
const base = `http://127.0.0.1:${port}`;

class El {
  constructor(tag, id) {
    this.tag = tag; this.id = id || ""; this.children = []; this.textContent = ""; this.className = ""; this.hidden = false;
    this.disabled = false; this.value = ""; this.dataset = {}; this.style = {}; this.scrollTop = 0; this.clientHeight = 100; this.scrollHeight = 100;
    this.checked = false; this.onclick = null; this.tagName = /^b[A-Z]/.test(id || "") ? "BUTTON" : (/^l[A-Z]/.test(id || "") ? "A" : "DIV");
  }
  append(...n) { this.children.push(...n); }
  replaceChildren(...n) { this.children = n; }
  createTHead() { return this; } createTBody() { return this; }
  insertRow() { const r = new El("tr"); this.children.push(r); return r; }
  insertCell() { const c = new El("td"); this.children.push(c); return c; }
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
const fail = m => { console.log("DOM-FAIL: " + m); process.exit(1); };

(async () => {
  const html = await (await fetch(`${base}/?t=${token}`)).text();
  const js = html.match(/<script>([\s\S]*)<\/script>/)[1];
  const els = new Map();
  const get = id => { if (!els.has(id)) els.set(id, new El("div", id)); return els.get(id); };
  const openers = [...html.matchAll(/data-open="([a-z_]+)"/g)].map(m => { const e = new El("button"); e.dataset.open = m[1]; return e; });
  const document = {
    getElementById: get, createElement: t => new El(t), body: new El("body"),
    querySelectorAll: sel => (sel === "[data-open]" ? openers : []),
  };
  get("dShadow").checked = true; get("dFrom").value = "2026-10-01"; get("dTo").value = "2026-10-01";
  const alerts = [];
  // a browser resolves "/api/..." against the page address; Node needs it spelled out
  const pageFetch = (p, o) => fetch(p.startsWith("/") ? base + p : p, o);
  new Function("document", "fetch", "confirm", "setInterval", js)(document, pageFetch, () => true, () => 0);

  // 1. the overview fills the page
  await sleep(700);
  if (get("modeBadge").textContent !== "TEST MODE") fail("mode badge is '" + get("modeBadge").textContent + "'");
  if (!get("whereFiles").textContent.includes("TestOutput")) fail("destination text: " + get("whereFiles").textContent);
  if (!get("attnText").textContent) fail("attention text empty");
  console.log("overview OK:", get("modeBadge").textContent, "|", get("whereFiles").textContent.slice(0, 60));

  // 2. start the demo and watch it finish
  await get("lDemo").onclick();
  await sleep(300);
  if (get("runTitle").textContent !== "Running…" && !get("runTitle").textContent.includes("Demo")) fail("run title: " + get("runTitle").textContent);
  for (let i = 0; i < 150 && !get("runTitle").textContent.includes("finished"); i++) await sleep(500);
  if (!get("runTitle").textContent.includes("Demo finished")) fail("run did not finish: " + get("runTitle").textContent);
  if (!get("log").textContent.includes("RUN SUMMARY")) fail("log has no summary");
  if (get("stats").children.length !== 9) fail("result cards: " + get("stats").children.length);
  const cards = get("stats").children.map(c => c.children[1].textContent + "=" + c.children[0].textContent).join(", ");
  console.log("demo OK:", get("runTitle").textContent, "|", cards);
  if (get("bStop").hidden !== true) fail("stop button still visible");

  // 3. refresh picked up the new numbers
  await sleep(1200);
  if (get("attnStats").children.length !== 5) fail("attention cards: " + get("attnStats").children.length);

  // 4. a bad action shows a friendly message instead of crashing
  get("msgFolder").value = "/definitely/not/here";
  await get("bMsgGo").onclick(); await sleep(400);
  if (!get("alert").textContent.includes("does not exist")) fail("alert: '" + get("alert").textContent + "'");
  console.log("friendly error OK:", get("alert").textContent.slice(0, 50));

  // 5. find mailboxes builds the table with a Save button
  await get("bDetect").onclick(); await sleep(1500);
  const box = get("mbTable").children;
  if (box.length !== 3) fail("mailbox table children: " + box.length);             // table, hint, Save button
  if (box[2].textContent !== "Save these") fail("no Save button: " + box[2].textContent);
  const firstRow = box[0].children.find(c => c.tag === "tr" && c.children.length && c.children[1] && c.children[1].children.length);
  if (!firstRow || firstRow.children[1].children[0].tag !== "select") fail("the 'What it is' column has no drop-down");
  console.log("mailboxes OK: rows =", box[0].children[1].children.length);

  // 6. opening a file that does not exist yet reports it
  await openers.find(b => b.dataset.open === "to_fix").onclick(); await sleep(400);
  console.log("open OK: '" + get("alert").textContent.slice(0, 50) + "'");
  console.log("DOM-OK");
  process.exit(0);
})().catch(e => fail(e.stack || String(e)));
