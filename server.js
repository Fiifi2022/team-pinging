const http = require("http");
const { WebSocketServer } = require("ws");
const fs = require("fs");
const path = require("path");
const os = require("os");

function getLocalIP() {
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name] || []) {
      if (net.family === "IPv4" && !net.internal) return net.address;
    }
  }
  return "localhost";
}

const PORT = process.env.PORT || 3000;
const localIP = getLocalIP();
const DATA_FILE = path.join(__dirname, "data.json");
const ARCHIVE_REPORT_FILE = path.join(__dirname, "archive-tracker.xls");
const DEFAULT_AUDIT_COUNTS = {
  total: 0,
  byNest: { "GH-1": 0, "GH-2": 0, "GH-3": 0, "GH-4": 0, "GH-5": 0, "GH-6": 0 }
};
const NESTS = Object.keys(DEFAULT_AUDIT_COUNTS.byNest);

function cloneDefaultAuditCounts() {
  return {
    total: 0,
    byNest: { ...DEFAULT_AUDIT_COUNTS.byNest }
  };
}

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, "utf8");
      const d = JSON.parse(raw);
      return {
        orders: new Map(Object.entries(d.orders || {})),
        auditCounts: d.auditCounts || cloneDefaultAuditCounts(),
        archivedOrders: Array.isArray(d.archivedOrders) ? d.archivedOrders : [],
        dailySummaries: Array.isArray(d.dailySummaries) ? d.dailySummaries : []
      };
    }
  } catch (e) {
    console.error("Failed to load data:", e.message);
  }

  return {
    orders: new Map(),
    auditCounts: cloneDefaultAuditCounts(),
    archivedOrders: [],
    dailySummaries: []
  };
}

let { orders, auditCounts, archivedOrders, dailySummaries } = loadData();
let sharedReminderTime = null;
let adminEvents = [];

function eventTime(ts = Date.now()) {
  return new Date(ts).toLocaleString([], {
    year: "numeric", month: "short", day: "2-digit",
    hour: "2-digit", minute: "2-digit"
  });
}

function logAdminEvent(action, name, detail = "") {
  const evt = { action, name, detail, timestamp: Date.now(), time: eventTime() };
  adminEvents.unshift(evt);
  adminEvents = adminEvents.slice(0, 100);
  broadcast({ type: "admin_event", event: evt });
}

function onlineMembers() {
  return [...clients.values()]
    .filter((c) => c.name && c.status !== "offline")
    .map((c) => ({
      name: c.name,
      status: c.status || "online",
      role: c.role || "sender",
      soundMuted: !!c.soundMuted,
      joinedAt: c.joinedAt || null,
      joinedTime: c.joinedAt ? eventTime(c.joinedAt) : ""
    }));
}

function saveData() {
  try {
    const d = {
      orders: Object.fromEntries(orders),
      auditCounts,
      archivedOrders,
      dailySummaries
    };
    fs.writeFileSync(DATA_FILE, JSON.stringify(d, null, 2), "utf8");
  } catch (e) {
    console.error("Failed to save data:", e.message);
  }
}

function writeArchiveReportFile() {
  try {
    fs.writeFileSync(ARCHIVE_REPORT_FILE, generateArchiveWorkbookHtml(), 'utf8');
  } catch (e) {
    console.error('Failed to write archive tracker:', e.message);
  }
}


function buildDailySummary(date, archivedBatch) {
  const summary = {
    date,
    archivedCount: archivedBatch.length,
    completedCount: 0,
    pendingCount: 0,
    medicalCount: 0,
    vaccineCount: 0,
    emergencyCount: 0,
    replenishmentCount: 0,
    scheduledCount: 0,
    avgCompletionMins: "",
    byNest: { ...DEFAULT_AUDIT_COUNTS.byNest }
  };

  let totalMins = 0;
  let totalTimedDone = 0;

  archivedBatch.forEach((order) => {
    if (order.done) summary.completedCount++;
    else summary.pendingCount++;

    if (order.type === "Medical") summary.medicalCount++;
    if (order.type === "Vaccine") summary.vaccineCount++;
    if (order.priority === "Emergency") summary.emergencyCount++;
    if (order.priority === "Replenishment") summary.replenishmentCount++;
    if (order.priority === "Scheduled") summary.scheduledCount++;

    if (order.nest && summary.byNest[order.nest] !== undefined) {
      summary.byNest[order.nest]++;
    }

    if (order.done && order.submitTimestamp && order.doneTimestamp) {
      totalMins += Math.round((order.doneTimestamp - order.submitTimestamp) / 60000);
      totalTimedDone++;
    }
  });

  summary.avgCompletionMins = totalTimedDone ? Math.round(totalMins / totalTimedDone) : "";
  return summary;
}

function appendDailySummary(summary) {
  const existing = dailySummaries.find((item) => item.date === summary.date);
  if (!existing) {
    dailySummaries.push(summary);
    return;
  }

  const prevCompleted = existing.completedCount;
  const prevAvg = Number(existing.avgCompletionMins || 0);
  const incomingCompleted = summary.completedCount;
  const incomingAvg = Number(summary.avgCompletionMins || 0);

  existing.archivedCount += summary.archivedCount;
  existing.completedCount += summary.completedCount;
  existing.pendingCount += summary.pendingCount;
  existing.medicalCount += summary.medicalCount;
  existing.vaccineCount += summary.vaccineCount;
  existing.emergencyCount += summary.emergencyCount;
  existing.replenishmentCount += summary.replenishmentCount;
  existing.scheduledCount += summary.scheduledCount;

  NESTS.forEach((nest) => {
    existing.byNest[nest] = (existing.byNest[nest] || 0) + (summary.byNest[nest] || 0);
  });

  const totalCompleted = prevCompleted + incomingCompleted;
  existing.avgCompletionMins = totalCompleted
    ? Math.round(((prevAvg * prevCompleted) + (incomingAvg * incomingCompleted)) / totalCompleted)
    : "";
}


function saveDailyXlsxFromServer(date = todayKey(), testMode = false) {
  const xlsxDir = path.join(__dirname, "xlsx-reports");
  if (!fs.existsSync(xlsxDir)) fs.mkdirSync(xlsxDir);
  const fileName = `${testMode ? "TeamPing_TEST" : "TeamPing"}_${date}.xlsx`;
  const outFile = path.join(xlsxDir, fileName);
  const { execSync } = require("child_process");
  const payload = JSON.stringify({
    date,
    orders: [...orders.values()],
    archivedOrders: archivedOrders.filter(o => o.archiveDate === date),
    dailySummaries: dailySummaries.filter(s => s.date === date),
    allSummaries: dailySummaries
  });
  const script = path.join(__dirname, "make_xlsx.py");
  execSync(`python3 ${script} '${outFile}' '${date}'`, {
    input: payload,
    encoding: "utf8",
    timeout: 15000
  });
  return fileName;
}

function archiveCurrentOrders(reason = "manual_reset") {
  const currentOrders = [...orders.values()];
  const date = todayKey();

  if (!currentOrders.length) {
    auditCounts = cloneDefaultAuditCounts();
    saveData();
    return { archivedBatch: [], date, reason };
  }

  const archivedBatch = currentOrders.map((order) => ({
    ...order,
    archiveDate: date,
    archiveReason: reason,
    archivedAt: new Date().toISOString()
  }));

  archivedOrders.push(...archivedBatch);
  appendDailySummary(buildDailySummary(date, archivedBatch));
  orders = new Map();
  auditCounts = cloneDefaultAuditCounts();
  saveData();

  // Immediately write the refreshed daily summary to the Excel reports folder.
  try {
    const file = saveDailyXlsxFromServer(date, false);
    console.log("Daily report Excel saved after summary refresh:", file);
  } catch (e) {
    console.error("Failed to save daily report after summary refresh:", e.message);
  }

  // Check if it's end of month — if so, schedule monthly combine
  const now = new Date();
  const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  if (tomorrow.getDate() === 1) {
    // End of month — combine will be triggered by client after xlsx save
    console.log("End of month detected — monthly combine will run after daily xlsx is saved.");
  }

  return { archivedBatch, date, reason };
}

function generateArchiveWorkbookHtml() {
  const summaries = [...dailySummaries].sort((a, b) => b.date.localeCompare(a.date));
  const archiveRows = [...archivedOrders].sort((a, b) => {
    const aDate = `${a.archiveDate || ""} ${a.archivedAt || ""}`;
    const bDate = `${b.archiveDate || ""} ${b.archivedAt || ""}`;
    return bDate.localeCompare(aDate);
  });

  const summaryHeaders = ["Date", "Archived", "Completed", "Pending", "Medical", "Vaccine", "Emergency", "Replenishment", "Scheduled", "Avg Completion (mins)", "GH-1", "GH-2", "GH-3", "GH-4", "GH-5", "GH-6"];
  const archiveHeaders = ["Archive Date", "Archived At", "Order ID", "Nest", "Priority", "Type", "Submitted By", "Submitted Time", "Status", "Completed By", "Completed At", "Time Taken (mins)", "Notes", "Link 1", "Link 2"];

  const summaryTable = summaries.length
    ? summaries.map((s) => `
      <tr>
        <td>${escapeHtml(s.date)}</td>
        <td>${escapeHtml(s.archivedCount)}</td>
        <td>${escapeHtml(s.completedCount)}</td>
        <td>${escapeHtml(s.pendingCount)}</td>
        <td>${escapeHtml(s.medicalCount)}</td>
        <td>${escapeHtml(s.vaccineCount)}</td>
        <td>${escapeHtml(s.emergencyCount)}</td>
        <td>${escapeHtml(s.replenishmentCount)}</td>
        <td>${escapeHtml(s.scheduledCount)}</td>
        <td>${escapeHtml(s.avgCompletionMins)}</td>
        <td>${escapeHtml(s.byNest?.["GH-1"] || 0)}</td>
        <td>${escapeHtml(s.byNest?.["GH-2"] || 0)}</td>
        <td>${escapeHtml(s.byNest?.["GH-3"] || 0)}</td>
        <td>${escapeHtml(s.byNest?.["GH-4"] || 0)}</td>
        <td>${escapeHtml(s.byNest?.["GH-5"] || 0)}</td>
        <td>${escapeHtml(s.byNest?.["GH-6"] || 0)}</td>
      </tr>
    `).join("")
    : `<tr><td colspan="${summaryHeaders.length}">No daily summaries yet.</td></tr>`;

  const archiveTable = archiveRows.length
    ? archiveRows.map((order) => {
        const mins = order.submitTimestamp && order.doneTimestamp ? Math.round((order.doneTimestamp - order.submitTimestamp) / 60000) : "";
        return `
          <tr>
            <td>${escapeHtml(order.archiveDate)}</td>
            <td>${escapeHtml(order.archivedAt)}</td>
            <td>${escapeHtml(order.orderID)}</td>
            <td>${escapeHtml(order.nest)}</td>
            <td>${escapeHtml(order.priority)}</td>
            <td>${escapeHtml(order.type)}</td>
            <td>${escapeHtml(order.from)}</td>
            <td>${escapeHtml(order.time)}</td>
            <td>${escapeHtml(order.done ? "Completed" : "Pending")}</td>
            <td>${escapeHtml(order.doneBy || "")}</td>
            <td>${escapeHtml(order.doneTime || "")}</td>
            <td>${escapeHtml(mins)}</td>
            <td>${escapeHtml(order.notes || "")}</td>
            <td>${escapeHtml(order.link || "")}</td>
            <td>${escapeHtml(order.link2 || "")}</td>
          </tr>
        `;
      }).join("")
    : `<tr><td colspan="${archiveHeaders.length}">No archived orders yet.</td></tr>`;

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8" />
<style>
  body { font-family: Arial, sans-serif; margin: 20px; color: #222; }
  h1 { color: #1d6fc4; margin-bottom: 6px; }
  h2 { margin: 24px 0 8px; color: #111; }
  p { margin: 0 0 12px; color: #555; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 28px; }
  th, td { border: 1px solid #d0d7de; padding: 8px 10px; font-size: 12px; text-align: left; }
  th { background: #1d6fc4; color: white; font-weight: 700; }
  tr:nth-child(even) td { background: #f8fafc; }
</style>
</head>
<body>
  <h1>Team Ping Archive Tracker</h1>
  <p>Generated: ${escapeHtml(new Date().toISOString())}</p>

  <h2>Daily Summary</h2>
  <table>
    <thead><tr>${summaryHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
    <tbody>${summaryTable}</tbody>
  </table>

  <h2>Archived Orders</h2>
  <table>
    <thead><tr>${archiveHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
    <tbody>${archiveTable}</tbody>
  </table>
</body>
</html>`;
}

console.log(`Loaded ${orders.size} active orders and ${archivedOrders.length} archived orders from disk.`);
writeArchiveReportFile();

const httpServer = http.createServer((req, res) => {
  if (req.url === "/" || req.url === "/index.html") {
    const filePath = path.join(__dirname, "client.html");
    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(500);
        res.end("Error");
        return;
      }
      res.writeHead(200, { "Content-Type": "text/html" });
      res.end(data);
    });
    return;
  }

  if (req.url === "/reset-day" && req.method === "POST") {
    archiveCurrentOrders("http_reset");
    broadcast({ type: "day_reset" });
    res.writeHead(200);
    res.end("ok");
    return;
  }

  if (req.url === "/archive-report.xls" && req.method === "GET") {
    const filename = `team-ping-archive-${todayKey()}.xls`;
    if (!fs.existsSync(ARCHIVE_REPORT_FILE)) writeArchiveReportFile();
    res.writeHead(200, {
      "Content-Type": "application/vnd.ms-excel",
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Cache-Control": "no-store"
    });
    fs.createReadStream(ARCHIVE_REPORT_FILE).pipe(res);
    return;
  }

  if (req.url === "/archive-report.json" && req.method === "GET") {
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ dailySummaries, archivedOrders }, null, 2));
    return;
  }

  if (req.url === "/archive-clear" && req.method === "POST") {
    archivedOrders = [];
    dailySummaries = [];
    saveData();
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ ok: true }));
    return;
  }

  // Save daily XLSX report
  if (req.url === "/save-daily-xlsx" && req.method === "POST") {
    const date = todayKey();
    const xlsxDir = path.join(__dirname, "xlsx-reports");
    if (!fs.existsSync(xlsxDir)) fs.mkdirSync(xlsxDir);
    const outFile = path.join(xlsxDir, `TeamPing_${date}.xlsx`);
    try {
      const fileName = saveDailyXlsxFromServer(date, false);
      broadcast({ type: "xlsx_saved", date, file: fileName });
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: true, file: fileName }));
    } catch (e) {
      console.error("xlsx generation failed:", e.message);
      res.writeHead(500);
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // Combine monthly XLSX into one reference workbook
  if (req.url.startsWith("/combine-monthly") && req.method === "POST") {
    const now = new Date();
    const urlObj = new URL(req.url, `http://${req.headers.host}`);
    const testMode = urlObj.searchParams.get("test") === "1";

    // If today is the 1st, the complete month is the previous month.
    // Otherwise use the current month for manual tests.
    const target = new Date(now.getFullYear(), now.getMonth() - (now.getDate() === 1 ? 1 : 0), 1);
    const ym = `${target.getFullYear()}-${String(target.getMonth()+1).padStart(2,"0")}`;

    const xlsxDir = path.join(__dirname, "xlsx-reports");
    if (!fs.existsSync(xlsxDir)) fs.mkdirSync(xlsxDir);

    const outName = testMode ? `TeamPing_Monthly_TEST_${ym}.xlsx` : `TeamPing_Monthly_${ym}.xlsx`;
    const outFile = path.join(xlsxDir, outName);

    try {
      const { execSync } = require("child_process");
      const script = path.join(__dirname, "combine_xlsx.py");
      execSync(`python3 ${script} '${xlsxDir}' '${outFile}' '${ym}'`, { timeout: 30000 });

      // Also create a server-side monthly summary record from archived daily summaries.
      const monthSummaries = dailySummaries.filter(s => String(s.date || "").startsWith(ym));
      const monthlySummary = monthSummaries.reduce((acc, s) => {
        acc.archivedCount += Number(s.archivedCount || 0);
        acc.completedCount += Number(s.completedCount || 0);
        acc.pendingCount += Number(s.pendingCount || 0);
        acc.medicalCount += Number(s.medicalCount || 0);
        acc.vaccineCount += Number(s.vaccineCount || 0);
        acc.emergencyCount += Number(s.emergencyCount || 0);
        acc.replenishmentCount += Number(s.replenishmentCount || 0);
        acc.scheduledCount += Number(s.scheduledCount || 0);
        for (const nest of NESTS) acc.byNest[nest] += Number(s.byNest?.[nest] || 0);
        return acc;
      }, {
        month: ym,
        archivedCount: 0,
        completedCount: 0,
        pendingCount: 0,
        medicalCount: 0,
        vaccineCount: 0,
        emergencyCount: 0,
        replenishmentCount: 0,
        scheduledCount: 0,
        byNest: { ...DEFAULT_AUDIT_COUNTS.byNest }
      });

      const timed = archivedOrders.filter(o => String(o.archiveDate || "").startsWith(ym) && o.done && o.submitTimestamp && o.doneTimestamp);
      monthlySummary.avgCompletionMins = timed.length
        ? Math.round(timed.reduce((sum, o) => sum + Math.round((o.doneTimestamp - o.submitTimestamp) / 60000), 0) / timed.length)
        : "";

      console.log("Monthly summary created:", monthlySummary);
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: true, file: outName, month: ym, summary: monthlySummary }));
    } catch (e) {
      console.error("combine failed:", e.message);
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  // List xlsx reports
  if (req.url === "/xlsx-list" && req.method === "GET") {
    const xlsxDir = path.join(__dirname, "xlsx-reports");
    if (!fs.existsSync(xlsxDir)) { res.writeHead(200, {"Content-Type":"application/json"}); res.end("[]"); return; }
    const files = fs.readdirSync(xlsxDir).filter(f => f.endsWith(".xlsx")).sort().reverse();
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(files));
    return;
  }

  // Download specific xlsx
  if (req.url.startsWith("/xlsx-download/") && req.method === "GET") {
    const fname = decodeURIComponent(req.url.replace("/xlsx-download/", "")).replace(/\.\./g, "");
    const xlsxDir = path.join(__dirname, "xlsx-reports");
    const fpath = path.join(xlsxDir, fname);
    if (!fs.existsSync(fpath) || !fname.endsWith(".xlsx")) { res.writeHead(404); res.end("Not found"); return; }
    res.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${fname}"`,
      "Cache-Control": "no-store"
    });
    fs.createReadStream(fpath).pipe(res);
    return;
  }


  // Test report settings and create a test XLSX report
  if (req.url === "/report-test" && req.method === "POST") {
    const date = todayKey();
    const xlsxDir = path.join(__dirname, "xlsx-reports");
    if (!fs.existsSync(xlsxDir)) fs.mkdirSync(xlsxDir);
    const outFile = path.join(xlsxDir, `TeamPing_TEST_${date}.xlsx`);
    try {
      const fileName = saveDailyXlsxFromServer(date, true);
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: true, file: fileName }));
    } catch (e) {
      console.error("report test failed:", e.message);
      res.writeHead(500, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: e.message }));
    }
    return;
  }

  if (req.url === "/admin-state" && req.method === "GET") {
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
    res.end(JSON.stringify({
      onlineMembers: onlineMembers(),
      activeOrders: [...orders.values()],
      dailySummaries,
      adminEvents
    }, null, 2));
    return;
  }

  res.writeHead(404);
  res.end("Not found");
});

const wss = new WebSocketServer({ server: httpServer });
const clients = new Map();

function broadcast(data, excludeWs = null) {
  const msg = JSON.stringify(data);
  for (const [ws] of clients) {
    if (ws !== excludeWs && ws.readyState === 1) ws.send(msg);
  }
}

function broadcastPresence() {
  const members = onlineMembers();
  const msg = JSON.stringify({ type: "presence_list", members });
  for (const [ws] of clients) {
    if (ws.readyState === 1) ws.send(msg);
  }
}

function sendToTargets(data, targets, excludeWs) {
  const msg = JSON.stringify(data);
  for (const [ws, info] of clients) {
    if (ws === excludeWs || ws.readyState !== 1) continue;
    if (!targets || targets.includes(info.name)) ws.send(msg);
  }
}

wss.on("connection", (ws) => {
  clients.set(ws, { name: null, status: "online", role: "sender", soundMuted: false, joinedAt: null });

  ws.on("message", (raw) => {
    let data;
    try {
      data = JSON.parse(raw);
    } catch {
      return;
    }

    if (data.type === "join") {
      const info = clients.get(ws);
      info.name = data.from;
      info.status = data.status || "online";
      info.role = data.role || "sender";
      info.soundMuted = !!info.soundMuted;
      info.joinedAt = Date.now();
      logAdminEvent("joined", data.from, `Joined as ${info.role}`);

      broadcastPresence();
      broadcast({ type: "join", from: data.from, status: info.status, role: info.role }, ws);

      ws.send(JSON.stringify({
        type: "state_sync",
        orders: [...orders.values()],
        auditCounts
      }));
      if(sharedReminderTime){
        ws.send(JSON.stringify({type:"reminder_set",from:"Team",time:sharedReminderTime,label:"Midday audit"}));
      }
      broadcastPresence();
      return;
    }

    if (data.type === "logout") {
      const info = clients.get(ws);
      if (info?.name) {
        logAdminEvent("left", info.name, "Signed out");
        broadcast({ type: "leave", from: info.name }, ws);
        info.name = null;
        info.status = "offline";
      }
      broadcastPresence();
      return;
    }

    if (data.type === "status") {
      const info = clients.get(ws);
      if (info) info.status = data.status;
      if (data.status === "offline") {
        logAdminEvent("went offline", data.from, "Changed status to offline");
        broadcast({ type: "leave", from: data.from }, ws);
      } else {
        logAdminEvent("came online", data.from, "Changed status to online");
        broadcast(data, ws);
      }
      broadcastPresence();
      return;
    }

    if (data.type === "ping") {
      const targets = data.to ? (Array.isArray(data.to) ? data.to : [data.to]) : null;
      sendToTargets(data, targets, ws);
      return;
    }

    if (data.type === "task_new") {
      orders.set(data.task.id, { ...data.task });
      saveData();
      broadcast(data, ws);
      return;
    }

    if (data.type === "task_done") {
      const order = orders.get(data.taskId);
      if (order && !order.done) {
        order.done = true;
        order.doneBy = data.by;
        order.doneTime = data.doneTime || new Date().toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit"
        });
        order.doneTimestamp = data.doneTimestamp || Date.now();

        auditCounts.total++;
        if (order.nest && auditCounts.byNest[order.nest] !== undefined) {
          auditCounts.byNest[order.nest]++;
        }

        saveData();
      }
      broadcast(data, ws);
      return;
    }

    if (data.type === "reset_day") {
      archiveCurrentOrders("socket_reset");
      broadcast({ type: "day_reset" });
      return;
    }

    if (data.type === "task_update") {
      const order = orders.get(data.taskId);
      if (order && data.patch) {
        Object.assign(order, data.patch);
        saveData();
      }
      broadcast(data, ws);
      return;
    }

    if (data.type === "task_note") {
      const order = orders.get(data.taskId);
      if (order) {
        order.notes = data.notes;
        saveData();
      }
      broadcast(data, ws);
      return;
    }

    if (data.type === "role_change") {
      const info = clients.get(ws);
      if (info) info.role = data.role;
      broadcastPresence();
      broadcast({ type: "role_change", from: data.from, role: data.role }, ws);
      return;
    }

    if (data.type === "comment_alert") {
      const msg = JSON.stringify(data);
      for (const [ws2, info] of clients) {
        if (ws2 !== ws && ws2.readyState === 1) {
          if (!data.target || info.name === data.target) ws2.send(msg);
        }
      }
    } else if (data.type === "comment_ack") {
      const order = orders.get(data.taskId);
      if (order) {
        const msg = JSON.stringify(data);
        for (const [ws2, info] of clients) {
          if (ws2 !== ws && ws2.readyState === 1 && info.name === order.from) ws2.send(msg);
        }
      }
    } else if (data.type === "task_accept") {
      const order = orders.get(data.taskId);
      if (order) { order.accepted = true; saveData(); }
      const msg = JSON.stringify(data);
      for (const [ws2, info] of clients) {
        if (ws2 !== ws && ws2.readyState === 1) {
          if (!order || info.name === order.from) ws2.send(msg);
        }
      }
    } else if (data.type === "mute_peer") {
      const msg = JSON.stringify(data);
      for (const [ws2, info] of clients) {
        if (ws2 !== ws && ws2.readyState === 1 && info.name === data.target) {
          ws2.send(msg);
        }
      }
      return;
    }


    if (data.type === "admin_remove_user") {
      const target = data.target;
      let removed = 0;
      for (const [ws2, info] of clients) {
        if (info.name === target && ws2.readyState === 1) {
          ws2.send(JSON.stringify({ type: "admin_removed", by: data.from || "Admin" }));
          info.status = "offline";
          removed++;
          setTimeout(() => { try { ws2.close(); } catch {} }, 100);
        }
      }
      logAdminEvent("removed", target, `Removed by ${data.from || "Admin"}${removed ? "" : " (user not found online)"}`);
      broadcastPresence();
      return;
    }

    if (data.type === "admin_mute_user") {
      const target = data.target;
      let changed = 0;
      for (const [ws2, info] of clients) {
        if (info.name === target && ws2.readyState === 1) {
          info.soundMuted = !!data.muted;
          ws2.send(JSON.stringify({ type: "mute_peer", from: data.from || "Admin", target, muted: !!data.muted }));
          changed++;
        }
      }
      logAdminEvent(data.muted ? "muted sound" : "unmuted sound", target, `${data.from || "Admin"} ${data.muted ? "muted" : "unmuted"} alerts${changed ? "" : " (user not found online)"}`);
      broadcastPresence();
      return;
    }

    if (data.type === "reminder_set") {
      sharedReminderTime = data.time || null;
      broadcast(data, ws);
      return;
    }

    if (data.type === "reminder_clear") {
      sharedReminderTime = null;
      broadcast(data, ws);
      return;
    }

    if (data.type === "msg_ack") {
      const msg = JSON.stringify(data);
      for (const [ws2, info] of clients) {
        if (ws2 !== ws && ws2.readyState === 1 && (!data.to || info.name === data.to)) ws2.send(msg);
      }
      return;
    }

    if (["ack", "msg"].includes(data.type)) {
      broadcast(data, ws);
    }
  });

  ws.on("close", () => {
    const info = clients.get(ws);
    if (info?.name && info.status !== "offline") {
      logAdminEvent("left", info.name, "Connection closed");
      broadcast({ type: "leave", from: info.name });
    }
    clients.delete(ws);
    broadcastPresence();
  });
});

httpServer.listen(PORT, "0.0.0.0", () => {
  console.log("\nTeam Ping Server running");
  console.log(`Local:   http://localhost:${PORT}`);
  console.log(`Network: http://${localIP}:${PORT}`);
  console.log("Archive report available at /archive-report.xls");
  console.log("Data saved to disk and survives restarts.\n");
});
