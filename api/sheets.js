export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");
  const API_KEY = process.env.GOOGLE_API_KEY;
  const SHEET_ID = process.env.GOOGLE_SHEET_ID;
  if (!API_KEY || !SHEET_ID) return res.status(500).json({ error: "Missing GOOGLE_API_KEY or GOOGLE_SHEET_ID." });
  const { action, sheetName } = req.query;
  try {
    if (action === "list") {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?key=${API_KEY}&fields=sheets.properties`;
      const resp = await fetch(url);
      if (!resp.ok) return res.status(resp.status).json({ error: await resp.text() });
      const data = await resp.json();
      const sheets = (data.sheets || []).map(s => ({ sheetId: s.properties.sheetId, title: s.properties.title, index: s.properties.index }));
      return res.status(200).json({ sheets });
    }
    if (action === "data" && sheetName) {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(sheetName)}?key=${API_KEY}`;
      const resp = await fetch(url);
      if (!resp.ok) return res.status(resp.status).json({ error: await resp.text() });
      const data = await resp.json();
      const rows = data.values || [];
      if (rows.length < 2) return res.status(200).json({ headers: [], data: [] });
      const headers = rows[0];
      const records = rows.slice(1).map(row => { const obj = {}; headers.forEach((h, i) => { obj[h] = row[i] || ""; }); return obj; });
      return res.status(200).json({ headers, data: records });
    }
    if (action === "newcustomer") {
      const targetMonth = parseInt(req.query.month) || 0, targetYear = parseInt(req.query.year) || 0;
      const listUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?key=${API_KEY}&fields=sheets.properties`;
      const listResp = await fetch(listUrl);
      if (!listResp.ok) return res.status(listResp.status).json({ error: await listResp.text() });
      const listData = await listResp.json();
      const dealerSheets = (listData.sheets || []).map(s => s.properties.title).filter(t => { const lc = t.toLowerCase(); return !lc.includes("overall") && !lc.includes("sheet1") && !lc.includes("template") && t.length > 1; });
      const ranges = dealerSheets.map(name => encodeURIComponent(name));
      const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchGet?key=${API_KEY}&ranges=${ranges.join("&ranges=")}`;
      const batchResp = await fetch(batchUrl);
      if (!batchResp.ok) return res.status(batchResp.status).json({ error: await batchResp.text() });
      const batchData = await batchResp.json();
      const valueRanges = batchData.valueRanges || [];
      const MONTH_NAMES = { january:1,february:2,march:3,april:4,may:5,june:6,july:7,august:8,september:9,october:10,november:11,december:12,jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
      const counts = {}; let total = 0; const condCounts = {}; const condTotal = { new: 0, used: 0 };
      for (let si = 0; si < valueRanges.length; si++) {
        const sheetTitle = dealerSheets[si]; const rows = valueRanges[si].values || [];
        if (rows.length < 2) continue;
        const headers = rows[0].map(h => (h || "").trim().toLowerCase());
        const prevCustIdx = headers.findIndex(h => h.includes("previous") && h.includes("customer"));
        const conditionIdx = headers.findIndex(h => h === "condition" || h === "new/used" || h.includes("condition"));
        const dateIdx = headers.findIndex(h => h.includes("date") || h === "month" || h.includes("sale date"));
        const monthIdx = headers.findIndex(h => h === "month");
        const yearIdx = headers.findIndex(h => h === "year");
        let newCustCount = 0, condNew = 0, condUsed = 0;
        for (let ri = 1; ri < rows.length; ri++) {
          const row = rows[ri]; let rowMonth = 0, rowYear = 0;
          if (monthIdx >= 0 && yearIdx >= 0) { const mVal = (row[monthIdx] || "").trim().toLowerCase(); rowMonth = MONTH_NAMES[mVal] || parseInt(mVal) || 0; rowYear = parseInt(row[yearIdx]) || 0; }
          else if (dateIdx >= 0) { const ds = (row[dateIdx] || "").trim(); const pd = new Date(ds); if (!isNaN(pd.getTime())) { rowMonth = pd.getMonth() + 1; rowYear = pd.getFullYear(); } else { const parts = ds.split(/[\s\/\-]+/); for (const p of parts) { const ml = p.toLowerCase(); if (MONTH_NAMES[ml]) rowMonth = MONTH_NAMES[ml]; const yr = parseInt(p); if (yr > 2000 && yr < 2100) rowYear = yr; } } }
          const monthMatch = !targetMonth || rowMonth === targetMonth, yearMatch = !targetYear || rowYear === targetYear;
          if (!monthMatch || !yearMatch) continue;
          if (prevCustIdx >= 0 && (row[prevCustIdx] || "").trim().toUpperCase() === "N") newCustCount++;
          if (conditionIdx >= 0) { const cond = (row[conditionIdx] || "").trim().toLowerCase(); if (cond === "new" || cond === "n") condNew++; else if (cond === "used" || cond === "u" || cond === "pre-owned") condUsed++; }
        }
        counts[sheetTitle] = newCustCount; total += newCustCount;
        condCounts[sheetTitle] = { new: condNew, used: condUsed }; condTotal.new += condNew; condTotal.used += condUsed;
      }
      return res.status(200).json({ counts, total, condCounts, condTotal });
    }
    return res.status(400).json({ error: "Invalid action." });
  } catch (err) { return res.status(500).json({ error: err.message }); }
}
