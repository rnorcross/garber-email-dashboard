// Vercel Serverless Function — /api/sheets.js
// Securely fetches dealership customer data from Google Sheets API
// The API key and Spreadsheet ID are stored as Vercel environment variables
// and never exposed to the browser.

export default async function handler(req, res) {
  // CORS headers (allows your Vercel frontend to call this)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");

  const API_KEY = process.env.GOOGLE_API_KEY;
  const SHEET_ID = process.env.GOOGLE_SHEET_ID;

  if (!API_KEY || !SHEET_ID) {
    return res.status(500).json({
      error: "Missing environment variables. Please set GOOGLE_API_KEY and GOOGLE_SHEET_ID in Vercel project settings.",
    });
  }

  const { action, sheetName } = req.query;

  try {
    // ACTION: list — returns all sheet names and IDs in the workbook
    if (action === "list") {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?key=${API_KEY}&fields=sheets.properties`;
      const resp = await fetch(url);
      if (!resp.ok) {
        const errText = await resp.text();
        return res.status(resp.status).json({ error: `Google API error: ${errText}` });
      }
      const data = await resp.json();
      const sheets = (data.sheets || []).map((s) => ({
        sheetId: s.properties.sheetId,
        title: s.properties.title,
        index: s.properties.index,
      }));
      return res.status(200).json({ sheets });
    }

    // ACTION: data — returns all rows from a specific sheet by name
    if (action === "data" && sheetName) {
      const encodedName = encodeURIComponent(sheetName);
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodedName}?key=${API_KEY}`;
      const resp = await fetch(url);
      if (!resp.ok) {
        const errText = await resp.text();
        return res.status(resp.status).json({ error: `Google API error: ${errText}` });
      }
      const data = await resp.json();
      const rows = data.values || [];
      if (rows.length < 2) {
        return res.status(200).json({ headers: [], data: [] });
      }
      // First row = headers, rest = data
      const headers = rows[0];
      const records = rows.slice(1).map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i] || "";
        });
        return obj;
      });
      return res.status(200).json({ headers, data: records });
    }

    // ACTION: newcustomer — counts "Previous Customer" = "N" per dealership for a given month/year
    // Query params: month (1-12), year (e.g. 2025)
    // Returns: { counts: { "Dealership Name": count, ... }, total: number }
    if (action === "newcustomer") {
      const targetMonth = parseInt(req.query.month) || 0;
      const targetYear = parseInt(req.query.year) || 0;

      // Step 1: Get all sheet names
      const listUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?key=${API_KEY}&fields=sheets.properties`;
      const listResp = await fetch(listUrl);
      if (!listResp.ok) {
        const errText = await listResp.text();
        return res.status(listResp.status).json({ error: `Google API error: ${errText}` });
      }
      const listData = await listResp.json();
      const dealerSheets = (listData.sheets || [])
        .map((s) => s.properties.title)
        .filter((t) => {
          const lc = t.toLowerCase();
          return !lc.includes("overall") && !lc.includes("sheet1") && !lc.includes("template") && t.length > 1;
        });

      // Step 2: Build a batch request for all dealership sheets at once
      const ranges = dealerSheets.map((name) => encodeURIComponent(name));
      const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values:batchGet?key=${API_KEY}&ranges=${ranges.join("&ranges=")}`;
      const batchResp = await fetch(batchUrl);
      if (!batchResp.ok) {
        const errText = await batchResp.text();
        return res.status(batchResp.status).json({ error: `Google API error: ${errText}` });
      }
      const batchData = await batchResp.json();
      const valueRanges = batchData.valueRanges || [];

      const MONTH_NAMES = { january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12, jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

      const counts = {};
      let total = 0;

      for (let si = 0; si < valueRanges.length; si++) {
        const sheetTitle = dealerSheets[si];
        const rows = valueRanges[si].values || [];
        if (rows.length < 2) continue;

        const headers = rows[0].map((h) => (h || "").trim().toLowerCase());

        // Find "Previous Customer" column (flexible matching)
        const prevCustIdx = headers.findIndex((h) =>
          h.includes("previous") && h.includes("customer")
        );
        if (prevCustIdx < 0) continue;

        // Find date/month column (flexible matching)
        const dateIdx = headers.findIndex((h) =>
          h.includes("date") || h === "month" || h.includes("sale date") || h.includes("sold date")
        );
        const monthIdx = headers.findIndex((h) => h === "month");
        const yearIdx = headers.findIndex((h) => h === "year");

        let count = 0;
        for (let ri = 1; ri < rows.length; ri++) {
          const row = rows[ri];
          const prevCust = (row[prevCustIdx] || "").trim().toUpperCase();
          if (prevCust !== "N") continue;

          // If no month/year filter requested, count all N's
          if (!targetMonth && !targetYear) {
            count++;
            continue;
          }

          // Try to match month/year from the row
          let rowMonth = 0, rowYear = 0;

          // Method A: Separate month and year columns
          if (monthIdx >= 0 && yearIdx >= 0) {
            const mVal = (row[monthIdx] || "").trim().toLowerCase();
            rowMonth = MONTH_NAMES[mVal] || parseInt(mVal) || 0;
            rowYear = parseInt(row[yearIdx]) || 0;
          }
          // Method B: Parse a date column
          else if (dateIdx >= 0) {
            const dateStr = (row[dateIdx] || "").trim();
            const parsed = new Date(dateStr);
            if (!isNaN(parsed.getTime())) {
              rowMonth = parsed.getMonth() + 1;
              rowYear = parsed.getFullYear();
            } else {
              // Try "Month Year" format like "January 2025"
              const parts = dateStr.split(/[\s\/\-]+/);
              for (const part of parts) {
                const ml = part.toLowerCase();
                if (MONTH_NAMES[ml]) rowMonth = MONTH_NAMES[ml];
                const yr = parseInt(part);
                if (yr > 2000 && yr < 2100) rowYear = yr;
              }
            }
          }

          const monthMatch = !targetMonth || rowMonth === targetMonth;
          const yearMatch = !targetYear || rowYear === targetYear;
          if (monthMatch && yearMatch) count++;
        }

        counts[sheetTitle] = count;
        total += count;
      }

      return res.status(200).json({ counts, total });
    }

    return res.status(400).json({ error: "Invalid request. Use ?action=list or ?action=data&sheetName=..." });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
