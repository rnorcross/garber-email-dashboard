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

    return res.status(400).json({ error: "Invalid request. Use ?action=list or ?action=data&sheetName=..." });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
