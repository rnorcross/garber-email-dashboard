import { put, list, del } from "@vercel/blob";
import crypto from "crypto";

const MAPPING_KEY = "slides-mapping.json";
const BLOB_DOMAIN = "vbaadun5aa5zljfi.public.blob.vercel-storage.com";

async function getAccessToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const b64 = (obj) => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const unsigned = `${b64({ alg: "RS256", typ: "JWT" })}.${b64({
    iss: sa.client_email,
    scope: "https://www.googleapis.com/auth/presentations",
    aud: "https://oauth2.googleapis.com/token",
    iat: now, exp: now + 3600,
  })}`;
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(unsigned);
  const jwt = `${unsigned}.${sign.sign(sa.private_key, "base64url")}`;
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error("Auth failed: " + JSON.stringify(data));
  return data.access_token;
}

// Load saved objectId → filename mapping from Blob storage
async function loadMapping() {
  try {
    const { blobs } = await list({ prefix: MAPPING_KEY });
    if (blobs.length === 0) return null;
    const resp = await fetch(blobs[0].url);
    if (!resp.ok) return null;
    return await resp.json();
  } catch { return null; }
}

// Save objectId → filename mapping to Blob storage
async function saveMapping(mapping) {
  await put(MAPPING_KEY, JSON.stringify(mapping, null, 2), {
    access: "public", addRandomSuffix: false, contentType: "application/json",
  });
}

// Extract all text from a slide's text elements
function getSlideText(slide) {
  const texts = [];
  for (const el of slide.pageElements || []) {
    if (el.shape && el.shape.text) {
      for (const te of el.shape.text.textElements || []) {
        if (te.textRun && te.textRun.content) texts.push(te.textRun.content.trim());
      }
    }
    if (el.table) {
      for (const row of el.table.tableRows || []) {
        for (const cell of row.tableCells || []) {
          if (cell.text) {
            for (const te of cell.text.textElements || []) {
              if (te.textRun && te.textRun.content) texts.push(te.textRun.content.trim());
            }
          }
        }
      }
    }
  }
  return texts.filter(t => t.length > 1).join(" ");
}

// Try to match slide text to a screenshot filename
function matchTextToFile(slideText, fileNames) {
  const text = slideText.toLowerCase();
  let bestMatch = null;
  let bestScore = 0;
  for (const fn of fileNames) {
    // Extract dealership name and type from filename like "Garber_Honda_Sales_Email.jpg"
    const parts = fn.replace(".jpg", "").split("_");
    // Determine the type (Sales_Email, Service_Email, Advertising)
    let type = "";
    let dealerParts = [];
    if (fn.includes("Sales_Email")) { type = "sales"; dealerParts = fn.replace("_Sales_Email.jpg", "").split("_"); }
    else if (fn.includes("Service_Email")) { type = "service"; dealerParts = fn.replace("_Service_Email.jpg", "").split("_"); }
    else if (fn.includes("Advertising")) { type = "advertising"; dealerParts = fn.replace("_Advertising.jpg", "").split("_"); }
    const dealerName = dealerParts.join(" ").toLowerCase();
    if (!dealerName) continue;
    // Score: how many dealer name words appear in the slide text?
    const words = dealerParts.map(w => w.toLowerCase());
    let score = 0;
    for (const w of words) { if (text.includes(w.toLowerCase())) score++; }
    // Bonus for type match
    if (type === "sales" && (text.includes("sales email") || text.includes("sales_email"))) score += 2;
    if (type === "service" && (text.includes("service email") || text.includes("service_email"))) score += 2;
    if (type === "advertising" && (text.includes("advertising") || text.includes("google ads") || text.includes("facebook ads"))) score += 2;
    if (score > bestScore) { bestScore = score; bestMatch = fn; }
  }
  // Require at least 3 matching signals to consider it a match
  return bestScore >= 3 ? bestMatch : null;
}

// Scan presentation and build mapping
async function buildMappingFromPresentation(token, presentationId, fileNames) {
  const resp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error("Slides API error: " + (await resp.text()));
  const pres = await resp.json();
  const mapping = {};
  const debug = [];
  for (const slide of pres.slides || []) {
    const slideText = getSlideText(slide);
    const images = (slide.pageElements || []).filter(e => e.image);
    if (images.length === 0) continue;
    const matchedFile = fileNames ? matchTextToFile(slideText, fileNames) : null;
    for (const img of images) {
      const contentUrl = img.image.contentUrl || "";
      const sourceUrl = img.image.sourceUrl || "";
      let matched = false;
      // Try URL-based matching first
      for (const url of [contentUrl, sourceUrl]) {
        if (url.includes(BLOB_DOMAIN) || url.includes("screenshots/")) {
          const m = url.match(/screenshots\/([^?]+)/);
          if (m) { mapping[img.objectId] = m[1]; matched = true; break; }
        }
      }
      // Try text-based matching
      if (!matched && matchedFile) {
        mapping[img.objectId] = matchedFile;
        matched = true;
      }
      debug.push({
        objectId: img.objectId,
        slideText: slideText.substring(0, 150),
        matchedFile: matched ? (mapping[img.objectId] || null) : null,
        contentUrl: contentUrl.substring(0, 80),
      });
    }
  }
  return { mapping, debug };
}

// Replace images using saved objectId mapping
async function replaceByMapping(token, presentationId, mapping, urlMap) {
  const requests = [];
  const matched = [];
  for (const [objectId, fileName] of Object.entries(mapping)) {
    if (urlMap[fileName]) {
      requests.push({
        replaceImage: {
          imageObjectId: objectId,
          imageReplaceMethod: "CENTER_INSIDE",
          url: urlMap[fileName] + "?t=" + Date.now(),
        },
      });
      matched.push(fileName);
    }
  }
  if (requests.length === 0) {
    return { replaced: 0, matched, message: "No matching files found for mapped images." };
  }
  const updateResp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ requests }),
  });
  if (!updateResp.ok) throw new Error("Slides batch update failed: " + (await updateResp.text()));
  return { replaced: requests.length, matched, message: `Refreshed ${requests.length} images in the presentation.` };
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const { action } = req.query;
  const slidesId = process.env.GOOGLE_SLIDES_PRESENTATION_ID;
  const saKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

  try {
    // TEST — verify blob storage works
    if (action === "test") {
      const testBlob = await put("test-connection.txt", "ok", { access: "public", addRandomSuffix: false });
      await del(testBlob.url);
      const mapping = await loadMapping();
      return res.status(200).json({
        success: true,
        message: "Vercel Blob storage is connected.",
        slidesConfigured: !!slidesId,
        mappingExists: !!mapping,
        mappedImages: mapping ? Object.keys(mapping).length : 0,
      });
    }

    // UPLOAD — upload a single screenshot
    if (action === "upload" && req.method === "POST") {
      const { fileName, imageData } = req.body;
      if (!fileName || !imageData) return res.status(400).json({ error: "fileName and imageData required." });
      const buffer = Buffer.from(imageData, "base64");
      const blob = await put(`screenshots/${fileName}`, buffer, {
        access: "public", addRandomSuffix: false, contentType: "image/jpeg",
      });
      return res.status(200).json({ fileName, url: blob.url });
    }

    // LIST — list all uploaded screenshots
    if (action === "list") {
      const { blobs } = await list({ prefix: "screenshots/" });
      const files = blobs.map((b) => ({ name: b.pathname.replace("screenshots/", ""), url: b.url, size: b.size }));
      return res.status(200).json({ count: files.length, files });
    }

    // URLS — get screenshot URLs grouped by tab
    if (action === "urls") {
      const { blobs } = await list({ prefix: "screenshots/" });
      const grouped = { Sales_Email: [], Service_Email: [], Advertising: [] };
      for (const b of blobs) {
        const name = b.pathname.replace("screenshots/", "");
        if (name.includes("_2026-") || name.includes("_2025-")) continue; // skip old dated files
        const entry = { name, url: b.url };
        if (name.includes("Sales_Email")) grouped.Sales_Email.push(entry);
        else if (name.includes("Service_Email")) grouped.Service_Email.push(entry);
        else if (name.includes("Advertising")) grouped.Advertising.push(entry);
      }
      return res.status(200).json(grouped);
    }

    // INIT-SLIDES — scan presentation, find images from blob URLs, save mapping
    // Run this ONCE after inserting images from blob URLs into your slides
    if (action === "init-slides") {
      if (!slidesId || !saKey) return res.status(400).json({ error: "GOOGLE_SLIDES_PRESENTATION_ID and GOOGLE_SERVICE_ACCOUNT_KEY required." });
      const sa = JSON.parse(saKey);
      const token = await getAccessToken(sa);
      const { blobs } = await list({ prefix: "screenshots/" });
      const fileNames = blobs.map(b => b.pathname.replace("screenshots/", "")).filter(n => !n.includes("_2026-") && !n.includes("_2025-"));
      const { mapping, debug } = await buildMappingFromPresentation(token, slidesId, fileNames);
      const count = Object.keys(mapping).length;
      if (count === 0) {
        return res.status(200).json({
          success: false, mapped: 0, debug: debug.slice(0, 20),
          message: "Could not match images to screenshots. Check debug to see what text is on each slide.",
        });
      }
      await saveMapping(mapping);
      return res.status(200).json({
        success: true, mapped: count, mapping,
        message: `Locked in ${count} image mappings. Future refreshes will use these IDs.`,
      });
    }

    // REFRESH — replace images in the presentation
    if (action === "refresh") {
      if (!slidesId || !saKey) return res.status(400).json({ error: "GOOGLE_SLIDES_PRESENTATION_ID and GOOGLE_SERVICE_ACCOUNT_KEY required." });
      const sa = JSON.parse(saKey);
      const token = await getAccessToken(sa);

      // Get all screenshot URLs
      const { blobs } = await list({ prefix: "screenshots/" });
      const urlMap = {};
      for (const b of blobs) {
        const name = b.pathname.replace("screenshots/", "");
        if (!name.includes("_2026-") && !name.includes("_2025-")) { // skip old dated files
          urlMap[name] = b.url;
        }
      }

      // Try saved mapping first (fast, reliable)
      let mapping = await loadMapping();
      if (mapping && Object.keys(mapping).length > 0) {
        const result = await replaceByMapping(token, slidesId, mapping, urlMap);
        return res.status(200).json({ ...result, method: "saved-mapping" });
      }

      // No saved mapping — try to build one from current presentation
      const fileNames = Object.keys(urlMap);
      const { mapping: newMapping } = await buildMappingFromPresentation(token, slidesId, fileNames);
      if (Object.keys(newMapping).length > 0) {
        await saveMapping(newMapping);
        const result = await replaceByMapping(token, slidesId, newMapping, urlMap);
        return res.status(200).json({ ...result, method: "auto-detected", note: "Mapping saved for future use." });
      }

      return res.status(200).json({
        replaced: 0, method: "none",
        message: "No mapping found and no blob URLs detected in the presentation. Run ?action=init-slides after inserting images from blob URLs.",
      });
    }

    // MAPPING — view or clear the saved mapping
    if (action === "mapping") {
      const mapping = await loadMapping();
      return res.status(200).json({ exists: !!mapping, count: mapping ? Object.keys(mapping).length : 0, mapping });
    }

    if (action === "clear-mapping") {
      try { const { blobs } = await list({ prefix: MAPPING_KEY }); for (const b of blobs) await del(b.url); } catch {}
      return res.status(200).json({ message: "Mapping cleared. Run init-slides to rebuild." });
    }

    return res.status(400).json({ error: "Actions: test, upload, list, urls, init-slides, refresh, mapping, clear-mapping" });
  } catch (err) {
    console.error("[publish.js]", err);
    return res.status(500).json({ error: err.message });
  }
}
