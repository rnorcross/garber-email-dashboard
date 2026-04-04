import { put, list, del } from "@vercel/blob";
import crypto from "crypto";

async function getAccessToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: sa.client_email,
    scope: "https://www.googleapis.com/auth/presentations",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };
  const b64 = (obj) => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const unsigned = `${b64(header)}.${b64(payload)}`;
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(unsigned);
  const signature = sign.sign(sa.private_key, "base64url");
  const jwt = `${unsigned}.${signature}`;
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await resp.json();
  if (!data.access_token) throw new Error("Slides auth failed: " + JSON.stringify(data));
  return data.access_token;
}

async function refreshSlides(token, presentationId, urlMap) {
  const resp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error("Slides API error: " + (await resp.text()));
  const pres = await resp.json();
  const requests = [];
  for (const slide of pres.slides || []) {
    for (const element of slide.pageElements || []) {
      if (element.image && element.image.contentUrl) {
        const imgUrl = element.image.contentUrl;
        for (const [fileName, blobUrl] of Object.entries(urlMap)) {
          const nameWithoutExt = fileName.replace(".jpg", "");
          if (imgUrl.includes(nameWithoutExt) || imgUrl.includes(encodeURIComponent(nameWithoutExt))) {
            requests.push({
              replaceImage: {
                imageObjectId: element.objectId,
                imageReplaceMethod: "CENTER_INSIDE",
                url: blobUrl + "?t=" + Date.now(),
              },
            });
            break;
          }
        }
      }
    }
  }
  if (requests.length === 0) {
    return { replaced: 0, message: "No matching images found in the presentation. Insert images from the screenshot URLs into your slides first." };
  }
  const updateResp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ requests }),
  });
  if (!updateResp.ok) throw new Error("Slides update failed: " + (await updateResp.text()));
  return { replaced: requests.length, message: `Refreshed ${requests.length} images in the presentation.` };
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
    if (action === "test") {
      const testBlob = await put("test-connection.txt", "ok", { access: "public", addRandomSuffix: false });
      await del(testBlob.url);
      return res.status(200).json({ success: true, message: "Vercel Blob storage is connected.", slidesConfigured: !!slidesId });
    }

    if (action === "upload" && req.method === "POST") {
      const { fileName, imageData } = req.body;
      if (!fileName || !imageData) return res.status(400).json({ error: "fileName and imageData required." });
      const buffer = Buffer.from(imageData, "base64");
      const blob = await put(`screenshots/${fileName}`, buffer, { access: "public", addRandomSuffix: false, contentType: "image/jpeg" });
      return res.status(200).json({ fileName, url: blob.url });
    }

    if (action === "list") {
      const { blobs } = await list({ prefix: "screenshots/" });
      const files = blobs.map((b) => ({ name: b.pathname.replace("screenshots/", ""), url: b.url, size: b.size }));
      return res.status(200).json({ count: files.length, files });
    }

    if (action === "refresh") {
      if (!slidesId) return res.status(400).json({ error: "GOOGLE_SLIDES_PRESENTATION_ID not set." });
      if (!saKey) return res.status(400).json({ error: "GOOGLE_SERVICE_ACCOUNT_KEY not set." });
      const sa = JSON.parse(saKey);
      const token = await getAccessToken(sa);
      const { blobs } = await list({ prefix: "screenshots/" });
      const urlMap = {};
      for (const b of blobs) { urlMap[b.pathname.replace("screenshots/", "")] = b.url; }
      const result = await refreshSlides(token, slidesId, urlMap);
      return res.status(200).json(result);
    }

    if (action === "urls") {
      const { blobs } = await list({ prefix: "screenshots/" });
      const grouped = { Sales_Email: [], Service_Email: [], Advertising: [] };
      for (const b of blobs) {
        const name = b.pathname.replace("screenshots/", "");
        const entry = { name, url: b.url };
        if (name.includes("Sales_Email")) grouped.Sales_Email.push(entry);
        else if (name.includes("Service_Email")) grouped.Service_Email.push(entry);
        else if (name.includes("Advertising")) grouped.Advertising.push(entry);
      }
      return res.status(200).json(grouped);
    }

    return res.status(400).json({ error: "Invalid action. Use test, upload, list, refresh, or urls." });
  } catch (err) {
    console.error("[publish.js]", err);
    return res.status(500).json({ error: err.message });
  }
}
