// Vercel Serverless Function — /api/publish.js
// Uploads screenshots to Google Drive and refreshes images in Google Slides
// Uses a Google Cloud Service Account for authentication

import crypto from "crypto";

// ── Auth: Generate access token from service account JSON ──
async function getAccessToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: sa.client_email,
    scope: "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/presentations",
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
  if (!data.access_token) throw new Error("Auth failed: " + JSON.stringify(data));
  return data.access_token;
}

// ── Drive: Find existing file by name in folder ──
async function findFile(token, folderId, fileName) {
  const q = encodeURIComponent(`'${folderId}' in parents and name='${fileName}' and trashed=false`);
  const resp = await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await resp.json();
  return data.files && data.files.length > 0 ? data.files[0].id : null;
}

// ── Drive: Upload or overwrite a file ──
async function uploadFile(token, folderId, fileName, imageBuffer, mimeType) {
  const existingId = await findFile(token, folderId, fileName);

  if (existingId) {
    // Overwrite existing file content
    const resp = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${existingId}?uploadType=media`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": mimeType },
      body: imageBuffer,
    });
    if (!resp.ok) throw new Error("Drive update failed: " + (await resp.text()));
    return existingId;
  } else {
    // Create new file with multipart upload
    const metadata = JSON.stringify({ name: fileName, parents: [folderId] });
    const boundary = "----FormBoundary" + Date.now();
    const body = Buffer.concat([
      Buffer.from(
        `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`
      ),
      imageBuffer,
      Buffer.from(`\r\n--${boundary}--`),
    ]);
    const resp = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
      body,
    });
    if (!resp.ok) throw new Error("Drive create failed: " + (await resp.text()));
    const data = await resp.json();

    // Make file publicly viewable
    await fetch(`https://www.googleapis.com/drive/v3/files/${data.id}/permissions`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ role: "reader", type: "anyone" }),
    });

    return data.id;
  }
}

// ── Slides: Refresh all images that come from our Drive folder ──
async function refreshSlides(token, presentationId, fileMap) {
  // Get the presentation
  const resp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error("Slides API error: " + (await resp.text()));
  const pres = await resp.json();

  // Build a set of our Drive file IDs for matching
  const ourFileIds = new Set(Object.values(fileMap));

  // Find all images in the presentation and match to our files
  const requests = [];
  for (const slide of pres.slides || []) {
    for (const element of slide.pageElements || []) {
      if (element.image && element.image.contentUrl) {
        const url = element.image.contentUrl;
        // Check if this image came from one of our Drive files
        for (const [fileName, fileId] of Object.entries(fileMap)) {
          if (url.includes(fileId)) {
            // Force Slides to re-fetch this image
            requests.push({
              replaceImage: {
                imageObjectId: element.objectId,
                imageReplaceMethod: "CENTER_INSIDE",
                url: `https://drive.google.com/uc?export=download&id=${fileId}&t=${Date.now()}`,
              },
            });
            break;
          }
        }
      }
    }
  }

  if (requests.length === 0) {
    return { replaced: 0, message: "No matching images found in the presentation. Make sure you've inserted the images from the Drive folder into your slides first." };
  }

  // Execute batch update
  const updateResp = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ requests }),
  });
  if (!updateResp.ok) throw new Error("Slides batch update failed: " + (await updateResp.text()));

  return { replaced: requests.length, message: `Successfully refreshed ${requests.length} images in the presentation.` };
}

// ── Main handler ──
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  const saKey = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  const slidesId = process.env.GOOGLE_SLIDES_PRESENTATION_ID;

  if (!saKey) return res.status(500).json({ error: "Missing GOOGLE_SERVICE_ACCOUNT_KEY environment variable." });
  if (!folderId) return res.status(500).json({ error: "Missing GOOGLE_DRIVE_FOLDER_ID environment variable." });

  let sa;
  try { sa = JSON.parse(saKey); } catch (e) {
    return res.status(500).json({ error: "GOOGLE_SERVICE_ACCOUNT_KEY is not valid JSON." });
  }

  const { action } = req.query;

  try {
    const token = await getAccessToken(sa);

    // ACTION: upload — upload a single image to Drive
    if (action === "upload" && req.method === "POST") {
      const { fileName, imageData } = req.body; // imageData is base64
      if (!fileName || !imageData) return res.status(400).json({ error: "fileName and imageData required." });

      const buffer = Buffer.from(imageData, "base64");
      const fileId = await uploadFile(token, folderId, fileName, buffer, "image/jpeg");
      const url = `https://drive.google.com/uc?export=view&id=${fileId}`;

      return res.status(200).json({ fileId, url, fileName });
    }

    // ACTION: upload-batch — upload multiple images at once
    if (action === "upload-batch" && req.method === "POST") {
      const { files } = req.body; // [{fileName, imageData}]
      if (!files || !Array.isArray(files)) return res.status(400).json({ error: "files array required." });

      const results = [];
      const fileMap = {};
      for (const f of files) {
        const buffer = Buffer.from(f.imageData, "base64");
        const fileId = await uploadFile(token, folderId, f.fileName, buffer, "image/jpeg");
        fileMap[f.fileName] = fileId;
        results.push({ fileName: f.fileName, fileId, url: `https://drive.google.com/uc?export=view&id=${fileId}` });
      }

      // Auto-refresh slides if presentation ID is configured
      let slidesResult = null;
      if (slidesId) {
        try {
          slidesResult = await refreshSlides(token, slidesId, fileMap);
        } catch (e) {
          slidesResult = { error: e.message };
        }
      }

      return res.status(200).json({ uploaded: results.length, files: results, slides: slidesResult });
    }

    // ACTION: refresh — manually refresh slides
    if (action === "refresh") {
      if (!slidesId) return res.status(400).json({ error: "GOOGLE_SLIDES_PRESENTATION_ID not configured." });

      // Get all files in our folder to build the map
      const listResp = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${folderId}' in parents and trashed=false`)}&fields=files(id,name)&pageSize=200`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const listData = await listResp.json();
      const fileMap = {};
      for (const f of listData.files || []) { fileMap[f.name] = f.id; }

      const result = await refreshSlides(token, slidesId, fileMap);
      return res.status(200).json(result);
    }

    // ACTION: list — list all files in the Drive folder
    if (action === "list") {
      const listResp = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${folderId}' in parents and trashed=false`)}&fields=files(id,name,modifiedTime)&pageSize=200&orderBy=name`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const listData = await listResp.json();
      const files = (listData.files || []).map((f) => ({
        name: f.name,
        fileId: f.id,
        url: `https://drive.google.com/uc?export=view&id=${f.id}`,
        modified: f.modifiedTime,
      }));
      return res.status(200).json({ files });
    }

    // ACTION: test — verify service account auth and folder access
    if (action === "test") {
      const listResp = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(`'${folderId}' in parents and trashed=false`)}&fields=files(id,name)&pageSize=5`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!listResp.ok) {
        const errText = await listResp.text();
        return res.status(200).json({ success: false, error: `Drive API error (${listResp.status}): ${errText}`, config: { hasSaKey: !!saKey, folderId, slidesId: slidesId || "not set", saEmail: sa.client_email } });
      }
      const listData = await listResp.json();
      return res.status(200).json({ success: true, message: "Service account authenticated and folder accessible.", filesInFolder: (listData.files || []).length, config: { saEmail: sa.client_email, folderId, slidesId: slidesId || "not set" } });
    }

    return res.status(400).json({ error: "Invalid action. Use test, upload, upload-batch, refresh, or list." });
  } catch (err) {
    console.error("[publish.js] Error:", err);
    return res.status(500).json({ error: err.message, stack: err.stack?.split("\n").slice(0, 3) });
  }
}
