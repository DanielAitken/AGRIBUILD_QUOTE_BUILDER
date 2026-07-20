const express = require("express");
const multer = require("multer");
const nodemailer = require("nodemailer");
const PDFDocument = require("pdfkit");
const path = require("path");
const https = require("https");
const fs = require("fs/promises");
const crypto = require("crypto");

require("dotenv").config();

const app = express();
app.disable("x-powered-by");
app.set("trust proxy", 1);

const allowedUploads = new Map([
  [".pdf", new Set(["application/pdf"])],
  [".jpg", new Set(["image/jpeg"])],
  [".jpeg", new Set(["image/jpeg"])],
  [".png", new Set(["image/png"])],
]);

const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (_req, file, callback) => {
    const extension = path.extname(file.originalname || "").toLowerCase();
    const allowedTypes = allowedUploads.get(extension);
    if (!allowedTypes || !allowedTypes.has(String(file.mimetype || "").toLowerCase())) {
      const error = new Error("Unsupported upload type");
      error.code = "UNSUPPORTED_UPLOAD_TYPE";
      callback(error);
      return;
    }
    callback(null, true);
  },
  limits: {
    fileSize: Number(process.env.MAX_UPLOAD_FILE_MB || 15) * 1024 * 1024,
    files: Number(process.env.MAX_UPLOAD_FILES || 10),
    fields: 150,
    fieldNameSize: 100,
    fieldSize: 100 * 1024,
    parts: 165,
  },
});

app.use((_req, res, next) => {
  res.setHeader("Content-Security-Policy", [
    "default-src 'self'",
    "base-uri 'self'",
    "connect-src 'self'",
    "font-src 'self' https://fonts.gstatic.com",
    "form-action 'self'",
    "frame-ancestors 'none'",
    "img-src 'self' data: https://krgrp.co.uk",
    "object-src 'none'",
    "script-src 'self' 'unsafe-inline'",
    "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  ].join("; "));
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
  if (_req.secure || _req.get("x-forwarded-proto") === "https") {
    res.setHeader("Strict-Transport-Security", "max-age=31536000; includeSubDomains");
  }
  next();
});

app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

app.use(
  "/assets",
  express.static(path.join(__dirname, "assets"), {
    dotfiles: "deny",
    index: false,
    maxAge: "1d",
  })
);

function firstDefined(...values) {
  for (const value of values) {
    if (value && String(value).trim()) return String(value).trim();
  }
  return "";
}

const submissionsRoot = firstDefined(
  process.env.SUBMISSIONS_DIR,
  path.join(__dirname, "submissions")
);

function formatField(label, value) {
  if (!value) return `${label}:`;
  const text = Array.isArray(value) ? value.join(", ") : String(value);
  return `${label}: ${text}`;
}

function valueOrDefault(value) {
  const text = Array.isArray(value) ? value.join(", ") : String(value || "");
  return text.trim() || "Not provided";
}

function isLikelyEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function sendQuoteResponse(req, res, statusCode, ok, message, data) {
  res.setHeader("Cache-Control", "no-store");
  const accept = String(req.headers.accept || "");
  if (accept.includes("application/json")) {
    return res.status(statusCode).json({ ok, message, ...(data || {}) });
  }
  return res.status(statusCode).send(message);
}

const quoteRateWindowMs = 15 * 60 * 1000;
const quoteRateLimit = Number(process.env.QUOTE_RATE_LIMIT || 20);
const quoteAttempts = new Map();

function limitQuoteRequests(req, res, next) {
  const now = Date.now();
  const key = req.ip || req.socket.remoteAddress || "unknown";
  const previous = quoteAttempts.get(key);
  const state = !previous || now - previous.startedAt >= quoteRateWindowMs
    ? { startedAt: now, count: 0 }
    : previous;
  state.count += 1;
  quoteAttempts.set(key, state);

  if (state.count > quoteRateLimit) {
    const retrySeconds = Math.max(1, Math.ceil((state.startedAt + quoteRateWindowMs - now) / 1000));
    res.setHeader("Retry-After", String(retrySeconds));
    sendQuoteResponse(req, res, 429, false, "Too many quote attempts. Please wait and try again.");
    return;
  }

  if (quoteAttempts.size > 500) {
    for (const [attemptKey, attempt] of quoteAttempts) {
      if (now - attempt.startedAt >= quoteRateWindowMs) quoteAttempts.delete(attemptKey);
    }
  }
  next();
}

function requireSameOrigin(req, res, next) {
  const origin = req.get("origin");
  if (!origin) {
    next();
    return;
  }
  try {
    if (new URL(origin).host === req.get("host")) {
      next();
      return;
    }
  } catch (_error) {
    // Invalid origins are rejected below.
  }
  sendQuoteResponse(req, res, 403, false, "This quote request did not come from the KR AgriBuild form.");
}

function singleLine(value, maxLength = 100) {
  return String(value || "").replace(/[\r\n]+/g, " ").trim().slice(0, maxLength);
}

function safeFilename(value) {
  const cleaned = String(value || "file").replace(/[^\w.\- ]+/g, "_").trim();
  return cleaned || "file";
}

async function saveSubmission({ body, files, lines, pdfBuffer }) {
  const submissionId = `${new Date().toISOString().replace(/[:.]/g, "-")}-${crypto
    .randomBytes(3)
    .toString("hex")}`;
  const submissionDir = path.join(submissionsRoot, submissionId);
  const uploadsDir = path.join(submissionDir, "uploads");

  await fs.mkdir(uploadsDir, { recursive: true });
  await fs.writeFile(path.join(submissionDir, "quote-request.pdf"), pdfBuffer);
  await fs.writeFile(path.join(submissionDir, "answers.txt"), lines.join("\n"), "utf8");
  await fs.writeFile(
    path.join(submissionDir, "form.json"),
    JSON.stringify(body, null, 2),
    "utf8"
  );

  let fileIndex = 0;
  for (const file of files || []) {
    fileIndex += 1;
    const filename = `${String(fileIndex).padStart(2, "0")}-${safeFilename(
      file.originalname
    )}`;
    await fs.writeFile(path.join(uploadsDir, filename), file.buffer);
  }

  return submissionId;
}

function parseJsonSafe(value) {
  try {
    return JSON.parse(value);
  } catch (_err) {
    return null;
  }
}

function httpsRequest({ url, method, headers, body }) {
  return new Promise((resolve, reject) => {
    const request = https.request(
      url,
      { method: method || "GET", headers: headers || {} },
      (response) => {
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => {
          resolve({
            statusCode: response.statusCode || 0,
            body: Buffer.concat(chunks).toString("utf8"),
          });
        });
      }
    );

    request.on("error", reject);
    if (body) request.write(body);
    request.end();
  });
}

function getGraphConfig() {
  const tenantId = firstDefined(process.env.GRAPH_TENANT_ID);
  const clientId = firstDefined(process.env.GRAPH_CLIENT_ID);
  const clientSecret = firstDefined(process.env.GRAPH_CLIENT_SECRET);
  if (!tenantId || !clientId || !clientSecret) return null;
  return { tenantId, clientId, clientSecret };
}

function getSmtpConfig() {
  if (!process.env.SMTP_USER || !process.env.SMTP_PASS) return null;
  const smtpPort = Number(process.env.SMTP_PORT || 587);
  const smtpSecure = process.env.SMTP_SECURE === "true" || smtpPort === 465;
  return {
    host: process.env.SMTP_HOST || "smtp.office365.com",
    port: smtpPort,
    secure: smtpSecure,
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  };
}

async function getGraphAccessToken(config) {
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    config.tenantId
  )}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  }).toString();

  const response = await httpsRequest({
    url: tokenUrl,
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Content-Length": Buffer.byteLength(body),
    },
    body,
  });

  const parsed = parseJsonSafe(response.body) || {};
  if (response.statusCode < 200 || response.statusCode >= 300 || !parsed.access_token) {
    const err = new Error(
      `Graph token request failed (${response.statusCode}): ${response.body}`
    );
    err.userMessage =
      "Microsoft Graph authentication failed. Check GRAPH_TENANT_ID, GRAPH_CLIENT_ID, and GRAPH_CLIENT_SECRET.";
    throw err;
  }

  return parsed.access_token;
}

function toGraphAttachment(attachment) {
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: attachment.filename,
    contentType: attachment.contentType || "application/octet-stream",
    contentBytes: Buffer.from(attachment.content).toString("base64"),
  };
}

async function sendViaGraph(graphConfig, mail) {
  const accessToken = await getGraphAccessToken(graphConfig);
  const endpoint = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(
    mail.fromAddress
  )}/sendMail`;

  const payload = {
    message: {
      subject: mail.subject,
      body: {
        contentType: "Text",
        content: mail.text,
      },
      toRecipients: [{ emailAddress: { address: mail.forwardTo } }],
      attachments: (mail.attachments || []).map(toGraphAttachment),
    },
    saveToSentItems: true,
  };

  if (mail.replyTo) {
    payload.message.replyTo = [{ emailAddress: { address: mail.replyTo } }];
  }

  const body = JSON.stringify(payload);
  const response = await httpsRequest({
    url: endpoint,
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      "Content-Length": Buffer.byteLength(body),
    },
    body,
  });

  if (response.statusCode !== 202) {
    const parsed = parseJsonSafe(response.body);
    const graphMessage =
      parsed &&
      parsed.error &&
      (parsed.error.message || parsed.error.code);
    const err = new Error(
      `Graph sendMail failed (${response.statusCode}): ${graphMessage || response.body}`
    );
    if (response.statusCode === 403) {
      err.userMessage =
        "Graph API permission denied. Grant Mail.Send application permission and admin consent.";
    } else {
      err.userMessage = "Could not send through Microsoft Graph. Check Azure app/mailbox setup.";
    }
    throw err;
  }
}

async function sendViaSmtp(smtpConfig, mail) {
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: smtpConfig.port,
    secure: smtpConfig.secure,
    auth: {
      user: smtpConfig.user,
      pass: smtpConfig.pass,
    },
    tls: { ciphers: "TLSv1.2" },
  });

  await transporter.sendMail({
    from: mail.fromAddress,
    to: mail.forwardTo,
    subject: mail.subject,
    text: mail.text,
    replyTo: mail.replyTo || undefined,
    attachments: mail.attachments,
  });
}

function buildQuotePdf(body, files) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 48 });
    const chunks = [];

    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const submittedAt = new Date().toLocaleString("en-GB", {
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
    });

    doc.fontSize(21).text("AgriBuild Quote Request");
    doc.moveDown(0.25);
    doc
      .fontSize(10)
      .fillColor("#6b7280")
      .text(`Submitted: ${submittedAt}`)
      .fillColor("#111111");
    doc.moveDown(1);

    const sections = [
      {
        title: "Project Details",
        fields: [
          ["Package required", "supply_option"],
          ["Division", "division"],
          ["Proposed use", "proposed_use"],
          ["Building type", "building_type"],
          ["Units", "units"],
          ["Length", "length"],
          ["Width", "width"],
          ["Height", "height"],
          ["Project notes", "project_notes"],
        ],
      },
      {
        title: "Specification",
        fields: [
          ["Steelwork finish", "steelwork_finish"],
          ["Roof material", "roof_material"],
          ["Other roof material details", "roof_material_custom"],
          ["Wall material", "wall_material"],
          ["Other wall material details", "wall_material_custom"],
          ["Heated", "heated"],
          ["Rainwater approach", "rainwater_approach"],
          ["Natural daylight", "daylight_preference"],
          ["Cladding preference", "cladding"],
          ["Door types", "door_types"],
          ["Door details", "door_details"],
          ["Internal fittings", "internal_fittings"],
        ],
      },
      {
        title: "Site & Delivery",
        fields: [
          ["Site postcode", "site_postcode"],
          ["Site address", "site_address"],
          ["Site setting", "site_setting"],
          ["Planning status", "planning_status"],
          ["Timescales", "timescales"],
          ["Additional works required", "groundworks_options"],
          ["Additional works details", "groundworks"],
          ["Other info", "other_info"],
        ],
      },
      {
        title: "Contact Details",
        fields: [
          ["First name", "first_name"],
          ["Surname", "last_name"],
          ["Email", "email"],
          ["Telephone", "telephone"],
          ["Return date", "return_date"],
          ["Additional requirements", "client_message"],
          ["Preferred next step", "preferred_next_step"],
          ["Heard about us", "hear_about"],
          ["Marketing consent", "marketing"],
        ],
      },
    ];

    sections.forEach((section) => {
      doc
        .font("Helvetica-Bold")
        .fontSize(13)
        .fillColor("#1b2f6b")
        .text(section.title);
      doc.moveDown(0.35);

      section.fields.forEach(([label, key]) => {
        const value = valueOrDefault(body[key]);
        doc
          .font("Helvetica-Bold")
          .fontSize(10)
          .fillColor("#111111")
          .text(`${label}: `, { continued: true })
          .font("Helvetica")
          .text(value);
      });

      doc.moveDown(0.9);
    });

    doc
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor("#1b2f6b")
      .text("Uploaded Drawings");
    doc.moveDown(0.35);

    if (!files || files.length === 0) {
      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#111111")
        .text("No files uploaded.");
    } else {
      files.forEach((file, index) => {
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(`${index + 1}. ${file.originalname}`);
      });
    }

    doc.end();
  });
}

app.get("/", (_req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.sendFile(path.join(__dirname, "index.html"));
});

app.get("/health", (_req, res) => {
  res.status(200).json({ ok: true });
});

app.get("/quote", (_req, res) => {
  res.redirect(303, "/");
});

const parseQuoteUploads = (req, res, next) => {
  upload.array("drawings")(req, res, (err) => {
    if (err) return next(err);
    next();
  });
};

function validateUploadedFileSignatures(req, _res, next) {
  const signatures = {
    ".pdf": Buffer.from("%PDF-"),
    ".jpg": Buffer.from([0xff, 0xd8, 0xff]),
    ".jpeg": Buffer.from([0xff, 0xd8, 0xff]),
    ".png": Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  };
  const invalid = (req.files || []).some((file) => {
    const signature = signatures[path.extname(file.originalname || "").toLowerCase()];
    return !signature || !file.buffer || !file.buffer.subarray(0, signature.length).equals(signature);
  });
  if (invalid) {
    const error = new Error("Invalid upload content");
    error.code = "INVALID_UPLOAD_CONTENT";
    next(error);
    return;
  }
  next();
}

app.post("/quote", requireSameOrigin, limitQuoteRequests, parseQuoteUploads, validateUploadedFileSignatures, async (req, res) => {
  try {
    const forwardTo = firstDefined(process.env.FORWARD_TO, process.env.MAIL_TO);
    const fromAddress = firstDefined(
      process.env.APP_MAILBOX,
      process.env.SMTP_FROM,
      process.env.SMTP_USER
    );
    const graphConfig = getGraphConfig();
    const smtpConfig = getSmtpConfig();

    const body = req.body || {};
    const lines = [
      "New AgriBuild quote request",
      "--------------------------------",
      "",
      formatField("Package required", body.supply_option),
      formatField("Division", body.division),
      formatField("Proposed use", body.proposed_use),
      formatField("Building type", body.building_type),
      formatField("Units", body.units),
      formatField("Length", body.length),
      formatField("Width", body.width),
      formatField("Height", body.height),
      formatField("Project notes", body.project_notes),
      "",
      formatField("Steelwork finish", body.steelwork_finish),
      formatField("Roof material", body.roof_material),
      formatField("Other roof material details", body.roof_material_custom),
      formatField("Wall material", body.wall_material),
      formatField("Other wall material details", body.wall_material_custom),
      formatField("Heated", body.heated),
      formatField("Rainwater approach", body.rainwater_approach),
      formatField("Natural daylight", body.daylight_preference),
      formatField("Cladding preference", body.cladding),
      formatField("Door types", body.door_types),
      formatField("Door details", body.door_details),
      formatField("Internal fittings", body.internal_fittings),
      "",
      formatField("Site postcode", body.site_postcode),
      formatField("Site address", body.site_address),
      formatField("Site setting", body.site_setting),
      formatField("Planning status", body.planning_status),
      formatField("Timescales", body.timescales),
      formatField("Additional works required", body.groundworks_options),
      formatField("Additional works details", body.groundworks),
      formatField("Other info", body.other_info),
      "",
      formatField("First name", body.first_name),
      formatField("Surname", body.last_name),
      formatField("Email", body.email),
      formatField("Telephone", body.telephone),
      formatField("Return date", body.return_date),
      formatField("Additional requirements", body.client_message),
      formatField("Preferred next step", body.preferred_next_step),
      formatField("Heard about us", body.hear_about),
      formatField("Marketing consent", body.marketing),
    ];

    const pdfBuffer = await buildQuotePdf(body, req.files || []);
    const pdfFilename = `quote-request-${Date.now()}.pdf`;

    const attachments = (req.files || []).map((file) => ({
      filename: safeFilename(file.originalname),
      content: file.buffer,
      contentType: file.mimetype,
    }));
    attachments.unshift({
      filename: pdfFilename,
      content: pdfBuffer,
      contentType: "application/pdf",
    });

    const subjectPieces = [singleLine(body.first_name), singleLine(body.last_name)].filter(Boolean);
    const subject =
      subjectPieces.length > 0
        ? `New quote request: ${subjectPieces.join(" ")}`
        : "New quote request";

    const mail = {
      fromAddress,
      forwardTo,
      subject,
      text: lines.join("\n"),
      replyTo: isLikelyEmail(body.email) ? body.email : "",
      attachments,
    };
    const submissionId = await saveSubmission({
      body,
      files: req.files || [],
      lines,
      pdfBuffer,
    });

    const hasTransport = Boolean(graphConfig || smtpConfig);
    const hasValidAddresses =
      isLikelyEmail(fromAddress) && isLikelyEmail(forwardTo);

    if (hasTransport && hasValidAddresses) {
      try {
        if (graphConfig) {
          await sendViaGraph(graphConfig, mail);
        } else {
          await sendViaSmtp(smtpConfig, mail);
        }
        sendQuoteResponse(
          req,
          res,
          200,
          true,
          `Thanks! Your quote request has been sent. Reference: ${submissionId}.`,
          { reference: submissionId, emailed: true }
        );
        return;
      } catch (mailErr) {
        console.error("Email send failed (submission saved locally):", mailErr);
        sendQuoteResponse(
          req,
          res,
          200,
          true,
          `Your request was saved successfully (reference: ${submissionId}). Email delivery is currently unavailable.`,
          { reference: submissionId, emailed: false }
        );
        return;
      }
    }

    sendQuoteResponse(
      req,
      res,
      200,
      true,
      `Your request was saved successfully (reference: ${submissionId}).`,
      { reference: submissionId, emailed: false }
    );
  } catch (err) {
    console.error("Submission failed:", err);
    const userMessage =
      err && err.userMessage
        ? err.userMessage
        : "Sorry, something went wrong saving your request.";
    sendQuoteResponse(
      req,
      res,
      500,
      false,
      userMessage
    );
  }
});

app.get("/thank-you", (_req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.sendFile(path.join(__dirname, "thank-you.html"));
});

app.use((_req, res) => {
  res.status(404).send("Not Found. Please open / and submit the form.");
});

app.use((err, req, res, _next) => {
  if (err && err.code === "INVALID_UPLOAD_CONTENT") {
    sendQuoteResponse(req, res, 400, false, "One of the uploaded files does not match its stated file type.");
    return;
  }

  if (err && err.code === "UNSUPPORTED_UPLOAD_TYPE") {
    sendQuoteResponse(
      req,
      res,
      400,
      false,
      "Unsupported drawing type. Please upload PDF, JPG or PNG files only."
    );
    return;
  }

  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE") {
      sendQuoteResponse(
        req,
        res,
        413,
        false,
        `One of the uploaded files is too large. Max per file is ${Number(process.env.MAX_UPLOAD_FILE_MB || 15)}MB.`
      );
      return;
    }
    if (err.code === "LIMIT_FILE_COUNT") {
      sendQuoteResponse(
        req,
        res,
        413,
        false,
        `Too many uploaded files. The maximum is ${Number(process.env.MAX_UPLOAD_FILES || 10)}.`
      );
      return;
    }
    sendQuoteResponse(
      req,
      res,
      400,
      false,
      "Upload failed. Please check your files and try again."
    );
    return;
  }

  console.error("Unhandled error:", err);
  sendQuoteResponse(req, res, 500, false, "Sorry, something went wrong.");
});

const port = Number(process.env.PORT || 3000);
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
