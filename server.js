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
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: Number(process.env.MAX_UPLOAD_FILE_MB || 15) * 1024 * 1024,
    files: Number(process.env.MAX_UPLOAD_FILES || 10),
  },
});

app.use((req, _res, next) => {
  console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
  next();
});

app.use(express.static(path.join(__dirname)));

function firstDefined(...values) {
  for (const value of values) {
    if (value && String(value).trim()) return String(value).trim();
  }
  return "";
}

function formatField(label, value) {
  if (!value) return `${label}:`;
  return `${label}: ${value}`;
}

function valueOrDefault(value) {
  return value && String(value).trim() ? String(value).trim() : "Not provided";
}

function isLikelyEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function sendQuoteResponse(req, res, statusCode, ok, message, data) {
  const accept = String(req.headers.accept || "");
  if (accept.includes("application/json")) {
    return res.status(statusCode).json({ ok, message, ...(data || {}) });
  }
  return res.status(statusCode).send(message);
}

function safeFilename(value) {
  const cleaned = String(value || "file").replace(/[^\w.\- ]+/g, "_").trim();
  return cleaned || "file";
}

async function saveSubmission({ body, files, lines, pdfBuffer }) {
  const submissionsRoot = path.join(__dirname, "submissions");
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
          ["Groundworks", "groundworks"],
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
  res.sendFile(path.join(__dirname, "agri.html"));
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

app.post("/quote", parseQuoteUploads, async (req, res) => {
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
      formatField("Groundworks", body.groundworks),
      formatField("Other info", body.other_info),
      "",
      formatField("First name", body.first_name),
      formatField("Surname", body.last_name),
      formatField("Email", body.email),
      formatField("Telephone", body.telephone),
      formatField("Return date", body.return_date),
      formatField("Additional requirements", body.client_message),
      formatField("Heard about us", body.hear_about),
      formatField("Marketing consent", body.marketing),
    ];

    const pdfBuffer = await buildQuotePdf(body, req.files || []);
    const pdfFilename = `quote-request-${Date.now()}.pdf`;

    const attachments = (req.files || []).map((file) => ({
      filename: file.originalname,
      content: file.buffer,
      contentType: file.mimetype,
    }));
    attachments.unshift({
      filename: pdfFilename,
      content: pdfBuffer,
      contentType: "application/pdf",
    });

    const subjectPieces = [body.first_name, body.last_name].filter(Boolean);
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
  res.sendFile(path.join(__dirname, "thank-you.html"));
});

app.use((_req, res) => {
  res.status(404).send("Not Found. Please open / and submit the form.");
});

app.use((err, req, res, _next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE") {
      sendQuoteResponse(
        req,
        res,
        413,
        false,
        "One of the uploaded files is too large. Max per file is 15MB."
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
