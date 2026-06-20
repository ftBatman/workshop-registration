// ============================================================
//  Workshop Registration App - Server (server.js)
//  Stack: Express.js + docx npm package
// ============================================================

const express    = require("express");
const cors       = require("cors");
const bodyParser = require("body-parser");
const fs         = require("fs");
const path       = require("path");

// Import everything we need from the docx package
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  HeadingLevel,
  VerticalAlign,
} = require("docx");

// ── App setup ────────────────────────────────────────────────
const app  = express();
const PORT = 3001;

// Path to the Word document that stores all registrations
const DOC_PATH = path.join(__dirname, "registrations.docx");

// In-memory store – loaded once at startup, kept in sync with the .docx
let participants = [];

// ── Middleware ───────────────────────────────────────────────
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));


// ════════════════════════════════════════════════════════════
//  DOCX HELPERS
// ════════════════════════════════════════════════════════════

/**
 * Build and write the registrations.docx file from the current
 * participants array.  Called after every successful registration.
 */
async function saveDocx() {
  // ── Reusable border style ──
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: border, bottom: border, left: border, right: border };

  // ── Helper: make a styled header cell ──
  const headerCell = (text, width) =>
    new TableCell({
      borders,
      width: { size: width, type: WidthType.DXA },
      shading: { fill: "4F46E5", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text,
              bold: true,
              color: "FFFFFF",
              size: 20,
              font: "Arial",
            }),
          ],
        }),
      ],
    });

  // ── Helper: make a normal data cell ──
  const dataCell = (text, width, shade) =>
    new TableCell({
      borders,
      width: { size: width, type: WidthType.DXA },
      shading: { fill: shade, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [new TextRun({ text: String(text), size: 18, font: "Arial" })],
        }),
      ],
    });

  // Column widths in DXA (total = 9360 for US Letter with 1" margins)
  const COL = [400, 1600, 2000, 1800, 1700, 1860]; // #, Name, Email, Phone, Workshop, Registered On

  // ── Header row ──
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell("#",              COL[0]),
      headerCell("Full Name",      COL[1]),
      headerCell("Email",          COL[2]),
      headerCell("Phone",          COL[3]),
      headerCell("Workshop",       COL[4]),
      headerCell("Registered On",  COL[5]),
    ],
  });

  // ── Data rows ──
  const dataRows = participants.map((p, i) => {
    const shade = i % 2 === 0 ? "F8F7FF" : "FFFFFF";
    return new TableRow({
      children: [
        dataCell(i + 1,                 COL[0], shade),
        dataCell(p.name,                COL[1], shade),
        dataCell(p.email,               COL[2], shade),
        dataCell(p.phone,               COL[3], shade),
        dataCell(p.workshop,            COL[4], shade),
        dataCell(p.registeredAt,        COL[5], shade),
      ],
    });
  });

  // ── Assemble document ──
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 22 } },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 15840, height: 12240 }, // A4 Landscape
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        children: [
          // ── Title ──
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 200 },
            children: [
              new TextRun({
                text: "🎓 Workshop Registration List",
                bold: true,
                size: 36,
                color: "4F46E5",
                font: "Arial",
              }),
            ],
          }),

          // ── Subtitle / meta line ──
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 400 },
            children: [
              new TextRun({
                text: `Total Participants: ${participants.length}   |   Last Updated: ${new Date().toLocaleString()}`,
                size: 18,
                color: "6B7280",
                font: "Arial",
              }),
            ],
          }),

          // ── Table ──
          new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: COL,
            rows: participants.length > 0
              ? [headerRow, ...dataRows]
              : [
                  headerRow,
                  new TableRow({
                    children: [
                      new TableCell({
                        borders,
                        columnSpan: 6,
                        margins: { top: 120, bottom: 120, left: 140, right: 140 },
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: "No registrations yet.",
                                italics: true,
                                color: "9CA3AF",
                                size: 18,
                                font: "Arial",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
          }),

          // ── Footer note ──
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { before: 300 },
            children: [
              new TextRun({
                text: "Generated by Workshop Registration App",
                size: 16,
                color: "9CA3AF",
                italics: true,
                font: "Arial",
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(DOC_PATH, buffer);
}

/**
 * Load existing participants from the .docx file.
 * We store a companion JSON sidecar (registrations.json) for fast
 * in-memory reads – the .docx is the canonical "download" artifact.
 */
const JSON_PATH = path.join(__dirname, "registrations.json");

function loadParticipants() {
  if (fs.existsSync(JSON_PATH)) {
    try {
      participants = JSON.parse(fs.readFileSync(JSON_PATH, "utf8"));
      console.log(`📋 Loaded ${participants.length} participant(s) from disk.`);
    } catch {
      participants = [];
    }
  }
}

function saveJson() {
  fs.writeFileSync(JSON_PATH, JSON.stringify(participants, null, 2));
}


// ════════════════════════════════════════════════════════════
//  ROUTES
// ════════════════════════════════════════════════════════════

// ── POST /register ─────────────────────────────────────────
app.post("/register", async (req, res) => {
  const { name, email, phone, workshop } = req.body;

  // Basic validation
  if (!name || !email || !phone || !workshop) {
    return res.status(400).json({ success: false, message: "All fields are required." });
  }

  // Normalize
  const normEmail = email.trim().toLowerCase();
  const normPhone = phone.trim().replace(/\D/g, ""); // digits only for comparison

  // ── Duplicate check ──
  const duplicate = participants.find(
    (p) =>
      p.email.toLowerCase() === normEmail ||
      p.phone.replace(/\D/g, "") === normPhone
  );

  if (duplicate) {
    const reason =
      duplicate.email.toLowerCase() === normEmail
        ? "email address"
        : "phone number";
    return res.status(409).json({
      success: false,
      message: `A registration with this ${reason} already exists.`,
    });
  }

  // ── Save participant ──
  const participant = {
    id:           Date.now(),
    name:         name.trim(),
    email:        normEmail,
    phone:        phone.trim(),
    workshop:     workshop.trim(),
    registeredAt: new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
  };

  participants.push(participant);
  saveJson();           // fast JSON sidecar
  await saveDocx();     // rebuild the .docx

  return res.status(201).json({
    success: true,
    message: `🎉 Registration successful! Welcome, ${participant.name}!`,
    participant,
  });
});


// ── GET /participants ───────────────────────────────────────
app.get("/participants", (req, res) => {
  res.json({ success: true, count: participants.length, participants });
});


// ── GET /download ───────────────────────────────────────────
app.get("/download", (req, res) => {
  if (!fs.existsSync(DOC_PATH)) {
    return res.status(404).json({ success: false, message: "No document found yet." });
  }
  res.download(DOC_PATH, "registrations.docx");
});


// ════════════════════════════════════════════════════════════
//  STARTUP
// ════════════════════════════════════════════════════════════
async function init() {
  loadParticipants();

  // Always generate a fresh .docx on startup so the file exists
  await saveDocx();
  console.log("📄 registrations.docx ready.");

  app.listen(PORT, () => {
    console.log(`\n✅ Server running at http://localhost:${PORT}\n`);
  });
}

init();
