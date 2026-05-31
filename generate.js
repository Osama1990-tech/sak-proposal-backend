const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, PageBreak, LevelFormat, Header, Footer, SimpleField,

} = require('docx');
const fs = require('fs');
const path = require('path');

// ── DATA (passed via JSON arg or use sample) ────────────────────────────────
const dataArg = process.argv[2];
const DATA = dataArg ? JSON.parse(dataArg) : {
  ref: "DMM-BLOCK12-BID-2026-326-R01",
  date: "13TH MAY. 2026",
  rev: "R01",
  client_name: "Abdul Aziz Al Ajlan Sons for Trading & Real Estate Investments Co.",
  client_address: "2nd Industrial City, Dammam",
  client_attn: "AMR KAMEL",
  client_phone: "+966 56 170 4050",
  proj_scope: "Construction of Warehouse (Scope 04)",
  proj_block: "Block 12",
  proj_location: "Scoop Industrial City, Dammam",
  contract_type: "LUMPSUM",
  duration: "12",
  duration_unit: "Months",
  pay_advance: "20",
  pay_progress: "75",
  pay_retention: "5",
  validity: "15",
  base_price: 0,
  additions: 0,
  omissions: 0,
  missed: 0,
  grand_total: 0,
  extras_desc: "",
  extras_amount: 0,
  sprinkler_delta: 0,
  price_rows: [
    { div: "DIV 01", desc: "General Requirements", amount: 0 },
    { div: "DIV 03", desc: "Concrete", amount: 11908252.66 },
    { div: "DIV 04", desc: "Masonry", amount: 5710780.00 },
    { div: "DIV 05", desc: "Metals", amount: 11175293.60 },
    { div: "DIV 07", desc: "Thermal & Moisture Protection", amount: 2117684.00 },
    { div: "DIV 08", desc: "Openings", amount: 1992362.00 },
    { div: "DIV 09", desc: "Finishing", amount: 1501471.63 },
    { div: "DIV 21", desc: "Fire Suppression", amount: 4368894.00 },
    { div: "DIV 22", desc: "Plumbing", amount: 411154.00 },
    { div: "DIV 23", desc: "HVAC", amount: 864951.00 },
    { div: "DIV 26", desc: "Electrical", amount: 3832133.00 },
    { div: "DIV 31", desc: "Earth Work", amount: 5485000.00 },
    { div: "DIV 32", desc: "Exterior Improvement", amount: 1661475.00 }
  ],
  assumptions: [
    "Mobilization: 1 No. Portable office for Client & Consultant including suitable furniture, AC unit, and MEP services.",
    "Mobilization: 1 No. Portable for Contractor as per our standard.",
    "Any item not mentioned excluded from contractor scope of Mobilization.",
    "Our offer is based on the vendor list attached to the offer. Soroh Al-Khair has the right to choose the suitable supplier from this list.",
    "Any additional requirement not included in our offer is Excluded and will be additional charges."
  ],
  exclusions: [
    "Any items not mentioned in BOQ.",
    "Supply, Unloading and Installation of Overhead Cranes if any.",
    "Supply, Unloading, and Installation of Elevator if any.",
    "Generator Set / SCECO RMU / Switchgear.",
    "Furniture.",
    "External Interlock & Curbstone outside the perimeter of Lot boundary line.",
    "Building Permit fees and any other Authority fees.",
    "Any type of insurance by client.",
    "Any Government Fees.",
    "Coordination with Governmental Authorities.",
    "Removal & Re-routing of existing MEP services, if any.",
    "Demolition works not mentioned in BOQ.",
    "Any Dewatering and Shoring works excluded from scope.",
    "Inherent Defect Insurance (Tawuniya) excluded from scope.",
    "Fireproof paint for steel structure building excluded from scope."
  ],
  vendor_arch: [
    ["PEB + Sandwich Panel", "KIRBY", "Al Zamil", "ZAMIL"],
    ["SFRC Sub-contractors", "DAC (CBS)", "CONTROTEC", "-"],
    ["Steel Reinforcement", "Al Jazeera Steel", "Saudi National Steel (SNS)", "Al Nuaimi"],
    ["Supply Concrete", "Riyadh Kingdom (RK)", "Al Naqoul", "-"],
    ["Masonry Block", "Al Naqoul", "SIBC", "-"],
    ["Painting", "Jotun", "Al Jazeera", "Al Jazera"],
    ["Ceramic & Porcelain Tile", "Future Ceramic", "Al Jawadah Ceramic", "Saudi Ceramic"],
    ["False Ceiling", "USG Boral", "Bamar", "-"],
    ["Aluminum & Glass Doors & Windows", "Saraya Section", "Royal Section", "Al Manara"],
    ["Galvanized Steel Rollup Doors", "Petronet", "Randa", "-"],
    ["Metal / Fire Rated Doors", "Al Thabet", "Al Barak", "-"],
    ["Epoxy Floor Paint", "Sika Floor 264 SG - 500 Micron", "Fosroc-Nitoflor Fc150", "Al Jazeera Factory"]
  ],
  vendor_mech: [
    ["Water Pump", "Samnan", "Grand", "KSB"],
    ["Sewer & Drainage Pipe Line", "Bahrain Pipe Fab", "Fabco", "Nepro"],
    ["Water Supply Pipe", "Bahrain Pipe", "KTP", "API"],
    ["AC Units", "Midea", "Hisense", "-"],
    ["Package Unit & VRF System", "Midea", "Craf", "-"],
    ["Exhaust Fans - Industrial", "Saudi Fan", "Rosenberg", "Zamil"],
    ["Fire Pump", "Al Nahdi (ANC)", "SMI", "Al Nahdi (ANC)"],
    ["FM200 System", "Safety Hi Tech", "Fireless-227", "-"],
    ["Fire Hose Cabinet", "HEBA", "SFFECO", "-"],
    ["Sprinkler", "Reliable", "Lifco", "-"],
    ["HDPE PN16", "Al-Jubail Sanitary Pipe Factory", "Al Wasail", "-"]
  ],
  vendor_elec: [
    ["Lighting Fixtures", "NLC", "Optimal", "KYD"],
    ["Wires and Cables", "Bahra", "Elsewedy", "Riyadh"],
    ["Wiring Devices", "Legrand", "Belanko S M", "Alfanar"],
    ["EMT Conduit", "Alroof", "Tablasa", "ITTC"],
    ["Panel Boards (Ready Built)", "ABB", "Schneider", "Fabco"],
    ["Circuit Breakers", "ABB", "Schneider", "-"],
    ["CCTV System", "Hikvision", "Dahua", "Hickvision"],
    ["Fire Alarm System", "Fire Light", "Fire Sole", "-"],
    ["Low Current Cables", "Belden", "3M", "-"],
    ["Switches", "D-Link", "Hickvision", "Linksys"]
  ]
};

// ── HELPERS ──────────────────────────────────────────────────────────────────
const SAK_BLUE = "1A3A6B";
const SAK_LIGHT = "E8EDF5";
const SAK_GRAY = "F5F5F5";
const RED = "C00000";
const WHITE = "FFFFFF";

function fmt(n) {
  return Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const allBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const noBorder = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function cell(text, opts = {}) {
  const {
    bold = false, color = "000000", bg = null, span = 1,
    w = 2000, align = AlignmentType.LEFT, size = 20, italic = false,
    valign = VerticalAlign.CENTER, borders = allBorders
  } = opts;
  return new TableCell({
    columnSpan: span,
    width: { size: w, type: WidthType.DXA },
    verticalAlign: valign,
    borders,
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, color, size, italics: italic, font: "Arial" })]
    })]
  });
}

function hCell(text, w = 2000, span = 1) {
  return cell(text, { bold: true, color: WHITE, bg: SAK_BLUE, w, span, size: 19 });
}

function gap(spBefore = 80, spAfter = 80) {
  return new Paragraph({ spacing: { before: spBefore, after: spAfter }, children: [] });
}

function heading(text, level = 1) {
  const sizes = { 1: 28, 2: 24, 3: 22 };
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, size: sizes[level] || 22, color: SAK_BLUE, font: "Arial" })]
  });
}

function para(text, opts = {}) {
  const { bold = false, size = 20, color = "000000", align = AlignmentType.LEFT, before = 40, after = 40 } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before, after },
    children: [new TextRun({ text, bold, size, color, font: "Arial" })]
  });
}

function dividerLine() {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: SAK_BLUE, space: 1 } },
    children: []
  });
}

function numberedList(items, prefix = "") {
  return items.map((item, i) =>
    new Paragraph({
      spacing: { before: 30, after: 30 },
      indent: { left: 360, hanging: 280 },
      children: [new TextRun({ text: `${i + 1}. ${prefix}${item}`, size: 20, font: "Arial" })]
    })
  );
}

// ── COVER PAGE ────────────────────────────────────────────────────────────────
function makeCoverSection() {
  const scope = `${DATA.proj_scope} – ${DATA.proj_block}`;
  return [
    gap(600),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: "CONSTRUCTION OF", bold: true, size: 32, color: SAK_BLUE, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: scope, bold: true, size: 40, color: "000000", font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: DATA.proj_location, size: 24, color: "555555", font: "Arial" })] }),
    gap(400),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: `TECHNO-COMMERCIAL PROPOSAL - ${DATA.rev}`, bold: true, size: 28, color: "000000", font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: DATA.date.toUpperCase(), bold: true, size: 24, color: SAK_BLUE, font: "Arial" })] }),
    gap(600),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ── COVER LETTER ─────────────────────────────────────────────────────────────
function makeCoverLetter() {
  const base = DATA.price_rows.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0);
  const grand = base + (parseFloat(DATA.additions) || 0) - (parseFloat(DATA.omissions) || 0) + (parseFloat(DATA.missed) || 0) + (parseFloat(DATA.extras_amount) || 0);

  const letterTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2200, 7160],
    borders: noBorders,
    rows: [
      new TableRow({ children: [
        cell("Date:", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.date, { w: 7160, borders: noBorders })
      ]}),
      new TableRow({ children: [
        cell("Ref:", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.ref, { w: 7160, borders: noBorders })
      ]}),
      new TableRow({ children: [
        cell("M/S", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.client_name, { bold: true, w: 7160, borders: noBorders })
      ]}),
      new TableRow({ children: [
        cell("Address:", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.client_address, { w: 7160, borders: noBorders })
      ]}),
      new TableRow({ children: [
        cell("Attention:", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.client_attn, { w: 7160, borders: noBorders })
      ]}),
      new TableRow({ children: [
        cell("Phone:", { bold: true, w: 2200, borders: noBorders }),
        cell(DATA.client_phone, { w: 7160, borders: noBorders })
      ]})
    ]
  });

  const summaryRows = [
    new TableRow({ children: [
      hCell("Sl. No", 1800), hCell("Description", 5660), hCell("Total Amount (SAR)", 1900)
    ]}),
    new TableRow({ children: [
      cell("BASE PRICE", { bold: true, bg: SAK_GRAY, w: 1800 }),
      cell(`${DATA.proj_scope} – ${DATA.proj_block}`, { w: 5660 }),
      cell(fmt(base), { w: 1900, align: AlignmentType.RIGHT })
    ]})
  ];

  if (parseFloat(DATA.additions) > 0) {
    summaryRows.push(new TableRow({ children: [
      cell("ADDITIONS_QUANTITIES.", { bold: true, color: SAK_BLUE, bg: SAK_GRAY, w: 1800, span: 2 }),
      cell(fmt(DATA.additions), { w: 1900, align: AlignmentType.RIGHT, bold: true, color: SAK_BLUE })
    ]}));
  }
  if (parseFloat(DATA.omissions) > 0) {
    summaryRows.push(new TableRow({ children: [
      cell("OMISSIONS_QUANTITIES.", { bold: true, color: SAK_BLUE, bg: SAK_GRAY, w: 1800, span: 2 }),
      cell(fmt(DATA.omissions), { w: 1900, align: AlignmentType.RIGHT, bold: true, color: SAK_BLUE })
    ]}));
  }
  if (parseFloat(DATA.missed) > 0) {
    summaryRows.push(new TableRow({ children: [
      cell("MISSED ITEMS.", { bold: true, color: SAK_BLUE, bg: SAK_GRAY, w: 1800, span: 2 }),
      cell(fmt(DATA.missed), { w: 1900, align: AlignmentType.RIGHT, bold: true, color: SAK_BLUE })
    ]}));
  }
  summaryRows.push(new TableRow({ children: [
    cell("TOTAL AMOUNT - (EXCLUDED VAT)", { bold: true, bg: SAK_BLUE, color: WHITE, w: 1800, span: 2 }),
    cell(fmt(grand), { bold: true, bg: SAK_BLUE, color: WHITE, w: 1900, align: AlignmentType.RIGHT })
  ]}));

  if (DATA.extras_desc && parseFloat(DATA.extras_amount) > 0) {
    summaryRows.push(new TableRow({ children: [
      cell(DATA.extras_desc, { w: 1800, span: 2 }),
      cell(fmt(DATA.extras_amount), { w: 1900, align: AlignmentType.RIGHT })
    ]}));
  }
  if (parseFloat(DATA.sprinkler_delta) > 0) {
    summaryRows.push(new TableRow({ children: [
      cell("Estimated cost impact resulting from changing the sprinkler system key factor to K16.8.", { w: 1800, span: 2 }),
      cell(fmt(DATA.sprinkler_delta), { w: 1900, align: AlignmentType.RIGHT })
    ]}));
  }

  const summaryTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1800, 5660, 1900],
    rows: summaryRows
  });

  const appendixList = [
    "Appendix 1 – Terms and Conditions.",
    "Appendix 2 – Assumptions and Exclusions.",
    "Appendix 3 – Proposed Vendor List.",
    "Appendix 4 – Priced Bill of Quantities (BOQ)."
  ].map(t => new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: { left: 360 },
    children: [new TextRun({ text: `\u2666 ${t}`, size: 20, font: "Arial" })]
  }));

  return [
    gap(80),
    letterTable,
    gap(100),
    dividerLine(),
    gap(60),
    para(`THE SCOPE: ${DATA.proj_scope} – ${DATA.proj_block}.`, { bold: true }),
    para(`SUBJECT: "Techno-Commercial Proposal-${DATA.rev}"`, { bold: true }),
    gap(60),
    para("Dear Sir,"),
    gap(40),
    new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [
        new TextRun({ text: "Soroh Al-Khair General Contracting would like to thank ", size: 20, font: "Arial" }),
        new TextRun({ text: `M/S ${DATA.client_name}`, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: " for the opportunity extended to us in tendering of the above-mentioned project. In this regard, we have the pleasure to submit our ", size: 20, font: "Arial" }),
        new TextRun({ text: "Techno-Commercial Proposal", bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: " for the construction of the Works as follows:", size: 20, font: "Arial" })
      ]
    }),
    gap(80),
    summaryTable,
    gap(100),
    new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [
        new TextRun({ text: "Validity: ", bold: true, italics: true, size: 20, font: "Arial", underline: {} }),
        new TextRun({ text: "This proposal is valid for a period of ", size: 20, font: "Arial" }),
        new TextRun({ text: `${DATA.validity} days`, bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: " from the date of this proposal.", size: 20, font: "Arial" })
      ]
    }),
    para("Enclosed herewith the following appendices which are integral part of this proposal."),
    gap(60),
    ...appendixList,
    gap(120),
    para("For and on behalf of", { bold: true }),
    para("Soroh Al-Khair General Contracting", { bold: true }),
    gap(60),
    para("HAMAD AL-HAJIRI", { bold: true }),
    para("GENERAL MANAGER"),
    para("E: h.alhajri@sakcontracting.com"),
    gap(80),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ── APPENDIX 01 — TERMS ───────────────────────────────────────────────────────
function makeAppendix01() {
  const ct = DATA.contract_type || "LUMPSUM";
  const dur = `${DATA.duration} ${DATA.duration_unit}`;
  return [
    heading("APPENDIX-01"),
    heading("TERMS & CONDITIONS", 2),
    dividerLine(), gap(80),
    heading("A. CONTRACT TYPE:", 3),
    para(`The contract shall be ${ct}.`),
    gap(80),
    heading("B. TIME FOR COMPLETION:", 3),
    new Paragraph({ spacing: { before: 40, after: 40 }, children: [
      new TextRun({ text: "Contractor shall require a total unobstructed duration of ", size: 20, font: "Arial" }),
      new TextRun({ text: dur, bold: true, size: 20, color: RED, font: "Arial" }),
      new TextRun({ text: " from commencement date for building completion and handover.", size: 20, font: "Arial" })
    ]}),
    gap(80),
    heading("C. COMMENCEMENT DATE:", 3),
    para("The Commencement Date shall be defined as the latest of the following events:", { bold: true }),
    gap(40),
    ...["The date upon which the Contractor obtains the Unconditional Building Permit enabling to proceed with the Works.",
       "The date upon which the Contractor obtains possession and access to the Site and the Site is clear from all existing utility services and similar obstructions.",
       "The date upon which the Contractor receives officially the signed Contract from the Employer.",
       "The date upon which the Contractor receives the agreed Advance payment.",
       "The date upon receiving the approved IFC drawings from client and Consultant."
    ].map((t, i) => new Paragraph({
      spacing: { before: 30, after: 30 },
      indent: { left: 360, hanging: 280 },
      children: [new TextRun({ text: `${i + 1}. ${t}`, size: 20, font: "Arial" })]
    })),
    gap(80),
    heading("D. COMMERCIAL TERMS & CONDITIONS:", 3),
    new Paragraph({ spacing: { before: 40, after: 20 }, children: [
      new TextRun({ text: "Advance Payment: ", bold: true, size: 20, font: "Arial" })
    ]}),
    para(`EMPLOYER shall pay to the Contractor (in one installment) an advance payment equivalent (${DATA.pay_advance}%) of the Accepted Contract Amount.`),
    gap(40),
    new Paragraph({ spacing: { before: 40, after: 20 }, children: [
      new TextRun({ text: "Payment:", bold: true, size: 20, font: "Arial" })
    ]}),
    para(`${DATA.pay_advance}% Advance Payment.`),
    para(`${DATA.pay_progress}% Progress Invoice as per site condition.`),
    para(`${DATA.pay_retention}% Retention.`),
    gap(80),
    heading("E. VALUE-ADDED TAX (VAT):", 3),
    para("Any Consideration or sum payable under this proposal Value is exclusive of Value Added Tax (VAT)."),
    new Paragraph({ spacing: { before: 40, after: 40 }, children: [
      new TextRun({ text: "If VAT", bold: true, size: 20, font: "Arial" }),
      new TextRun({ text: " is chargeable by sum or part thereof payable for a supply made pursuant to this proposal, the amount of ", size: 20, font: "Arial" }),
      new TextRun({ text: "VAT", bold: true, size: 20, font: "Arial" }),
      new TextRun({ text: " chargeable by law shall be payable in addition to the sum or relevant part thereof, subject to issuance of a valid VAT invoice.", size: 20, font: "Arial" })
    ]}),
    gap(60),
    para("All other Terms & Conditions including Payment Terms shall be discussed and mutually agreed prior to the contract award.", { italic: true }),
    gap(80),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ── APPENDIX 02 — ASSUMPTIONS & EXCLUSIONS ───────────────────────────────────
function makeAppendix02() {
  return [
    heading("APPENDIX-02"),
    heading("ASSUMPTIONS & EXCLUSIONS", 2),
    dividerLine(), gap(80),
    heading("ASSUMPTIONS:", 3),
    gap(40),
    ...numberedList(DATA.assumptions),
    gap(120),
    heading("EXCLUSIONS:", 3),
    gap(40),
    ...numberedList(DATA.exclusions),
    gap(80),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ── APPENDIX 03 — VENDOR LIST ─────────────────────────────────────────────────
function makeVendorTable(rows, colW) {
  const header = new TableRow({ children: [
    hCell("Materials", colW[0]),
    hCell("1st Supplier", colW[1]),
    hCell("2nd Supplier", colW[2]),
    hCell("3rd Supplier", colW[3])
  ]});
  const dataRows = rows.map((r, idx) => new TableRow({
    children: r.map((v, ci) => cell(v || "-", {
      w: colW[ci],
      bg: idx % 2 === 0 ? WHITE : SAK_GRAY,
      size: 18
    }))
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colW,
    rows: [header, ...dataRows]
  });
}

function makeAppendix03() {
  const colW = [2800, 2200, 2200, 2160];
  const noteText = "ANY MATERIALS LISTED IN THE VENDOR LIST THAT FALL OUTSIDE SAK'S SCOPE OF WORK WILL BE DEEMED INVALID. SAK IS ENTITLED TO SELECT ANY OF THE MENTIONED VENDORS REGARDLESS OF THEIR SEQUENTIAL ORDER OR RANKING.";
  return [
    heading("APPENDIX-03"),
    heading("PROPOSED VENDOR LIST", 2),
    dividerLine(), gap(80),
    heading("ARCHITECTURAL & STRUCTURAL", 3),
    gap(40),
    makeVendorTable(DATA.vendor_arch, colW),
    gap(120),
    heading("MECHANICAL (HVAC + PLUMBING + FIRE FIGHTING)", 3),
    gap(40),
    makeVendorTable(DATA.vendor_mech, colW),
    gap(120),
    heading("ELECTRICAL WORKS", 3),
    gap(40),
    makeVendorTable(DATA.vendor_elec, colW),
    gap(80),
    new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: `Notes: ${noteText}`, size: 16, color: "666666", italics: true, font: "Arial" })]
    }),
    gap(80),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ── APPENDIX 04 — BOQ SUMMARY ─────────────────────────────────────────────────
function makeAppendix04() {
  const base = DATA.price_rows.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0);
  const add = parseFloat(DATA.additions) || 0;
  const omit = parseFloat(DATA.omissions) || 0;
  const miss = parseFloat(DATA.missed) || 0;
  const ext = parseFloat(DATA.extras_amount) || 0;
  const grand = base + add - omit + miss + ext;

  const colW = [1200, 6160, 2000];
  const summaryHeader = new TableRow({ children: [
    hCell("Item", colW[0]), hCell("Description", colW[1]), hCell("Amount (SAR)", colW[2])
  ]});
  const divRows = DATA.price_rows.map((r, idx) => new TableRow({
    children: [
      cell(r.div, { w: colW[0], size: 18, bg: idx % 2 === 0 ? WHITE : SAK_GRAY }),
      cell(r.desc, { w: colW[1], size: 18, bg: idx % 2 === 0 ? WHITE : SAK_GRAY }),
      cell(fmt(r.amount), { w: colW[2], align: AlignmentType.RIGHT, size: 18, bg: idx % 2 === 0 ? WHITE : SAK_GRAY })
    ]
  }));

  const totalRows = [
    new TableRow({ children: [
      cell("TOTAL SUMMARY SAR", { bold: true, bg: SAK_BLUE, color: WHITE, w: colW[0] + colW[1], span: 2 }),
      cell(fmt(base), { bold: true, bg: SAK_BLUE, color: WHITE, w: colW[2], align: AlignmentType.RIGHT })
    ]})
  ];
  if (add > 0) totalRows.push(new TableRow({ children: [
    cell("ADDITIONS", { bold: true, color: SAK_BLUE, bg: SAK_GRAY, w: colW[0] + colW[1], span: 2 }),
    cell(fmt(add), { bold: true, color: SAK_BLUE, w: colW[2], align: AlignmentType.RIGHT })
  ]}));
  if (omit > 0) totalRows.push(new TableRow({ children: [
    cell("OMISSIONS", { bold: true, color: RED, bg: SAK_GRAY, w: colW[0] + colW[1], span: 2 }),
    cell(fmt(omit), { bold: true, color: RED, w: colW[2], align: AlignmentType.RIGHT })
  ]}));
  if (miss > 0) totalRows.push(new TableRow({ children: [
    cell("MISSED ITEMS", { bold: true, color: SAK_BLUE, bg: SAK_GRAY, w: colW[0] + colW[1], span: 2 }),
    cell(fmt(miss), { bold: true, color: SAK_BLUE, w: colW[2], align: AlignmentType.RIGHT })
  ]}));
  totalRows.push(new TableRow({ children: [
    cell("TOTAL SUMMARY SAR", { bold: true, bg: SAK_BLUE, color: WHITE, w: colW[0] + colW[1], span: 2 }),
    cell(fmt(grand), { bold: true, bg: SAK_BLUE, color: WHITE, w: colW[2], align: AlignmentType.RIGHT })
  ]}));

  return [
    heading("APPENDIX-04"),
    heading("PRICED BILL OF QUANTITIES (BOQ) — SUMMARY", 2),
    dividerLine(), gap(80),
    para(`${DATA.client_name}`, { bold: true }),
    para(`${DATA.proj_location}`),
    para(`${DATA.proj_scope} — ${DATA.proj_block}`, { bold: true }),
    gap(60),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: colW,
      rows: [summaryHeader, ...divRows, ...totalRows]
    }),
    gap(80),
    para("Notes:", { bold: true }),
    para("• This BOQ Summary reflects the division-level totals. The detailed Priced Bill of Quantities is provided as a separate attachment."),
    para("• All prices are in Saudi Arabian Riyals (SAR), exclusive of VAT."),
    para("• Rates are fixed for the duration of the contract and not subject to adjustment.")
  ];
}

// ── HEADER / FOOTER ───────────────────────────────────────────────────────────
function makeHeader() {
  return new Header({
    children: [
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [5500, 3860],
        borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 6, color: SAK_BLUE, space: 1 } },
        rows: [new TableRow({ children: [
          new TableCell({
            borders: noBorders,
            width: { size: 5500, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: "Soroh Al-Khair General Contracting", bold: true, size: 20, color: SAK_BLUE, font: "Arial" })]
            })]
          }),
          new TableCell({
            borders: noBorders,
            width: { size: 3860, type: WidthType.DXA },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: `Ref: ${DATA.ref}`, size: 16, color: "666666", font: "Arial" })]
            })]
          })
        ]})]
      })
    ]
  });
}

function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
        spacing: { before: 60 },
        children: [
          new TextRun({ text: "Bandariyah Center, Prince Faisal Bin Fahad Street, Bandariyah District, Al Khobar, KSA  |  C.R: 2052002095  |  www.sakcontracting.com", size: 16, color: "888888", font: "Arial" }),
          new TextRun({ text: "    Page ", size: 16, color: "888888", font: "Arial" }),
          new SimpleField("PAGE")
        ]
      })
    ]
  });
}

// ── BUILD DOCUMENT ────────────────────────────────────────────────────────────
async function buildDoc() {
  const children = [
    ...makeCoverSection(),
    ...makeCoverLetter(),
    ...makeAppendix01(),
    ...makeAppendix02(),
    ...makeAppendix03(),
    ...makeAppendix04()
  ];

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 20 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
        }
      },
      headers: { default: makeHeader() },
      footers: { default: makeFooter() },
      children
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const outDir = process.env.SAK_OUTDIR || __dirname;
  const outPath = require('path').join(outDir, 'SAK_Proposal_' + DATA.proj_block.replace(/\s+/g,'_') + '_' + DATA.rev + '.docx');
  fs.writeFileSync(outPath, buffer);
  console.log("DOCX:" + outPath);
  return outPath;
}

buildDoc().catch(e => { console.error(e); process.exit(1); });
