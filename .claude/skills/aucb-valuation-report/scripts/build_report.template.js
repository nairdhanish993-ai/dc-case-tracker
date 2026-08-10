// AUCB valuation report — docx-js template.
//
// This is a WORKING EXAMPLE from a real case (Anoop, Kommady, Alappuzha),
// not a generic template with blanks. To use it for a new case:
//   1. Copy this file into the case's scratch/working directory.
//   2. Edit the field values inside the `rows`/`rows2` sections and the
//      Remarks/Declaration paragraphs below to match the new case's data.
//   3. Point `photoDir` and the `imgCell(...)` calls at the new case's
//      extracted photo files (see SKILL.md for how to pull them out of a
//      site-inspection PDF with `pdfimages`).
//   4. Update the output filename at the bottom.
// Keep the layout/structure (table shapes, section order, TOTAL_W, the
// separate 4-column boundary table) — that part matches the bank's fixed
// format and does not change between cases. See references/aucb-field-list.md
// for the authoritative field list this structure implements.
//
// Run with: cd <dir with this file> && npm install docx && node build_report.template.js

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, HeadingLevel,
  ImageRun, VerticalAlign, PageBreak
} = require("docx");

const TAN = "C9BFA0";
const LIGHTGRAY = "F2F2F2";
const TOTAL_W = 10300; // dxa, fits within A4 usable width (11907 - 2*700 margins = 10507)
const COL_NUM = 500;
const COL_LABEL = 4200;
const COL_VAL = TOTAL_W - COL_NUM - COL_LABEL;

function borderSet() {
  const b = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
  return { top: b, bottom: b, left: b, right: b };
}

function cell(text, {width, bold=false, shade=null, colSpan=1, align=AlignmentType.LEFT, size=19, italic=false, color=null} = {}) {
  const runOpts = { text: String(text), bold, italics: italic, size };
  if (color) runOpts.color = color;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: colSpan,
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    borders: borderSet(),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: [new TextRun(runOpts)] })]
  });
}

function sectionHeaderRow(title) {
  return new TableRow({
    children: [cell(title, {width: TOTAL_W, colSpan: 3, bold: true, shade: TAN, size: 20})]
  });
}

function fieldRow(num, label, value, opts={}) {
  const flagged = opts.flag;
  return new TableRow({
    children: [
      cell(num, {width: COL_NUM, bold: true, align: AlignmentType.CENTER}),
      cell(label, {width: COL_LABEL, bold: false}),
      cell(value, {width: COL_VAL, italic: !!flagged, color: flagged ? "B45309" : null})
    ]
  });
}

function boundaryHeaderRow(title) {
  return new TableRow({
    children: [cell(title, {width: TOTAL_W, colSpan: 4, bold: true, shade: TAN, size: 20})]
  });
}
function boundaryRow(dir1, val1, dir2, val2) {
  const w = TOTAL_W/4;
  return new TableRow({
    children: [
      cell(dir1, {width: w, bold: true}),
      cell(val1, {width: w}),
      cell(dir2, {width: w, bold: true}),
      cell(val2, {width: w})
    ]
  });
}

function plainPara(text, opts={}) {
  return new Paragraph({
    spacing: { after: 120 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, bold: !!opts.bold, size: opts.size || 20, italics: !!opts.italic, color: opts.color || null })]
  });
}

function flagNote(text) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: "⚠ " + text, italics: true, size: 18, color: "B45309" })]
  });
}

// ---------- HEADER TABLE ----------
const headerTable = new Table({
  width: { size: TOTAL_W, type: WidthType.DXA },
  columnWidths: [3200, 4400, 3200],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: {size:3200, type: WidthType.DXA}, borders: borderSet(), shading: {type:ShadingType.CLEAR, fill:TAN},
          verticalAlign: VerticalAlign.CENTER, margins:{top:100,bottom:100,left:120,right:120},
          children: [
            plainPara("Dhanish Nair", {bold:true, size:20}),
            plainPara(""),
            plainPara("IBBI registered valuer", {size:16}),
            plainPara("Chartered Engineer", {size:16}),
            plainPara("Empaneled valuer for Alleppey Urban Co-operative Bank", {size:16}),
          ]
        }),
        new TableCell({
          width: {size:4400, type: WidthType.DXA}, borders: borderSet(),
          verticalAlign: VerticalAlign.CENTER, margins:{top:100,bottom:100,left:120,right:120},
          children: [ new Paragraph({alignment: AlignmentType.CENTER, children:[new TextRun({text:"Valuation Report", bold:true, size:36, color:"1F4E79"})]}) ]
        }),
        new TableCell({
          width: {size:3200, type: WidthType.DXA}, borders: borderSet(),
          verticalAlign: VerticalAlign.CENTER, margins:{top:100,bottom:100,left:120,right:120},
          children: [
            new Paragraph({alignment: AlignmentType.CENTER, children:[new TextRun({text:"AUCB", bold:true, size:24})]}),
            new Paragraph({alignment: AlignmentType.CENTER, children:[new TextRun({text:"Alleppey Urban Co-operative Bank", bold:true, size:16})]}),
          ]
        }),
      ]
    })
  ]
});

// ---------- MAIN DATA TABLE ----------
const rows = [];

rows.push(sectionHeaderRow("Customer Details"));
rows.push(fieldRow(1, "Branch Name", "Not specified — not provided with case papers", {flag:true}));
rows.push(fieldRow(2, "Application Number", "Not specified — not provided with case papers", {flag:true}));
rows.push(fieldRow(3, "Date of site visit", "06-08-2026"));
rows.push(fieldRow(4, "Date of Report", "10-08-2026"));
rows.push(fieldRow(5, "Purpose of Valuation", "To assess the present market value of the property."));
rows.push(fieldRow(6, "Name of Customer", "Anoop"));
rows.push(fieldRow(7, "Contact number of the customer", "Not specified — not provided with case papers", {flag:true}));

rows.push(sectionHeaderRow("Property Details"));
rows.push(fieldRow(1, "Property Address with Pin code", "Poomkavu Church Road, Kommady, Alappuzha, Kerala – 688007 (per site photograph geolocation)"));
rows.push(fieldRow(2, "Legal address", "Alappuzha District (Village/Taluk/Block not confirmed this round — title deed handwriting not reliably legible; recommend confirming from location sketch)", {flag:true}));
rows.push(fieldRow(3, "Landmark", "Not specified in site notes", {flag:true}));
rows.push(fieldRow(4, "Owner Name", "Anoop (as per site inspection file reference — recommend confirming exact registered name(s) against title deed / tax receipts)", {flag:true}));
rows.push(fieldRow(5, "Type of Property", "Residential"));
rows.push(fieldRow(6, "Approved Usage of Property", "Residential"));
rows.push(fieldRow(7, "Actual Use of the property", "Residential"));
rows.push(fieldRow(8, "Property occupancy Details", "Owner Occupied (Self)"));
rows.push(fieldRow(9, "Name of Person Occupying the property", "Anoop"));
rows.push(fieldRow(10, "If Rented, Tenant Name & Rent", "Nil"));
rows.push(fieldRow(11, "Expected/Rent Received", "Nil"));
rows.push(fieldRow(12, "Name & Contact details of person met site", "Not specified in site notes", {flag:true}));
rows.push(fieldRow(13, "Name of House/Door/Society Board", "Not specified in site notes", {flag:true}));
rows.push(fieldRow(14, "Proximity to Civic amenities", "Municipal road on East and South sides (per site boundary notes)"));

rows.push(sectionHeaderRow("Documents provided for valuation"));
rows.push(new TableRow({children:[
  cell("Document", {width: COL_NUM+COL_LABEL, bold:true, shade: LIGHTGRAY}),
  cell("DOC NO / Dated", {width: COL_VAL, bold:true, shade: LIGHTGRAY, align: AlignmentType.CENTER})
]}));
rows.push(new TableRow({children:[cell("Title deed", {width: COL_NUM+COL_LABEL}), cell("Handwritten deed dated 2001; Re-Survey Nos. 387/1/4 & 384/6B referenced — exact document number/date and executant names not reliably legible on the scanned copy provided", {width: COL_VAL, italic:true, color:"B45309"})]}));
rows.push(new TableRow({children:[cell("Land tax receipt", {width: COL_NUM+COL_LABEL}), cell("Not provided", {width: COL_VAL, italic:true, color:"B45309"})]}));
rows.push(new TableRow({children:[cell("Possession certificate", {width: COL_NUM+COL_LABEL}), cell("Not provided", {width: COL_VAL, italic:true, color:"B45309"})]}));
rows.push(new TableRow({children:[cell("Building tax receipt", {width: COL_NUM+COL_LABEL}), cell("Not provided", {width: COL_VAL, italic:true, color:"B45309"})]}));
rows.push(new TableRow({children:[cell("Location sketch", {width: COL_NUM+COL_LABEL}), cell("Not provided", {width: COL_VAL, italic:true, color:"B45309"})]}));

rows.push(sectionHeaderRow("Land and Building Approval Details"));
rows.push(fieldRow(1, "Approved plan provided, Yes/No", "No / Not provided"));
rows.push(fieldRow(2, "Approved plan number & Date", "NA"));
rows.push(fieldRow(3, "Occupancy Certificate available, Y/N", "No / Not provided"));
rows.push(fieldRow(4, "Occupancy Certificate number & Date", "NA"));
rows.push(fieldRow(5, "Approving Authority Name", "NA"));
rows.push(fieldRow(6, "Whether layout plan approved", "Not verifiable — approved plan not furnished"));
rows.push(fieldRow(7, "Land classification", "Not specified in site notes/legible deed portion — recommend confirmation", {flag:true}));

const table1 = new Table({
  width: {size: TOTAL_W, type: WidthType.DXA},
  columnWidths: [COL_NUM, COL_LABEL, COL_VAL],
  rows
});

// ---------- BOUNDARY TABLE (4 columns) ----------
const bRows = [];
bRows.push(boundaryHeaderRow("Boundaries as per Title deed"));
bRows.push(boundaryRow("North", "Not legible / not recorded", "East", "Not legible / not recorded"));
bRows.push(boundaryRow("South", "Not legible / not recorded", "West", "Not legible / not recorded"));

bRows.push(boundaryHeaderRow("Boundaries as per Location sketch"));
bRows.push(boundaryRow("North", "Not provided", "East", "Not provided"));
bRows.push(boundaryRow("South", "Not provided", "West", "Not provided"));

bRows.push(boundaryHeaderRow("Boundaries as per Site"));
bRows.push(boundaryRow("North", "Binu (property of) and pathway", "East", "Municipal way"));
bRows.push(boundaryRow("South", "Municipal way", "West", "Alosius (property of)"));

const boundaryTable = new Table({
  width: {size: TOTAL_W, type: WidthType.DXA},
  columnWidths: [TOTAL_W/4, TOTAL_W/4, TOTAL_W/4, TOTAL_W/4],
  rows: bRows
});

// ---------- TABLE 2: Structural + Valuation ----------
const rows2 = [];
rows2.push(sectionHeaderRow("Structural Details"));
rows2.push(fieldRow(1, "Type of Structure", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(2, "Type of Roofing & terracing", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(3, "Structure Type with flooring", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(4, "Type of Masonry", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(5, "No of Floors", "Not specified in site notes — appears single-storey based on BUA/photographs, recommend confirmation", {flag:true}));
rows2.push(fieldRow(6, "Is setback provision done", "Yes — Front 3m, Side 2m, Rear 21m (as recorded; rear figure appears unusually large compared to other setbacks — recommend field verification)", {flag:true}));
rows2.push(fieldRow(7, "Foundation Type", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(8, "Age of the property", "25 years"));
rows2.push(fieldRow(9, "Estimated future life/Residual life", "Not specified in site notes — recommend confirmation", {flag:true}));
rows2.push(fieldRow(10, "Stage of construction-%", "100% (assumed complete — owner-occupied)", {flag:true}));
rows2.push(fieldRow(11, "Quality of Construction", "Average (inferred from recorded 'Average' condition rating)", {flag:true}));
rows2.push(fieldRow(12, "Appearance & maintenance of the property", "Average"));
rows2.push(fieldRow(13, "Demolition Risk if any with reasons", "Low — no adverse structural concerns reported on site"));
rows2.push(fieldRow(14, "Amenities if any", "Not specified in site notes — assumed Nil", {flag:true}));

rows2.push(sectionHeaderRow("Valuation Method"));
rows2.push(fieldRow(1, "Valuation method", "Land and Building (Market approach)"));
rows2.push(fieldRow(2, "Land area", "6.17 cents"));
rows2.push(fieldRow(3, "Land value/cent", "Rs. 2.5 to 3 lakhs/cent (ongoing rate per site enquiry)"));
rows2.push(fieldRow(4, "Land value considered", "Rs. 2.5 lakhs/cent (adopted, per inspecting engineer's recommendation)"));
rows2.push(fieldRow(5, "Total land value", "Rs. 15,42,500  (6.17 cents x Rs. 2,50,000/cent)"));
rows2.push(fieldRow(6, "Approved built up area and usage", "Residential"));
rows2.push(fieldRow(7, "Area as per plan (Sq.ft)", "Not furnished (no approved plan provided)"));
rows2.push(fieldRow(8, "Area as per site measurement (Sqft)", "602.77 sqft"));
rows2.push(fieldRow(9, "Area as per K-smart (Sqft)", "Not verified/checked"));
rows2.push(fieldRow(10, "Area considered for valuation (Sqft)", "602.77 sqft (site-measured)"));
rows2.push(fieldRow(11, "Building construction rate", "Rs. 1,600/sqft — ASSUMED (new-construction basis for Average quality finish; not specified in site notes). Please confirm/override this rate.", {flag:true}));
rows2.push(fieldRow(12, "Total construction cost of the building", "602.77 sqft x Rs. 1,600/sqft = Rs. 9,64,432 (new-construction basis)", {flag:true}));
rows2.push(fieldRow(13, "Depreciation applied", "~40% (assumed — 25-year-old structure, Average condition, on an assumed 60-year useful life). Please confirm/override.", {flag:true}));
rows2.push(fieldRow(14, "Building value based on current stage", "Rs. 5,78,659 (after ~40% age/condition depreciation on Rs. 9,64,432)", {flag:true}));
rows2.push(fieldRow(15, "Market value of property", "Rs. 15,42,500 (Land) + Rs. 5,78,659 (Building) = Rs. 21,21,159", {flag:true}));
rows2.push(fieldRow(16, "Govt. Guideline value", "Not verified — recommend querying Kerala Registration Dept. Fair Value portal (igr.kerala.gov.in)", {flag:true}));
rows2.push(fieldRow(17, "Realizable value of property (85%)", "Rs. 18,02,985", {flag:true}));
rows2.push(fieldRow(18, "Distress value of property (75%)", "Rs. 15,90,869", {flag:true}));
rows2.push(fieldRow(19, "Whether property is in Sensitive Area/Concentrated area", "Flagged by inspecting engineer for CRZ verification — not conclusively confirmed either way; recommend legal/CRZ clearance check before finalizing", {flag:true}));
rows2.push(fieldRow(20, "Longitude & Latitude of property", "9.506299, 76.316547  (±9m device GPS accuracy)"));
rows2.push(fieldRow(21, "Whether proper fencing of the Plot is done or not?", "Yes — compound wall on all sides"));
rows2.push(fieldRow(22, "Width of Road", "2.7m wide municipal road"));

const table2 = new Table({
  width: {size: TOTAL_W, type: WidthType.DXA},
  columnWidths: [COL_NUM, COL_LABEL, COL_VAL],
  rows: rows2
});

// ---------- REMARKS / DECLARATION ----------
const remarksPara = plainPara(
  "The Subject property consists of a residential building of BUA 602.77 sqft (as per site measurement) within a land extent of 6.17 cents, located at Poomkavu Church Road, Kommady, Alappuzha. The property is bounded on the North by Binu's property and a pathway, East and South by Municipal ways, and West by Alosius's property (all as per site inspection). The land extent of 6.17 cents is stated as consistent between the available documents and site measurement per the inspecting engineer's notes. The property is demarcated with a compound wall on all sides and is accessed through a 2.7m wide municipal road.",
  {size: 18}
);
const remarksPara2 = plainPara(
  "TITLE DEED: The title deed provided is a handwritten document dated 2001, referencing Re-Survey Nos. 387/1/4 and 384/6B. The cursive handwriting on the scanned copy could not be read with sufficient confidence to confirm the exact document number, transacting parties, or deed-boundary details — recommend verification against the original registered instrument before disbursement.",
  {size:18, italic:true, color:"B45309"}
);
const remarksPara3 = plainPara(
  "BUILDING RATE & DEPRECIATION NOT SPECIFIED: Site notes did not record a building construction rate, structure type, roofing, flooring, masonry, or residual life. A rate of Rs. 1,600/sqft (new-construction, Average quality) with ~40% depreciation (25-year-old structure, Average condition) has been assumed to arrive at a current-stage building value of Rs. 5,78,659 — please confirm or adjust these assumptions.",
  {size:18, italic:true, color:"B45309"}
);
const remarksPara4 = plainPara(
  "CRZ: The inspecting engineer flagged this property for a CRZ (Coastal Regulation Zone) check; this has not been independently confirmed either way and should be verified before finalizing the report, given the property's location within Alappuzha town.",
  {size:18, italic:true, color:"B45309"}
);
const remarksPara5 = plainPara(
  "Land rate of Rs. 2.5 lakhs/cent has been adopted per the inspecting engineer's recommendation, within the informed ongoing range of Rs. 2.5–3 lakhs/cent for the locality.",
  {size:18}
);

const declarationPara = plainPara(
  "Declaration from the Valuers — I hereby declare that:\n" +
  "a) The property was inspected on 06-08-2026 by Er. Sreenath S on behalf of the undersigned.\n" +
  "b) The information furnished in this valuation report dated 10-08-2026 is true and correct to the best of my knowledge and belief, and I have made an impartial and true valuation of the property.\n" +
  "c) I have no direct or indirect interest in the property valued.\n" +
  "d) Legal documents/aspects fall beyond the scope of valuation.",
  {size: 18}
);
// docx TextRun doesn't render \n as line breaks; split manually below instead.

function declarationParas() {
  return [
    plainPara("Declaration from the Valuers — I hereby declare that:", {size:18, bold:true}),
    plainPara("a) The property was inspected on 06-08-2026 by Er. Sreenath S on behalf of the undersigned.", {size:18}),
    plainPara("b) The information furnished in this valuation report dated 10-08-2026 is true and correct to the best of my knowledge and belief, and I have made an impartial and true valuation of the property.", {size:18}),
    plainPara("c) I have no direct or indirect interest in the property valued.", {size:18}),
    plainPara("d) Legal documents/aspects fall beyond the scope of valuation.", {size:18}),
  ];
}

const complianceFlag = flagNote(
  "COMPLIANCE NOTE: The bank's standard declaration states that the IBBI registered valuer personally inspected the property. Per the site inspection sign-off, the physical site visit was conducted by Er. Sreenath S, not Mr. Dhanish Nair personally. The declaration wording above has been adjusted accordingly — please review and confirm this is acceptable, or amend per your firm's certification practice."
);

// ---------- SITE IMAGES ----------
function imgCell(path, caption) {
  const data = fs.readFileSync(path);
  return new TableCell({
    width: {size: TOTAL_W/2, type: WidthType.DXA},
    borders: borderSet(),
    margins:{top:80,bottom:80,left:80,right:80},
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [ new ImageRun({ data, type:"jpg", transformation: { width: 320, height: 240 } }) ]
      }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({text: caption, bold:true, size:16})]})
    ]
  });
}

const photoDir = __dirname;
const siteImagesTable = new Table({
  width: {size: TOTAL_W, type: WidthType.DXA},
  columnWidths: [TOTAL_W/2, TOTAL_W/2],
  rows: [
    new TableRow({children: [
      imgCell(photoDir+"/img-003-000.jpg", "Front Elevation / Approach"),
      imgCell(photoDir+"/img-003-004.jpg", "Gate / Frontage View")
    ]}),
    new TableRow({children: [
      imgCell(photoDir+"/img-003-001.jpg", "Interior View — Room"),
      imgCell(photoDir+"/img-003-002.jpg", "Interior View — Kitchen")
    ]}),
    new TableRow({children: [
      imgCell(photoDir+"/img-003-003.jpg", "Road Access / Approach Lane"),
      imgCell(photoDir+"/img-003-005.jpg", "Street View / Access Lane")
    ]}),
  ]
});

// ---------- DOCUMENT ----------
const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 11907, height: 16840 }, margins:{top:600, bottom:600, left:700, right:700} } },
    children: [
      headerTable,
      plainPara(""),
      table1,
      plainPara(""),
      boundaryTable,
      plainPara(""),
      table2,
      plainPara(""),

      new Paragraph({children:[new TextRun({text:"Remarks", bold:true, size:22})]}),
      remarksPara, remarksPara2, remarksPara3, remarksPara4, remarksPara5,
      plainPara(""),

      new Paragraph({children:[new TextRun({text:"Locality & Marketability of the Property: Average", bold:true, size:20})]}),
      plainPara(""),

      ...declarationParas(),
      plainPara(""),
      complianceFlag,
      plainPara(""),

      plainPara("Place: Alappuzha                                                              Signature:", {size:20, bold:true}),
      plainPara("Date: 10-08-2026", {size:20, bold:true}),

      new Paragraph({ children:[new PageBreak()] }),
      new Paragraph({heading: HeadingLevel.HEADING_2, children:[new TextRun({text:"Site Photographs", bold:true})]}),
      plainPara(""),
      siteImagesTable,
      plainPara(""),
      flagNote("Note: One additional photograph submitted with this case (a small tiled-roof house on Vatikad Road, Palakkulam, Alappuzha 688006, GPS 9.514714/76.339258, timestamped 06/08/2026 02:21 PM) has GPS coordinates and a timestamp that do not match this property's location (Poomkavu Church Road, Kommady — GPS ~9.5063/76.3166, ~03:48–03:50 PM) or the site inspection form. It has been excluded from this report as it does not appear to belong to this case — please confirm which case it relates to."),

      new Paragraph({ children:[new PageBreak()] }),
      new Paragraph({heading: HeadingLevel.HEADING_2, children:[new TextRun({text:"Google Location Map", bold:true})]}),
      plainPara(""),
      plainPara("Latitude & Longitude: 9.506299, 76.316547", {size:20}),
      plainPara("Map link: https://maps.google.com/?q=9.506299,76.316547", {size:20}),
      flagNote("Satellite/map screenshot not available in the supplied files — please capture and insert for the final record copy."),
      plainPara(""),

      new Paragraph({heading: HeadingLevel.HEADING_2, children:[new TextRun({text:"Fair Value", bold:true})]}),
      plainPara(""),
      flagNote("Fair Value not verified for this property — recommend querying the Kerala Registration Department Fair Value portal (igr.kerala.gov.in). Village/Block/Survey details required for this lookup could not be confirmed from the documents provided (see Legal Address note above)."),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(__dirname + "/Valuation_Report_Anoop_AUCB.docx", buf);
  console.log("done");
});
