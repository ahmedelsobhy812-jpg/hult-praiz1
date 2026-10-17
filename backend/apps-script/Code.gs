/**
 * Registration backend (Google Apps Script web app).
 *
 * Deploy as: Execute as "Me", access "Anyone". The web app URL is public, so the
 * protections below live here: server-side validation, duplicate handling, rate
 * limiting, a honeypot field and a per-application token for the PDF.
 * No secrets go in the website's JavaScript. Keep the Sheet and Drive folder private.
 */

var CONFIG = {
  SHEET_ID: "1V4YMdcOAxRwaVdm1c2oJzqMHA_UXVbOewSIOLth6yNM",            // ID of the Google Sheet (the long string in its URL)
  SHEET_NAME: "Applications",
  PDF_FOLDER_ID: "",       // Drive folder for PDF copies. Empty = do not store copies.
  COMPANY_NAME: "Hult Prize \u2013 Borg El Arab Technological University",
  DUPLICATE_POLICY: "reject", // TODO: "reject" blocks a second application with the same email
  MAX_REQUESTS_PER_MINUTE: 30,
  MIN_SKILLS: 1
};

var HEADERS = [
  "Submitted At", "Reference", "Request ID", "Token Hash", "Full Name", "Date of Birth", "Email", "Gender",
  "Country", "Governorate", "City", "Faculty", "Department", "Academic Year", "Phone", "Academic ID",
  "National ID", "Committee", "Skills", "PDF File Link"
];

var FACULTIES = {
  "Faculty of Industry and Energy Technology": {
    fromYear: 1,
    departments: ["Information Technology", "Railway Technology", "Food Manufacturing Technology", "Tractors and Equipment Technology", "Spinning and Weaving Technology"]
  },
  "Faculty of Applied Health Sciences Technology": {
    fromYear: 2,
    departments: ["Medical Laboratory Technology", "Radiology Technology", "Physical Therapy Technology", "Nursing Technology"] // TODO: placeholders
  }
};

// Keep in sync with js/committees.js
var COMMITTEES = {
  "Events Team": ["Event Planning", "Logistics", "Hosting & Moderation", "Budget Tracking"],
  "Experts & Judges Support Team": ["Professional Outreach", "Scheduling", "Hospitality", "Evaluation Support"],
  "Media Team": ["Graphic Design", "Video Editing", "Photography", "Content Writing", "Social Media Management"],
  "Partnerships & Relations Team": ["Sponsorship Outreach", "Negotiation", "Proposal Writing", "Relationship Management"],
  "Startups Team": ["Business Modeling", "Pitch Coaching", "Market Research", "Mentoring"]
};

/* ---------------- entry points ---------------- */

function doPost(e) {
  try {
    if (!rateLimitOk()) return json({ ok: false, error: "Too many requests. Please wait a minute and try again." });
    var body = JSON.parse((e.postData && e.postData.contents) || "{}");
    if (body.action !== "submit") return json({ ok: false, error: "Unknown action." });
    if (body.website) return json({ ok: true, ref: "HP-0000000", pdf: { status: "failed" } }); // honeypot: pretend success

    var clean = validate(body);
    if (clean.error) return json({ ok: false, error: clean.error });
    return json(saveApplication(clean.value));
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: "Server error. Please try again." });
  }
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action !== "pdf" || !p.ref || !p.token) return json({ ok: false, error: "Bad request." });
    var row = findRow("Reference", p.ref);
    if (!row || row.values[3] !== hash(p.token)) return json({ ok: false, error: "Not found." });
    return json({ ok: true, pdf: makePdf(rowToApplication(row.values), p.ref) });
  } catch (err) {
    console.error("PDF retry failed: " + err);
    return json({ ok: false, error: "PDF failed: " + String(err && err.message || err).slice(0, 160) });
  }
}

/* ---------------- saving ---------------- */

function saveApplication(a) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet();

    // A retry of the same request returns the same reference instead of a second row
    var same = findRow("Request ID", a.requestId);
    if (same) return { ok: true, ref: same.values[1], token: "", duplicate: true, pdf: { status: "failed" } };

    if (CONFIG.DUPLICATE_POLICY === "reject" && findRow("Email", a.email)) {
      return { ok: false, error: "An application with this email already exists." };
    }

    var ref = nextReference(sheet);
    var token = Utilities.getUuid();
    sheet.appendRow([
      new Date(), ref, a.requestId, hash(token), safe(a.fullName), safe(a.dob), safe(a.email), safe(a.gender),
      safe(a.country), safe(a.governorate), safe(a.city), safe(a.faculty), safe(a.department), safe(a.year),
      safe(a.phone), safe(a.academicId), safe(a.nationalId), safe(a.committee), safe(a.skills.join(", ")), ""
    ]);
    SpreadsheetApp.flush();

    // The PDF is created only after the row is saved. A PDF failure never loses the registration.
    var pdf = { status: "failed" };
    try {
      pdf = makePdf(a, ref);
      var link = storePdf(pdf);
      if (link) sheet.getRange(sheet.getLastRow(), HEADERS.indexOf("PDF File Link") + 1).setValue(link);
    } catch (pdfErr) {
      console.error("PDF failed: " + pdfErr);
      pdf = { status: "failed", error: String(pdfErr && pdfErr.message || pdfErr).slice(0, 160) };
    }
    return { ok: true, ref: ref, token: token, pdf: pdf };
  } finally {
    lock.releaseLock();
  }
}

function nextReference(sheet) {
  var n = Math.max(sheet.getLastRow() - 1, 0) + 1; // header row excluded
  return "HP-" + Utilities.formatDate(new Date(), "UTC", "yyyy") + "-" + ("0000" + n).slice(-4);
}

/* ---------------- PDF ---------------- */

function makePdf(a, ref) {
  try { return buildPdf(a, ref, true); }
  catch (e) {
    console.error("PDF with logo failed, retrying without it: " + e);
    return buildPdf(a, ref, false);
  }
}

function buildPdf(a, ref, withLogo) {
  var logo = withLogo
    ? '<img src="data:image/png;base64,' + HP_LOGO_B64 + '" height="72" style="height:72px">'
    : '<b style="font-size:16px;color:#4e276d">Hult Prize</b>';
  var html =
    '<html><body style="font-family:Arial,sans-serif;color:#2b1442;margin:0">' +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
      '<td style="width:90px;padding:20px 0 14px 28px;vertical-align:middle">' + logo + '</td>' +
      '<td style="padding:20px 28px 14px 14px;vertical-align:middle">' +
        '<div style="font-size:12px;letter-spacing:2px;color:#f63a9c;font-weight:bold">APPLICATION</div>' +
        '<div style="font-size:21px;font-weight:bold;color:#4e276d;margin-top:4px">Application for Joining ' + esc(CONFIG.COMPANY_NAME) + '</div>' +
      '</td></tr></table>' +
    '<div style="height:5px;background:#f63a9c"></div>' +
    '<div style="padding:18px 28px">' +
    '<p style="margin:0 0 14px"><b>Reference:</b> ' + esc(ref) + ' &nbsp;&nbsp; <b>Submitted:</b> ' + Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd HH:mm") + ' UTC</p>' +
    section("Personal Information", [["Full Name", a.fullName], ["Date of Birth", a.dob], ["Email", a.email], ["Gender", a.gender],
      ["Location", [a.city, a.governorate, a.country].filter(String).join(", ")]]) +
    section("Academic Information", [["Faculty", a.faculty], ["Academic Year", a.year], ["Department", a.department || "-"],
      ["Phone", a.phone], ["Academic ID", a.academicId], ["National ID", a.nationalId]]) +
    section("Committee", [["Committee", a.committee], ["Skills", (a.skills || []).join(", ")]]) +
    '</div></body></html>';
  var filename = "Application-" + ref + ".pdf";
  var blob = htmlToPdfBlob(html, filename);
  return { status: "ready", filename: filename, base64: Utilities.base64Encode(blob.getBytes()) };
}

// Direct conversion first. If it is refused, convert through a temporary Drive file.
function htmlToPdfBlob(html, filename) {
  try {
    return Utilities.newBlob(html, "text/html", filename + ".html").getAs("application/pdf").setName(filename);
  } catch (e1) {
    console.error("Direct PDF conversion failed: " + e1);
    var tmp = DriveApp.createFile(filename + ".html", html, MimeType.HTML);
    try { return tmp.getAs(MimeType.PDF).setName(filename); }
    finally { tmp.setTrashed(true); }
  }
}

// Run this once from the editor to approve permissions and check that PDFs work.
function testPdf() {
  var pdf = buildPdf({ fullName: "Test User", dob: "2004-01-01", email: "test@example.com", gender: "Male", country: "Egypt",
    governorate: "Alexandria", city: "Agami", faculty: "Faculty of Industry and Energy Technology", year: "1", department: "Information Technology",
    phone: "+20 1012345678", academicId: "AB1234", nationalId: "30101010101010", committee: "Media Team", skills: ["Photography"] }, "HP-TEST", true);
  Logger.log("PDF ok, base64 length: " + pdf.base64.length);
}

function section(title, rows) {
  var h = '<h3 style="color:#4e276d;border-bottom:2px solid #f63a9c;padding-bottom:4px;margin:18px 0 8px">' + title + '</h3><table style="width:100%;border-collapse:collapse">';
  rows.forEach(function (r) {
    h += '<tr><td style="width:32%;padding:5px 0;color:#6e5f82;vertical-align:top">' + esc(r[0]) + '</td><td style="padding:5px 0;font-weight:bold;word-break:break-word">' + esc(r[1]) + '</td></tr>';
  });
  return h + "</table>";
}

function storePdf(pdf) {
  if (!CONFIG.PDF_FOLDER_ID || pdf.status !== "ready") return "";
  var blob = Utilities.newBlob(Utilities.base64Decode(pdf.base64), "application/pdf", pdf.filename);
  return DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID).createFile(blob).getUrl(); // access follows the folder's sharing
}

/* ---------------- validation (mirrors js/form-validation.js) ---------------- */

function validate(b) {
  var v = {
    requestId: str(b.requestId, 80), fullName: str(b.fullName, 80), dob: str(b.dob, 10), email: str(b.email, 120).toLowerCase(),
    gender: str(b.gender, 30), country: str(b.country, 60), governorate: str(b.governorate, 60), city: str(b.city, 60),
    faculty: str(b.faculty, 100), year: str(b.year, 2), department: str(b.department, 100), phone: str(b.phone, 24),
    academicId: str(b.academicId, 20), nationalId: str(b.nationalId, 14), committee: str(b.committee, 80),
    skills: (Array.isArray(b.skills) ? b.skills : []).map(function (s) { return str(s, 60); })
  };
  if (!v.requestId) return fail("Missing request ID.");
  if (v.fullName.split(/\s+/).length < 2 || v.fullName.length < 3) return fail("Enter your full name.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v.dob) || isNaN(new Date(v.dob + "T00:00:00Z"))) return fail("Invalid date of birth.");
  var age = (new Date() - new Date(v.dob + "T00:00:00Z")) / 31557600000;
  if (age < 16 || age > 45) return fail("Age must be between 16 and 45.");
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(v.email)) return fail("Invalid email.");
  if (["Female", "Male", "Prefer not to say"].indexOf(v.gender) < 0) return fail("Invalid gender.");
  if (!v.country || !v.governorate || !v.city) return fail("Location is incomplete.");
  var fac = FACULTIES[v.faculty];
  if (!fac) return fail("Invalid faculty.");
  var yr = parseInt(v.year, 10);
  if (!(yr >= 1 && yr <= 4)) return fail("Invalid academic year.");
  if (yr >= fac.fromYear) { if (fac.departments.indexOf(v.department) < 0) return fail("Invalid department."); }
  else v.department = "";
  if (!/^\+\d{1,3} \d{6,12}$/.test(v.phone)) return fail("Invalid phone number.");
  if (!/^[A-Za-z0-9-]{4,20}$/.test(v.academicId)) return fail("Invalid academic ID.");
  if (!/^\d{14}$/.test(v.nationalId)) return fail("Invalid national ID.");
  var allowed = COMMITTEES[v.committee];
  if (!allowed) return fail("Invalid committee.");
  var skills = [];
  v.skills.forEach(function (s) { if (allowed.indexOf(s) >= 0 && skills.indexOf(s) < 0) skills.push(s); });
  if (skills.length !== v.skills.length || skills.length < CONFIG.MIN_SKILLS) return fail("Invalid skills for this committee.");
  v.skills = skills; // names only, no descriptions
  return { value: v };
}
function fail(msg) { return { error: msg }; }
function str(x, max) { return (x == null ? "" : String(x)).trim().slice(0, max); }

/* ---------------- helpers ---------------- */

function getSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) { sheet.appendRow(HEADERS); sheet.setFrozenRows(1); }
  return sheet;
}
function findRow(header, value) {
  var sheet = getSheet(), last = sheet.getLastRow();
  if (last < 2) return null;
  var col = HEADERS.indexOf(header) + 1;
  var vals = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === String(value).toLowerCase()) {
      return { index: i + 2, values: sheet.getRange(i + 2, 1, 1, HEADERS.length).getValues()[0] };
    }
  }
  return null;
}
function rowToApplication(r) {
  return { fullName: r[4], dob: formatDob(r[5]), email: r[6], gender: r[7], country: r[8], governorate: r[9], city: r[10], faculty: r[11],
    department: r[12], year: r[13], phone: r[14], academicId: r[15], nationalId: r[16], committee: r[17],
    skills: String(r[18]).split(", ").filter(String) };
}
function formatDob(v) { return v instanceof Date ? Utilities.formatDate(v, "UTC", "yyyy-MM-dd") : String(v); }
function hash(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s)));
}
// A cell that starts with = + - @ would run as a formula, so prefix it with an apostrophe
function safe(v) { v = v == null ? "" : String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function rateLimitOk() {
  var cache = CacheService.getScriptCache(), key = "rl-" + Math.floor(Date.now() / 60000);
  var n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 90);
  return n <= CONFIG.MAX_REQUESTS_PER_MINUTE;
}

/* Logo for the PDF header (PNG, base64). */
var HP_LOGO_B64 = "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";
