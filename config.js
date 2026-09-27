// ---------------------------------------------------------------------
// EDIT THIS FILE to match your real submission details.
// ---------------------------------------------------------------------
const LEAVE_FORM_CONFIG = {
  applicantDisplayName: "Awangku Muhammad Khairul Amir Pengiran Darma Putra",
  applicantTitle: "Lecturer",

  // Recipient for the submission email — same for everyone.
  emailTo: "maziyyah@micronetbrunei.com",

  // CC depends on which branch the applicant belongs to (see "branch" on
  // each entry in LEAVE_FORM_STAFF_ROSTER below): Jerudong staff CC both
  // Sharon and Aqilah, Gadong staff CC Sharon only. handleEmail() in
  // app.js picks the right one for whoever is selected in the Name
  // dropdown — don't reference a single flat emailCc anywhere else.
  emailCcByBranch: {
    Jerudong: "sharon@micronetbrunei.com, aqilah@micronetbrunei.com",
    Gadong: "sharon@micronetbrunei.com"
  },

  emailGreetingName: "Ms. Maziyyah",

  // ---------------------------------------------------------------------
  // Submission email sign-off. The "Prepare Submission Email" button signs
  // the email as whoever is picked in the Name dropdown, not always
  // Khairul:
  //   - A name listed here in signatureOverrides gets its full, personal
  //     block (title, direct extension, mobile, email) exactly as before.
  //   - Any other staff name gets their own full name + the general
  //     College address block below (genericSignatureCollegeBlock) — no
  //     personal extension/mobile/email, since those aren't on file per
  //     staff.
  // Keyed by SHORT name (the "short" field in LEAVE_FORM_STAFF_ROSTER
  // below), not the full name shown in the dropdown.
  // ---------------------------------------------------------------------
  signatureOverrides: {
    "Khairul": [
      "Best Regards,",
      "",
      "Awangku Muhammad Khairul Amir Pengiran Darma Putra",
      "Lecturer",
      "",
      "MICRONET INTERNATIONAL COLLEGE",
      "Gadong Campus [Head Office]:",
      "No. 11 & 12, Kompleks Hj Tahir 2, Sungai Gadong Menglait, BSB BE4119 Brunei Darussalam    P O Box 933 Gadong BE3978",
      "O: +673-2451133 ext 13        M: +673-7250492    F:  +673-2450888",
      "E:  khairul@micronet.com.bn       www.micronet.com.bn",
      "",
      "Jerudong Branch:",
      "Unit C8, C9, C10, Complex Jerudong, Simpang 508, Jalan Jerudong, BSB  BG3122 Brunei Darussalam",
      "O:  +673-2611133",
      "",
      "PEARSON BTEC APPROVED CENTRE  |  UTB - MIC COMPUTING DEGREE PROGRAMMES  |",
      "UTB - SP BRIDGING PROGRAMME IN COMPUTING (BRICOMP)  |  IBTE APPROVED CENTRE"
    ].join("\n")
  },

  genericSignatureCollegeBlock: [
    "MICRONET INTERNATIONAL COLLEGE",
    "Gadong Campus [Head Office]:",
    "No. 11 & 12, Kompleks Hj Tahir 2, Sungai Gadong Menglait, BSB BE4119 Brunei Darussalam   P O Box 933 Gadong BE3978",
    "O: +673-2451133    F: +673-2450888",
    "www.micronet.com.bn"
  ].join("\n")
};

// ---------------------------------------------------------------------
// Staff roster: short name (used internally — StaffAuth login, the
// review/import backend, signatureOverrides keys, and the exported
// filename), full legal name (shown in the Leave Application Form's Name
// dropdown and written into the printed docx + email sign-off), which
// branch they're based at (drives the CC list above — see
// emailCcByBranch), and their role, which drives the Annual Leave
// accrual schedule on the Staff Leave Breakdown page (see
// ANNUAL_ACCRUAL_SHAPES in leave-breakdown.js):
//   - "lecturer"    — +1 day/month Jan–Sep, +2 day/month Oct–Dec (15/yr)
//   - "admin"       — +1 day/month Jan–Oct, +2 day/month Nov–Dec (14/yr)
//   - "gm"          — +1 day/month Jan–May, +2 day/month Jun–Dec (19/yr),
//                     currently just Sharon (General Manager)
//   - "part-timer"  — no annual leave accrual; the form is used purely
//                     as a formality for this person (currently
//                     Veronica)
// Full names are as listed in Staff Micronet 2026.xlsx.
//
// This is the single source of truth for the public Leave Application
// Form (index.html). Add a new staff member here as they start using the
// form.
// ---------------------------------------------------------------------
const LEAVE_FORM_STAFF_ROSTER = [
  { short: "Khairul", full: "Awangku Muhammad Khairul Amir Pengiran Darma Putra", branch: "Jerudong", role: "lecturer" },
  { short: "Amal", full: "Amal Rafidah bte Haji Hamzah", branch: "Jerudong", role: "lecturer" },
  { short: "Nurlizam", full: "Nurlizam bte Pungut/Ismail", branch: "Jerudong", role: "lecturer" },
  { short: "Lyana", full: "Lyana Anak Berawoh", branch: "Jerudong", role: "lecturer" },
  { short: "Veronica", full: "Veronica Anne binti Roddy", branch: "Jerudong", role: "part-timer" },
  { short: "Hariz", full: "Muhammad Hariz Mahyuddin bin Abdullah", branch: "Gadong", role: "lecturer" },
  { short: "Aqilah", full: "Hajah Nur' Aqilah binti Haji Ahmad", branch: "Gadong", role: "admin" },
  { short: "Norain", full: "Norain binti Haji Matusin", branch: "Gadong", role: "admin" },
  { short: "Maziyah", full: "Siti Fathin Maziyyah Amal Hayati binti Haji Metussin", branch: "Gadong", role: "admin" },
  { short: "Alisha", full: "Alisha bte Abdul Latip @ Alicecia Jata Anak Latip", branch: "Gadong", role: "lecturer" },
  { short: "Aslam", full: "Mohammad Afham Aslam bin Mohd Rajimi", branch: "Gadong", role: "admin" },
  { short: "Sheraden", full: "Sheraden Lubuguin Mayani", branch: "Gadong", role: "lecturer" },
  { short: "Ummi", full: "Dayangku Hajah Ummi Syahirah binti Pengiran Haji Setia Putra", branch: "Gadong", role: "admin" },
  { short: "Izzaty", full: "Nur Izzaty Faezattul Watiqah binti Haji Hairman", branch: "Gadong", role: "lecturer" },
  { short: "Azimah", full: "Azimah binti Mohammad Hishammudin", branch: "Gadong", role: "lecturer" },
  { short: "Nurkhtamal", full: "Nurkhtamal binti Alinafiah", branch: "Gadong", role: "admin" },
  { short: "Nurzahidah", full: "Nurzahidah binti Haji Sahari", branch: "Gadong", role: "admin" },
  { short: "Nur Amelea", full: "Nur Amelea Batrisyia binti Suhaili", branch: "Jerudong", role: "admin" },
  { short: "Mustadim", full: "Muhammad Mustadim bin Haji Mohd Lias", branch: "Gadong", role: "lecturer" },
  { short: "Sharon", full: "Sharon Chin Lee Fah", branch: "Gadong", role: "gm" },
  { short: "Crisanta", full: "Crisanta Joveres Dionglay", branch: "Gadong", role: "lecturer" },
  { short: "Kalau", full: "Muhammad Nur Aiman bin Abdullah @ Kalau Anak Misen", branch: "Gadong", role: "admin" }
];

// ---------------------------------------------------------------------
// Short names only. Used by the Staff Leave Breakdown page
// (leave-breakdown.js) to populate the sign-in / "Staff Member" pickers
// there, and as a fallback in app.js if a name ever needs matching
// without going through the roster above.
// ---------------------------------------------------------------------
const LEAVE_FORM_STAFF_NAMES = LEAVE_FORM_STAFF_ROSTER.map(function (r) { return r.short; });

// ---------------------------------------------------------------------
// Staff Leave Breakdown (leave-breakdown.html) data source.
//
// This page reads/writes leave records from a Google Sheet through a
// Google Apps Script Web App acting as a tiny API. Until you deploy that
// script and paste its URL below, the page will show setup instructions
// instead of data. See SETUP-STAFF-LEAVE-BREAKDOWN.md for the full,
// one-time setup steps (create the sheet, paste the script, deploy it).
// ---------------------------------------------------------------------
const LEAVE_LEDGER_API_URL = "https://script.google.com/macros/s/AKfycbz8jf69rhlOU-RzSpCvX6G90uYKeM0yHlga8T25eerxUKSq7ZyauD4rfqfxMNT2x2c56Q/exec";
