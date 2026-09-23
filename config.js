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
// dropdown and written into the printed docx + email sign-off), and
// which branch they're based at (drives the CC list above — see
// emailCcByBranch). Full names are as listed in Staff Micronet 2026.xlsx.
//
// This is the single source of truth for the public Leave Application
// Form (index.html). Add a new staff member here as they start using the
// form.
// ---------------------------------------------------------------------
const LEAVE_FORM_STAFF_ROSTER = [
  { short: "Khairul", full: "Awangku Muhammad Khairul Amir Pengiran Darma Putra", branch: "Jerudong" },
  { short: "Amal", full: "Amal Rafidah bte Haji Hamzah", branch: "Jerudong" },
  { short: "Nurlizam", full: "Nurlizam bte Pungut/Ismail", branch: "Jerudong" },
  { short: "Lyana", full: "Lyana Anak Berawoh", branch: "Jerudong" },
  { short: "Veronica", full: "Veronica Anne binti Roddy", branch: "Jerudong" },
  { short: "Hariz", full: "Muhammad Hariz Mahyuddin bin Abdullah", branch: "Gadong" },
  { short: "Aqilah", full: "Hajah Nur' Aqilah binti Haji Ahmad", branch: "Gadong" },
  { short: "Norain", full: "Norain binti Haji Matusin", branch: "Gadong" },
  { short: "Maziyah", full: "Siti Fathin Maziyyah Amal Hayati binti Haji Metussin", branch: "Gadong" },
  { short: "Alisha", full: "Alisha bte Abdul Latip @ Alicecia Jata Anak Latip", branch: "Gadong" },
  { short: "Aslam", full: "Mohammad Afham Aslam bin Mohd Rajimi", branch: "Gadong" },
  { short: "Sheraden", full: "Sheraden Lubuguin Mayani", branch: "Gadong" },
  { short: "Ummi", full: "Dayangku Hajah Ummi Syahirah binti Pengiran Haji Setia Putra", branch: "Gadong" },
  { short: "Izzaty", full: "Nur Izzaty Faezattul Watiqah binti Haji Hairman", branch: "Gadong" },
  { short: "Azimah", full: "Azimah binti Mohammad Hishammudin", branch: "Gadong" },
  { short: "Nurkhtamal", full: "Nurkhtamal binti Alinafiah", branch: "Gadong" },
  { short: "Nurzahidah", full: "Nurzahidah binti Haji Sahari", branch: "Gadong" },
  { short: "Nur Amelea", full: "Nur Amelea Batrisyia binti Suhaili", branch: "Jerudong" },
  { short: "Mustadim", full: "Muhammad Mustadim bin Haji Mohd Lias", branch: "Gadong" },
  { short: "Sharon", full: "Sharon Chin Lee Fah", branch: "Gadong" },
  { short: "Crisanta", full: "Crisanta Joveres Dionglay", branch: "Gadong" },
  { short: "Kalau", full: "Muhammad Nur Aiman bin Abdullah @ Kalau Anak Misen", branch: "Gadong" }
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
