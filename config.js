// ---------------------------------------------------------------------
// EDIT THIS FILE to match your real submission details.
// ---------------------------------------------------------------------
const LEAVE_FORM_CONFIG = {
  applicantDisplayName: "Awangku Muhammad Khairul Amir Pengiran Darma Putra",
  applicantTitle: "Lecturer",

  // Recipients for the submission email (comma-separate multiple addresses)
  emailTo: "maziyyah@micronetbrunei.com",
  emailCc: "sharon@micronetbrunei.com, aqilah@micronetbrunei.com",

  emailGreetingName: "Ms. Maziyyah",

  // Signature block appended to the email body, matching your usual sign-off
  signatureBlock: [
    "Best Regards,",
    "",
    "Awangku Muhammad Khairul Amir Pengiran Darma Putra",
    "Lecturer",
    "",
    "MICRONET INTERNATIONAL COLLEGE",
    "Gadong Campus [Head Office]:",
    "No. 11 & 12, Kompleks Hj Tahir 2, Sungai Gadong Menglait, BSB BE4119 Brunei Darussalam   P O Box 933 Gadong BE3978",
    "O: +673-2451133 ext 13    M: +673-7250492    F: +673-2450888",
    "E: khairul@micronet.com.bn    www.micronet.com.bn"
  ].join("\n")
};

// ---------------------------------------------------------------------
// Known staff first names, used to pick the short name in the exported
// filename: "Leave Form (start to end) - <ShortName>.docx". Whatever the
// applicant types into the Name field is matched against this list
// (case-insensitive, matches anywhere in the typed name); the first match
// is used. Add a new staff member's first name here as they start using
// the form. If nobody matches, the last word of the typed name is used
// instead.
// ---------------------------------------------------------------------
const LEAVE_FORM_STAFF_NAMES = [
  "Khairul",
  "Amal",
  "Nurlizam",
  "Lyana",
  "Veronica",
  "Hariz"
];
