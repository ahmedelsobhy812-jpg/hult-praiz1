/* Central settings. Items marked TODO are open decisions from phase 0 of the roadmap. */
window.HP_CONFIG = {
  COMPANY_NAME: "Hult Prize \u2013 Borg El Arab Technological University", // final name, used in the PDF
  BRAND_NAME: "Hult Prize",
  CAMPUS_NAME: "Borg El Arab Technological University",
  LOGO_SRC: "assets/images/logo.png",
  INTRO_TEXT:
    "The official Hult Prize campus community at Borg El Arab Technological University, empowering students to develop innovative, sustainable business ideas that address global challenges. We connect aspiring entrepreneurs with mentors, experts, and opportunities to turn their ideas into impactful startups through collaboration, creativity, and social entrepreneurship.",

  // Paste the Apps Script web app URL here after deployment (see README).
  // While empty, the site runs in demo mode: nothing is sent anywhere.
  ENDPOINT_URL: "https://script.google.com/macros/s/AKfycbystmUuKj6WUWd6mSVMkd7PqWeBfBDim-qzZDT1aX3uCMYq7MKHsdhy9QmeyyzUaNon/exec",

  MIN_SKILLS: 1, // TODO: confirm the skills rule (at least one, or all required)
  REQUEST_TIMEOUT_MS: 20000,
  SUBMIT_RETRIES: 2,
  DRAFT_KEY: "hp-registration-draft-v1",

  // TODO: confirm the official English names. Health department names are placeholders.
  FACULTIES: [
    {
      name: "Faculty of Industry and Energy Technology",
      deptFromYear: 1,
      departments: [
        "Information Technology",
        "Railway Technology",
        "Food Manufacturing Technology",
        "Tractors and Equipment Technology",
        "Spinning and Weaving Technology",
      ],
    },
    {
      name: "Faculty of Applied Health Sciences Technology",
      deptFromYear: 2, // first-year students do not choose a department
      departments: [
        "Medical Laboratory Technology",
        "Radiology Technology",
        "Physical Therapy Technology",
        "Nursing Technology",
      ],
    },
  ],
  YEARS: ["1", "2", "3", "4"],
  GENDERS: ["Female", "Male", "Prefer not to say"],
};
