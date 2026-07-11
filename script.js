const skillCatalog = [
  { label: "SQL", terms: ["sql", "postgres", "mysql", "database querying"] },
  { label: "Python", terms: ["python", "pandas", "numpy", "scikit"] },
  { label: "Power BI", terms: ["power bi", "powerbi", "dax"] },
  { label: "Tableau", terms: ["tableau"] },
  { label: "Excel", terms: ["excel", "spreadsheet", "pivot table"] },
  { label: "Dashboard design", terms: ["dashboard", "dashboards", "data visualization", "visualisation"] },
  { label: "Stakeholder reporting", terms: ["stakeholder reporting", "executive reporting", "presentation", "insights"] },
  { label: "Business analytics", terms: ["business analytics", "commercial analytics", "kpi", "metrics"] },
  { label: "HR operations", terms: ["hr operations", "hiring funnel", "workforce", "recruitment analytics"] },
  { label: "Machine learning", terms: ["machine learning", "ml", "modeling", "classification"] },
  { label: "JavaScript", terms: ["javascript", "typescript", "node", "react"] },
  { label: "Cloud", terms: ["aws", "azure", "gcp", "cloud"] },
  { label: "ATS", terms: ["ats", "greenhouse", "lever", "workday"] },
  { label: "Recruiting", terms: ["recruiting", "talent acquisition", "sourcing", "screening"] },
  { label: "Customer success", terms: ["customer success", "account management", "onboarding"] }
];

const sampleCandidates = [
  {
    name: "Amina Okafor",
    role: "Senior People Analytics Specialist",
    cv:
      "Senior analyst with 6 years of experience across HR operations, hiring funnel analytics, workforce planning, and executive reporting. Advanced SQL, Python, Power BI, DAX, dashboard design, Excel modeling, stakeholder reporting, KPI development, and presentation of insights to talent acquisition leaders. Built recruitment dashboards that reduced manual reporting time by 45%."
  },
  {
    name: "Daniel Mensah",
    role: "Data Analyst",
    cv:
      "Data analyst with 4 years of experience in business analytics, SQL, Tableau, Excel, and dashboard development. Strong stakeholder communication and weekly operations reporting. Built sales and finance scorecards, used Python for data cleaning, and collaborated with cross-functional teams. Limited direct HR operations exposure."
  },
  {
    name: "Maya Patel",
    role: "Talent Acquisition Lead",
    cv:
      "Recruiting leader with 8 years of experience in ATS management, candidate screening, sourcing, interview design, HR operations, and hiring manager coordination. Strong process design and stakeholder reporting. Uses Excel and Greenhouse analytics, but has limited Python and SQL hands-on experience."
  },
  {
    name: "Sofia Chen",
    role: "Machine Learning Engineer",
    cv:
      "Machine learning engineer with 5 years of Python, SQL, model evaluation, classification, data pipelines, and cloud deployment on AWS. Built ranking models and analytics services for product teams. Comfortable with dashboard requirements, but Power BI and HR operations are not recent focus areas."
  },
  {
    name: "Chinedu Adeyemi",
    role: "Customer Success Manager",
    cv:
      "Customer success manager with 7 years of onboarding, account management, stakeholder presentations, and customer health reporting. Experienced in Excel and CRM dashboards. Has worked with hiring teams during software rollout, but no advanced SQL, Python, or Power BI delivery experience."
  }
];

let candidates = [];

const refs = {
  form: document.querySelector("#screeningForm"),
  roleTitle: document.querySelector("#roleTitle"),
  mustHave: document.querySelector("#mustHave"),
  jobRequirements: document.querySelector("#jobRequirements"),
  candidateResults: document.querySelector("#candidateResults"),
  candidateCount: document.querySelector("#candidateCount"),
  candidateName: document.querySelector("#candidateName"),
  candidateRole: document.querySelector("#candidateRole"),
  candidateCv: document.querySelector("#candidateCv"),
  cvFiles: document.querySelector("#cvFiles"),
  contactForm: document.querySelector("#contactForm"),
  contactName: document.querySelector("#contactName"),
  contactEmail: document.querySelector("#contactEmail"),
  contactCompany: document.querySelector("#contactCompany"),
  contactVolume: document.querySelector("#contactVolume"),
  contactMessage: document.querySelector("#contactMessage"),
  contactStatus: document.querySelector("#contactStatus")
};

const stopWords = new Set([
  "with",
  "from",
  "that",
  "this",
  "into",
  "using",
  "need",
  "years",
  "year",
  "role",
  "will",
  "work",
  "strong",
  "should",
  "clear",
  "teams",
  "data",
  "candidate",
  "experience",
  "business",
  "requirements",
  "person",
  "cross",
  "functional"
]);

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9+#.]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function titleCase(value) {
  return value
    .trim()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function containsPhrase(text, phrase) {
  const haystack = ` ${normalize(text)} `;
  const needle = ` ${normalize(phrase)} `;
  return haystack.includes(needle);
}

function extractSkills(text) {
  const found = skillCatalog
    .filter((skill) => skill.terms.some((term) => containsPhrase(text, term)))
    .map((skill) => skill.label);
  return unique(found);
}

function extractYears(text) {
  const matches = [...String(text || "").matchAll(/(\d{1,2})\+?\s*(?:years|yrs|year)/gi)];
  if (!matches.length) return 0;
  return Math.max(...matches.map((match) => Number(match[1])));
}

function getRequirements() {
  const manualSkills = refs.mustHave.value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      const known = skillCatalog.find((skill) => containsPhrase(item, skill.label));
      return known ? known.label : titleCase(item);
    });

  const combinedText = `${refs.roleTitle.value} ${refs.mustHave.value} ${refs.jobRequirements.value}`;
  return {
    title: refs.roleTitle.value.trim(),
    text: combinedText,
    requiredSkills: unique([...extractSkills(combinedText), ...manualSkills]).slice(0, 12),
    requiredYears: extractYears(combinedText),
    keywords: extractKeywords(combinedText)
  };
}

function extractKeywords(text) {
  return unique(
    normalize(text)
      .split(" ")
      .filter((word) => word.length > 3 && !stopWords.has(word))
  ).slice(0, 30);
}

function skillMatches(cvText, skill) {
  const catalogMatch = skillCatalog.find((item) => item.label === skill);
  if (catalogMatch) {
    return catalogMatch.terms.some((term) => containsPhrase(cvText, term));
  }
  return containsPhrase(cvText, skill);
}

function scoreCandidate(candidate, requirements) {
  const cvText = `${candidate.name} ${candidate.role} ${candidate.cv}`;
  const matchedSkills = requirements.requiredSkills.filter((skill) => skillMatches(cvText, skill));
  const missingSkills = requirements.requiredSkills.filter((skill) => !matchedSkills.includes(skill));
  const extractedSkills = extractSkills(cvText);
  const keywordHits = requirements.keywords.filter((keyword) => containsPhrase(cvText, keyword));
  const candidateYears = extractYears(cvText);

  const skillScore = requirements.requiredSkills.length
    ? (matchedSkills.length / requirements.requiredSkills.length) * 56
    : 30;
  const keywordScore = requirements.keywords.length
    ? Math.min(20, (keywordHits.length / Math.min(requirements.keywords.length, 16)) * 20)
    : 10;
  const yearsScore = requirements.requiredYears
    ? Math.min(12, (candidateYears / requirements.requiredYears) * 12)
    : Math.min(12, candidateYears * 1.6);
  const titleKeywords = extractKeywords(requirements.title);
  const titleHits = titleKeywords.filter((word) => containsPhrase(`${candidate.role} ${candidate.cv}`, word));
  const titleScore = titleKeywords.length
    ? Math.min(8, (titleHits.length / Math.min(titleKeywords.length, 3)) * 8)
    : 0;
  const evidenceBonus = Math.min(4, extractedSkills.length * 0.5);
  const score = Math.round(Math.min(98, skillScore + keywordScore + yearsScore + titleScore + evidenceBonus));

  return {
    ...candidate,
    score,
    matchedSkills,
    missingSkills,
    extractedSkills,
    keywordHits,
    candidateYears,
    recommendation: score >= 76 ? "Shortlist" : score >= 56 ? "Review" : "Hold"
  };
}

function renderResults() {
  const requirements = getRequirements();
  const ranked = candidates
    .map((candidate) => scoreCandidate(candidate, requirements))
    .sort((a, b) => b.score - a.score);

  refs.candidateCount.textContent = `${ranked.length} CV${ranked.length === 1 ? "" : "s"}`;

  if (!ranked.length) {
    refs.candidateResults.innerHTML =
      '<div class="empty-state"><p>Add CVs or load the sample candidates to generate a ranked shortlist.</p></div>';
    return;
  }

  refs.candidateResults.innerHTML = ranked
    .map((candidate, index) => {
      const strengths = candidate.matchedSkills.slice(0, 6);
      const gaps = candidate.missingSkills.slice(0, 4);
      const summary = buildSummary(candidate);
      return `
        <article class="candidate-card">
          <div class="rank-badge">#${index + 1}</div>
          <div class="candidate-main">
            <div class="candidate-title">
              <h3>${escapeHtml(candidate.name)}</h3>
              <span class="role-pill">${escapeHtml(candidate.recommendation)}</span>
            </div>
            <p class="candidate-summary"><strong>${escapeHtml(candidate.role)}</strong> - ${summary}</p>
            <div class="skill-list" aria-label="Matched skills">
              ${strengths.map((skill) => `<span>${escapeHtml(skill)}</span>`).join("") || "<span>General match</span>"}
            </div>
            ${
              gaps.length
                ? `<div class="gap-list" aria-label="Missing skills">${gaps
                    .map((gap) => `<span>Missing: ${escapeHtml(gap)}</span>`)
                    .join("")}</div>`
                : ""
            }
          </div>
          <div class="score-block">
            <div class="score">${candidate.score}</div>
            <div class="score-label">Match</div>
          </div>
        </article>
      `;
    })
    .join("");
}

function buildSummary(candidate) {
  const years = candidate.candidateYears ? `${candidate.candidateYears}+ years found` : "experience depth not stated";
  const skills = candidate.matchedSkills.length
    ? `${candidate.matchedSkills.length} required skill${candidate.matchedSkills.length === 1 ? "" : "s"} matched`
    : "few direct requirement matches";
  const keywords = candidate.keywordHits.length
    ? `${candidate.keywordHits.length} role signals detected`
    : "limited role keyword evidence";
  return `${skills}, ${years}, ${keywords}.`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function loadSampleCandidates() {
  candidates = sampleCandidates.map((candidate) => ({ ...candidate }));
  renderResults();
}

function addCandidateFromForm() {
  const name = refs.candidateName.value.trim();
  const role = refs.candidateRole.value.trim();
  const cv = refs.candidateCv.value.trim();

  if (!name || !role || !cv) {
    refs.candidateCv.focus();
    return;
  }

  candidates.push({ name, role, cv });
  refs.candidateName.value = "";
  refs.candidateRole.value = "";
  refs.candidateCv.value = "";
  renderResults();
}

async function importFiles(event) {
  const files = [...event.target.files].filter((file) => /\.(txt|md)$/i.test(file.name));
  const imported = await Promise.all(
    files.map(async (file) => ({
      name: file.name.replace(/\.(txt|md)$/i, "").replace(/[-_]+/g, " "),
      role: "Imported CV",
      cv: await file.text()
    }))
  );

  candidates.push(...imported);
  refs.cvFiles.value = "";
  renderResults();
}

function handleContactSubmit(event) {
  event.preventDefault();

  if (!refs.contactForm.checkValidity()) {
    refs.contactForm.reportValidity();
    return;
  }

  const enquiry = {
    name: refs.contactName.value.trim(),
    email: refs.contactEmail.value.trim(),
    company: refs.contactCompany.value.trim(),
    volume: refs.contactVolume.value,
    message: refs.contactMessage.value.trim(),
    submittedAt: new Date().toISOString()
  };

  try {
    const savedEnquiries = JSON.parse(localStorage.getItem("getjobzEnquiries") || "[]");
    savedEnquiries.push(enquiry);
    localStorage.setItem("getjobzEnquiries", JSON.stringify(savedEnquiries));
  } catch (error) {
    console.warn("Could not save Getjobz enquiry locally.", error);
  }

  const subject = encodeURIComponent(`Getjobz enquiry from ${enquiry.company}`);
  const body = encodeURIComponent(
    [
      `Name: ${enquiry.name}`,
      `Email: ${enquiry.email}`,
      `Company: ${enquiry.company}`,
      `Hiring volume: ${enquiry.volume}`,
      "",
      "Hiring need:",
      enquiry.message
    ].join("\n")
  );

  refs.contactStatus.textContent = "Opening an email draft now.";
  window.location.href = `mailto:hello@getjobz.ai?subject=${subject}&body=${body}`;
  refs.contactForm.reset();
}

document.querySelector("[data-load-sample]").addEventListener("click", loadSampleCandidates);
document.querySelector("[data-clear-candidates]").addEventListener("click", () => {
  candidates = [];
  renderResults();
});
document.querySelector("[data-add-cv]").addEventListener("click", addCandidateFromForm);
refs.cvFiles.addEventListener("change", importFiles);
refs.contactForm.addEventListener("submit", handleContactSubmit);
refs.form.addEventListener("submit", (event) => {
  event.preventDefault();
  renderResults();
});

loadSampleCandidates();
