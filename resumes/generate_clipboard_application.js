/**
 * Generate Clipboard-tailored resume and cover letter (DOCX).
 * Run from repo root: node resumes/generate_clipboard_application.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  TabStopType,
  TabStopPosition,
} = require("docx");

const RESUMES_DIR = __dirname;
const COVER_LETTERS_DIR = path.join(RESUMES_DIR, "..", "cover_letters");

function center(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [
      new TextRun({
        text,
        bold: opts.bold ?? false,
        size: opts.size ?? 22,
        font: "Calibri",
      }),
    ],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
    children: [
      new TextRun({
        text,
        size: opts.size ?? 22,
        font: "Calibri",
        italics: opts.italics ?? false,
        bold: opts.bold ?? false,
      }),
    ],
  });
}

function sectionHeading(title) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    border: {
      bottom: { color: "333333", size: 6, style: BorderStyle.SINGLE },
    },
    children: [
      new TextRun({
        text: title.toUpperCase(),
        bold: true,
        size: 22,
        font: "Calibri",
      }),
    ],
  });
}

function roleHeader(company, title, meta) {
  return [
    new Paragraph({
      spacing: { before: 160, after: 40 },
      children: [
        new TextRun({ text: company, bold: true, size: 22, font: "Calibri" }),
        new TextRun({ text: ` — ${title}`, size: 22, font: "Calibri" }),
      ],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: meta,
          italics: true,
          size: 20,
          font: "Calibri",
        }),
      ],
    }),
  ];
}

function bullet(text) {
  return new Paragraph({
    spacing: { after: 60 },
    bullet: { level: 0 },
    children: [new TextRun({ text, size: 22, font: "Calibri" })],
  });
}

function buildResume() {
  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 720, right: 720 },
          },
        },
        children: [
          center("Nick Myers", { bold: true, size: 40 }),
          center(
            "Software Engineer (Frontend) · React, TypeScript, Node.js · Portland, OR · Remote",
            { size: 22 }
          ),
          center(
            "https://www.linkedin.com/in/nmyersdev/",
            { size: 20 }
          ),
          body(
            "Frontend-leaning full-stack engineer who ships user-facing web products in React and TypeScript, " +
              "partners with product and stakeholders on requirements, and keeps delivery testable from design through production support. " +
              "Four years building marketplace-style workflows, configurators, and high-traffic forms for education and public-sector users at Resource Data; " +
              "comfortable owning features end-to-end in Node.js/TypeScript and collaborating across distributed teams. " +
              "Seeking a remote role where customer impact, code quality, and continuous learning matter—aligned with Clipboard’s healthcare marketplace mission."
          ),
          sectionHeading("Technical skills"),
          body(
            "Frontend: TypeScript, JavaScript, React, Redux, Redux-Saga, HTML/CSS, Storybook, Bootstrap, SASS/LESS, responsive UI, accessibility-minded forms."
          ),
          body(
            "Backend & APIs: Node.js, .NET / .NET Core (MVC), GraphQL, REST, JSON; Postman-driven API testing and regression confidence."
          ),
          body(
            "Delivery: Git/GitHub, Agile/Scrum, trunk-style iteration, Heroku deployment experience; eager to deepen NestJS and AWS (ECS/Terraform) patterns."
          ),
          sectionHeading("Experience"),
          ...roleHeader(
            "Resource Data, Inc.",
            "Senior Programmer/Analyst & Programmer/Analyst",
            "May 2021 – January 2025 · Remote-friendly consulting · Custom web applications"
          ),
          bullet(
            "Built and maintained React/TypeScript experiences for manufacturing sales flows (IdeaRoom / American Steel)—configurable ordering, custom PDFs, and Storybook-documented components that reduced manual quoting friction."
          ),
          bullet(
            "Delivered public-sector web features (Washington DNR burn permitting; WSAC Career Launch)—complex forms, data tables, and map-driven workflows using React, Redux, and .NET APIs serving real end users under regulatory requirements."
          ),
          bullet(
            "Strengthened API reliability for e-commerce integrations (Bel/Cinch) with GraphQL/.NET services and structured Postman suites—surfacing defects early and improving release confidence."
          ),
          bullet(
            "Owned data migration and integration work for Epic Charter School (PowerSchool transitions, Edgenuity)—Node.js/TypeScript automations plus .NET/SQL Server performance tuning for concurrent, user-impacting workloads."
          ),
          bullet(
            "Partnered with product owners and client teams on requirements, demos, and production fixes; mentored developers on debugging practices and maintainable React patterns."
          ),
          ...roleHeader("Freelance Developer — NimblePath", "Full-stack product delivery", "2021 – 2023 · SaaS"),
          bullet(
            "Translated wireframes into a React/Redux UI and shipped a Java Spring API on Heroku—end-to-end ownership from problem definition through deployment."
          ),
          ...roleHeader("Team Lead — Lambda School", "Technical mentorship", "2020 – 2021"),
          bullet(
            "Coached ~10 students via weekly 1:1s, standups, and code review—building feedback loops similar to high-velocity, remote engineering teams."
          ),
          ...roleHeader("Earlier experience", "Customer-facing technical roles", "2016 – 2020"),
          bullet(
            "Phone support and field troubleshooting—strengthened communication, empathy, and calm problem-solving under time pressure (relevant to customer-centric product work)."
          ),
          sectionHeading("Education & certifications"),
          bullet("Full-Stack Web Development and Technical Interviewing — Lambda School"),
          bullet(
            "Boomi Associate MDH; Boomi Associate Developer; Boomi Professional Developer (integration background)"
          ),
        ],
      },
    ],
  });
}

function buildCoverLetter() {
  const today = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 720, right: 720 },
          },
        },
        children: [
          center("Nick Myers", { bold: true, size: 32 }),
          center("Portland, OR · Remote", { size: 22 }),
          center(
            "https://www.linkedin.com/in/nmyersdev/",
            { size: 20 }
          ),
          body(today, { after: 200 }),
          body("Hiring Team", { bold: true }),
          body("Clipboard — Software Engineer, Frontend", { after: 200 }),
          body(
            "Clipboard’s mission—connecting healthcare professionals with workplaces that need them—is the kind of product work I want next: software that directly improves people’s livelihoods and the care their communities receive. " +
              "I’m applying for the Frontend Software Engineer role because your stack (TypeScript, React, Node.js) and culture (remote-first, customer-centric, strong testing discipline) match how I’ve been shipping software for the past four years, and I’m motivated to grow with a profitable, high-impact YC team."
          ),
          body(
            "At Resource Data I spent most of my time on user-facing web applications in React and TypeScript—translating real requirements into maintainable UI, partnering with stakeholders through delivery, and supporting production when users depended on the system. " +
              "On IdeaRoom/American Steel I built configurable ordering flows and Storybook-backed components for a custom manufacturing sales experience. " +
              "For Washington DNR and WSAC Career Launch I improved complex forms, data views, and map-driven workflows where clarity and reliability mattered to non-technical users. " +
              "Across Bel/Cinch and Epic Charter School I paired frontend work with API testing and Node.js/TypeScript automation—habits that align with owning the full lifecycle, catching issues early, and keeping releases dependable."
          ),
          body(
            "What draws me to Clipboard specifically is the combination of pace and craft: trunk-based delivery, PR-driven interviews, and a “testing trophy” mindset are how strong teams scale without sacrificing user trust. " +
              "I thrive in async, written communication and have worked effectively with distributed clients and engineers; I’m also honest about growth areas—I’m eager to deepen NestJS and AWS deployment patterns (ECS/Terraform) on a team that values learning. " +
              "Customer focus is not new to me: earlier support and field roles taught me to listen first, and consulting taught me to tie every feature back to an outcome someone can feel."
          ),
          body(
            "I would welcome the chance to contribute to features nurses and facilities use every day, and to learn from your engineering team through the PR review process. " +
              "Thank you for your time and consideration—I’m happy to share more detail on any project above or walk through code samples."
          ),
          body("Sincerely,", { before: 160 }),
          body("Nick Myers", { bold: true }),
        ],
      },
    ],
  });
}

async function main() {
  fs.mkdirSync(COVER_LETTERS_DIR, { recursive: true });

  const resumePath = path.join(RESUMES_DIR, "Nick_Myers_Resume_CB.docx");
  const coverPath = path.join(
    COVER_LETTERS_DIR,
    "Nick Myers Cover Letter CB.docx"
  );

  const resumeBuf = await Packer.toBuffer(buildResume());
  const coverBuf = await Packer.toBuffer(buildCoverLetter());

  fs.writeFileSync(resumePath, resumeBuf);
  fs.writeFileSync(coverPath, coverBuf);

  console.log(`Wrote ${resumePath}`);
  console.log(`Wrote ${coverPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
