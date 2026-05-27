/**
 * Generate Trimble-tailored cover letter (DOCX).
 * Run from repo root: node resumes/generate_trimble_cover_letter.js
 *
 * Job context is summarized in jobs/trimble.md (official apply link).
 */
const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
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
          center("https://www.linkedin.com/in/nmyersdev/", {
            size: 20,
          }),
          body(today, { after: 200 }),
          body("Hiring Team", { bold: true }),
          body("Trimble — Software Engineer", { after: 200 }),
          body(
            "Trimble’s mission—connecting the physical and digital worlds to improve how essential work gets done—resonates with the product engineering I want next. " +
              "I am applying for the Software Engineer role because the posting emphasizes shipping robust features across the stack (relational data and web APIs through to the front-end experience), working in modern JavaScript frameworks alongside object-oriented services, and collaborating with product and QA in mature CI/CD environments. " +
              "I am based in the Portland metro and would welcome the chance to support Trimble’s Oregon-based teams in person when it helps delivery, while continuing to collaborate effectively in distributed settings."
          ),
          body(
            "At Resource Data I spent most of my time on user-facing web applications in React and TypeScript—translating requirements into maintainable UI, partnering with stakeholders through delivery, and supporting production when users depended on the system. " +
              "I routinely worked with .NET / .NET Core services, REST and GraphQL APIs, and SQL-backed data for complex forms, tables, and map-driven workflows (Washington DNR burn permitting; WSAC Career Launch). " +
              "On IdeaRoom/American Steel I built configurable ordering flows and Storybook-backed components for a manufacturing sales experience—domains where accuracy and traceability matter, not unlike digital construction and operations tooling. " +
              "Across Bel/Cinch and Epic Charter School I paired front-end work with structured API testing (Postman) and Node.js/TypeScript automation to catch regressions early and keep releases dependable."
          ),
          body(
            "Trimble’s focus on disciplined engineering and practical use of AI-assisted development matches how I already work: I use tools like GitHub Copilot for boilerplate, refactors, and tests, and I am eager to deepen that practice with a team that treats AI as part of how quality and velocity scale together. " +
              "I communicate clearly in writing, take feedback well in code review, and care about outcomes for non-technical users—habits I built in consulting and earlier customer-facing technical roles."
          ),
          body(
            "Thank you for your time and consideration. I would welcome a conversation about how I can contribute to Trimble’s software products, and I am happy to share more detail on any project above or walk through code samples."
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

  const coverPath = path.join(
    COVER_LETTERS_DIR,
    "Nick_Myers_Cover_Letter_Trimble.docx"
  );

  const coverBuf = await Packer.toBuffer(buildCoverLetter());
  fs.writeFileSync(coverPath, coverBuf);
  console.log(`Wrote ${coverPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
