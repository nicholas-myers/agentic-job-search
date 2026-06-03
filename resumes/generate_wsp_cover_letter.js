/**
 * Generate WSP-tailored cover letter (DOCX).
 * Run from repo root: node resumes/generate_wsp_cover_letter.js
 *
 * Job context: jobs/wsp.md
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
          center("Portland, OR · Open to Vancouver, WA / hybrid", { size: 22 }),
          center("https://www.linkedin.com/in/nmyersdev/", {
            size: 20,
          }),
          body(today, { after: 200 }),
          body("Hiring Team", { bold: true }),
          body("WSP — Senior Application Developer", { after: 200 }),
          body(
            "WSP’s work shaping infrastructure and communities is the kind of impact I want my software to support. " +
              "I am applying for the Senior Application Developer role in Software Solutions because it combines full-stack delivery on the Microsoft stack with polished, responsive UIs, Agile collaboration with project teams, and applications that help clients run complex programs—from reporting and workflows to map-driven tools."
          ),
          body(
            "At Resource Data I delivered production software for public-sector and commercial clients using React, TypeScript, Angular, and .NET / .NET Core (MVC), with SQL Server-backed data throughout. " +
              "For Washington DNR I improved burn-permit forms, data views, and Esri map workflows; for WSAC Career Launch and related programs I shipped full-stack features where clarity and reliability mattered to non-technical users. " +
              "I built Angular and .NET components for Alaska Housing Finance Corporation planning applications, configurable ordering and document flows for manufacturing clients (React, TypeScript, Storybook), and strengthened GraphQL and .NET APIs with structured Postman testing. " +
              "I have mentored developers, partnered with stakeholders to translate shifting priorities into workable backlogs, and worked in Git-based delivery with disciplined pull requests—habits that match your emphasis on Agile teamwork and maintainable custom applications."
          ),
          body(
            "I am based in the Portland metro (PST), a short commute from Vancouver, and I am authorized to work in the United States without sponsorship. " +
              "I would welcome the chance to grow further with Blazor, SharePoint, and Power BI–style reporting on a team that values strong UI/UX and data interpretation for project controls and technology solutions."
          ),
          body(
            "Thank you for your time and consideration. I am happy to discuss any project above in more detail or walk through code samples."
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
    "Nick_Myers_Cover_Letter_WSP.docx"
  );

  const coverBuf = await Packer.toBuffer(buildCoverLetter());
  fs.writeFileSync(coverPath, coverBuf);
  console.log(`Wrote ${coverPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
