// generate_report.js — Gulf Energy Development Weekly Progress Report Generator
// Usage: node generate_report.js data.json output.docx

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Footer, AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Constants ───────────────────────────────────────────────────────────────
const NAVY       = '1F3864';
const WHITE      = 'FFFFFF';
const GREEN_ROW  = 'E2EFDA';
const RED_ROW    = 'FFE0E0';
const GRAY_TEXT  = '808080';

// A4 page in DXA (1440 DXA = 1 inch / 25.4 mm)
// 11906 × 16838  →  margins 1417 all sides  →  content width ≈ 9072
const PAGE_W    = 11906;
const PAGE_H    = 16838;
const MARGIN    = 1417;
const CONTENT_W = PAGE_W - MARGIN * 2;  // 9072

// Summary table column widths (must sum to CONTENT_W = 9072)
const COL_ID       = 1200;
const COL_NAME     = 3500;
const COL_PDF      = 1200;
const COL_CONCERNS = 1586;
const COL_ACT      = 1586;
// sum = 9072

// ── Helpers ──────────────────────────────────────────────────────────────────

function navyShading() {
  return { fill: NAVY, type: ShadingType.CLEAR, color: NAVY };
}

function clearShading(fill) {
  return { fill, type: ShadingType.CLEAR, color: fill };
}

function cellBorders(color = 'FFFFFF') {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

function sectionHeader(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    shading: navyShading(),
    spacing: { before: 0, after: 240 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 36,          // 18pt
        color: WHITE,
        font: 'Calibri',
      }),
    ],
  });
}

function projectHeading(project) {
  return new Paragraph({
    spacing: { before: 240, after: 60 },
    children: [
      new TextRun({
        text: `${project.name}  (${project.id})`,
        bold: true,
        size: 28,          // 14pt
        color: NAVY,
        font: 'Calibri',
      }),
    ],
  });
}

function navyDivider() {
  return new Paragraph({
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 },
    },
    spacing: { before: 0, after: 120 },
    children: [],
  });
}

function bulletItem(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text, font: 'Calibri', size: 22 }),
    ],
  });
}

function pageBreakParagraph() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ── Cover page ────────────────────────────────────────────────────────────────

function buildCoverPage(data) {
  const monthNames = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
  ];
  const d = new Date(data.date);
  const dateStr = `${d.getDate()} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;

  return [
    // Top navy accent bar (simulated by shaded paragraph)
    new Paragraph({
      alignment: AlignmentType.CENTER,
      shading: navyShading(),
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: ' ', size: 48, color: WHITE, font: 'Calibri' }),
      ],
    }),

    // Spacer
    new Paragraph({ spacing: { before: 2880, after: 0 }, children: [] }),

    // Company title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
      children: [
        new TextRun({
          text: 'GULF ENERGY DEVELOPMENT',
          bold: true,
          size: 56,          // 28pt
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),

    // Subtitle
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
      children: [
        new TextRun({
          text: 'Project Progress Report',
          bold: true,
          size: 40,          // 20pt
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),

    // Week / Year
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
      children: [
        new TextRun({
          text: `Week ${data.week} / ${data.year}`,
          size: 32,          // 16pt
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),

    // Date
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 480 },
      children: [
        new TextRun({
          text: dateStr,
          size: 28,          // 14pt
          color: NAVY,
          font: 'Calibri',
        }),
      ],
    }),

    // Divider line
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 },
      },
      spacing: { before: 0, after: 240 },
      children: [],
    }),

    // Footer note
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 0 },
      children: [
        new TextRun({
          text: 'Auto-generated by Gulf Dashboard',
          size: 20,          // 10pt
          color: GRAY_TEXT,
          italics: true,
          font: 'Calibri',
        }),
      ],
    }),

    // Page break
    pageBreakParagraph(),
  ];
}

// ── Section 1 — Progress Summary ──────────────────────────────────────────────

function buildSummaryTable(projects) {
  const headerCell = (text, width) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: navyShading(),
      borders: cellBorders(NAVY),
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text, bold: true, size: 20, color: WHITE, font: 'Calibri' }),
          ],
        }),
      ],
    });

  const dataCell = (text, width, fill, align = AlignmentType.LEFT) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: clearShading(fill),
      borders: cellBorders('CCCCCC'),
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: align,
          children: [
            new TextRun({ text, size: 20, font: 'Calibri' }),
          ],
        }),
      ],
    });

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Project ID',  COL_ID),
      headerCell('Project Name', COL_NAME),
      headerCell('PDF Status',  COL_PDF),
      headerCell('Concerns',    COL_CONCERNS),
      headerCell('Activities',  COL_ACT),
    ],
  });

  const dataRows = projects.map(p => {
    const fill = p.pdf_found ? GREEN_ROW : RED_ROW;
    const pdfText = p.pdf_found ? '✓ Found' : '✗ Missing';
    return new TableRow({
      children: [
        dataCell(p.id,                               COL_ID,       fill),
        dataCell(p.name,                             COL_NAME,     fill),
        dataCell(pdfText,                            COL_PDF,      fill, AlignmentType.CENTER),
        dataCell(String(p.concerns.length),          COL_CONCERNS, fill, AlignmentType.CENTER),
        dataCell(String(p.activities.length),        COL_ACT,      fill, AlignmentType.CENTER),
      ],
    });
  });

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [COL_ID, COL_NAME, COL_PDF, COL_CONCERNS, COL_ACT],
    rows: [headerRow, ...dataRows],
  });
}

function buildSection1(projects) {
  return [
    sectionHeader('SECTION 1 — PROGRESS SUMMARY'),
    buildSummaryTable(projects),
    pageBreakParagraph(),
  ];
}

// ── Sections 2 & 3 — Concerns / Activities ────────────────────────────────────

function buildItemSection(headerText, projects, key) {
  const filtered = [...projects]
    .filter(p => p[key] && p[key].length > 0)
    .sort((a, b) => a.id.localeCompare(b.id));

  const content = [sectionHeader(headerText)];

  if (filtered.length === 0) {
    content.push(
      new Paragraph({
        spacing: { before: 240, after: 240 },
        children: [
          new TextRun({
            text: key === 'concerns'
              ? 'No concerns recorded this week.'
              : 'No activities recorded this week.',
            italics: true,
            size: 22,
            font: 'Calibri',
          }),
        ],
      })
    );
  } else {
    filtered.forEach((project, idx) => {
      content.push(projectHeading(project));
      content.push(navyDivider());
      project[key].forEach(item => content.push(bulletItem(item)));

      // Page break between projects, NOT after last
      if (idx < filtered.length - 1) {
        content.push(pageBreakParagraph());
      }
    });
  }

  content.push(pageBreakParagraph());
  return content;
}

// ── Footer ────────────────────────────────────────────────────────────────────

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        tabStops: [
          { type: TabStopType.RIGHT, position: CONTENT_W },
        ],
        border: {
          top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 1 },
        },
        children: [
          new TextRun({
            text: 'CONFIDENTIAL — Gulf Energy Development',
            size: 16,
            color: GRAY_TEXT,
            font: 'Calibri',
          }),
          new TextRun({
            text: '\t',
            font: 'Calibri',
          }),
          new TextRun({
            children: [PageNumber.CURRENT],
            size: 16,
            color: GRAY_TEXT,
            font: 'Calibri',
          }),
        ],
      }),
    ],
  });
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Usage: node generate_report.js data.json output.docx');
    process.exit(1);
  }

  const [inputPath, outputPath] = args;

  if (!fs.existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(1);
  }

  const data = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
  const { projects = [] } = data;

  // Cover page section (no footer)
  const coverSection = {
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    children: buildCoverPage(data),
  };

  // Main content section (with footer)
  const mainChildren = [
    ...buildSection1(projects),
    ...buildItemSection('SECTION 2 — CONCERNS', projects, 'concerns'),
    ...buildItemSection('SECTION 3 — NEXT PERIOD ACTIVITIES', projects, 'activities'),
  ];

  // Remove trailing page break if present
  if (
    mainChildren.length > 0 &&
    mainChildren[mainChildren.length - 1].children?.[0] instanceof PageBreak
  ) {
    mainChildren.pop();
  }

  const mainSection = {
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    footers: { default: buildFooter() },
    children: mainChildren,
  };

  const doc = new Document({
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: { indent: { left: 720, hanging: 360 } },
                run: { font: 'Symbol', size: 22 },
              },
            },
          ],
        },
      ],
    },
    sections: [coverSection, mainSection],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`Report written to: ${outputPath}`);
}

main().catch(err => {
  console.error('Error generating report:', err);
  process.exit(1);
});
