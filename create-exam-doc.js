const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, UnderlineType,
  LevelFormat, HeadingLevel
} = require('docx');
const fs = require('fs');

// ── HELPERS ───────────────────────────────────────────────────────
const borderThin = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

function cell(text, opts = {}) {
  const { bold=false, color="1C1C1C", bg="FFFFFF", width=7092, fontSize=22, italic=false } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({ children: [new TextRun({ text, bold, color, size: fontSize, font: "Arial", italics: italic })] })],
  });
}

function labelCell(text) { return cell(text, { bold: true, bg: "F5F5F5", color: "333333", width: 2268 }); }
function valueCell(text) { return cell(text, { width: 7092 }); }
function twoColRow(label, value) { return new TableRow({ children: [labelCell(label), valueCell(value)] }); }

function h1(text) {
  return new Paragraph({
    spacing: { before: 400, after: 160 },
    children: [new TextRun({ text, bold: true, size: 36, font: "Arial", color: "1C1C1C" })],
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: "1C1C1C",
      underline: { type: UnderlineType.SINGLE } })],
  });
}

function para(text, opts = {}) {
  const { italic=false, color="333333", before=60, after=120 } = opts;
  return new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text, font: "Arial", size: 22, italics: italic, color })],
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "• ", font: "Arial", size: 22, color: "F97316" }),
      new TextRun({ text, font: "Arial", size: 22, color: "333333" }),
    ],
  });
}

function redNote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({ text: "Важно: ", bold: true, color: "CC0000", font: "Arial", size: 22 }),
      new TextRun({ text, color: "CC0000", font: "Arial", size: 22 }),
    ],
  });
}

function spacer(before = 80, after = 80) {
  return new Paragraph({ spacing: { before, after }, children: [new TextRun("")] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "EEEEEE" } },
    children: [new TextRun("")],
  });
}

// ── CONTENT ───────────────────────────────────────────────────────

const PROMPT_TEXT =
  'Работя по проблем в сферата на малкия бизнес / фитнес услуги. ' +
  'Проблемът е: Имам малък фитнес клуб FitZone в кв. Люлин, София ' +
  '(бул. Царица Йоанна 47), работещ вече 7 години, без никакво онлайн ' +
  'присъствие — нямам уебсайт, намират ме само хора от квартала по препоръка. ' +
  'Искам потенциалните клиенти да намират информация за услугите ми ' +
  '(фитнес зала, кардио, групови тренировки, личен треньор), ' +
  'цените (от €5/посещение до €306/год.), работното време и да се ' +
  'запишат за безплатна пробна тренировка онлайн. ' +
  'Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения.';

const AI_SOLUTION =
  'Решението е едностранична уебстраница (landing page) за FitZone с секции: ' +
  'Hero с ключово послание, Услуги (4 карти), Цени в EUR (4 плана), ' +
  'Седмична програма на груповите тренировки, Форма за безплатна пробна тренировка, ' +
  'Google Maps с реалната локация и контакти. ' +
  'Страницата помага на хора от квартала, търсещи фитнес клуб в Google — ' +
  'виждат всичко необходимо и се записват без да се обаждат. ' +
  'Редизайнът следва UI/UX Pro Max Skill (Vibrant & Block-based: оранжево #F97316 + ' +
  'зелено #22C55E на тъмен фон), оптимизирана за мобилни устройства с hamburger меню. ' +
  'Хоствана безплатно в GitHub Pages, без месечен разход.';

const HTML_PROMPT =
  'Направи ми красива, responsive едностранична HTML страница за фитнес клуб FitZone. ' +
  'Стил: Vibrant & Block-based, тъмен фон #0F172A, оранжев акцент #F97316, зелен CTA #22C55E. ' +
  'Секции: Hero с H1 "Тренирай без компромис" + статистики, Услуги (4 карти), ' +
  'Цени в EUR (4 абонаментни плана), Групови тренировки (6 занятия), ' +
  'Форма за пробна тренировка, Google Maps embed, Footer. ' +
  'Мобилна оптимизация с hamburger меню. Шрифт: Barlow Condensed + Barlow от Google Fonts.';

const PAGE_URL = 'https://badmind.github.io/workshop_claude_agents_coursework/';
const GITHUB_URL = 'https://github.com/badmind/workshop_claude_agents_coursework';

// ── TABLES ────────────────────────────────────────────────────────

const studentTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Име", "Божидар Пачовски"),
    twoColRow("Username", "badmind"),
    twoColRow("Email", "pachovski@gmail.com"),
  ],
});

const topicTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Избрана сфера", "Малък бизнес / Фитнес услуги"),
    twoColRow("Проблем",
      "Фитнес клуб FitZone (кв. Люлин, София) работи 7 години без онлайн присъствие. " +
      "Клиентите го намират само по препоръка от квартала. Необходима е уебстраница " +
      "с услуги, цени в EUR, форма за запис и локация."),
  ],
});

const solutionTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Промпт", PROMPT_TEXT),
    twoColRow("Получено решение", AI_SOLUTION),
    twoColRow("Линк или screenshot", PAGE_URL),
  ],
});

const htmlTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Промпт за HTML", HTML_PROMPT),
    twoColRow("Линк към страница", PAGE_URL),
  ],
});

const techTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Технология", "HTML5 / CSS3 / JavaScript — без фреймуъркове"),
    twoColRow("Хостинг", "GitHub Pages (безплатно)"),
    twoColRow("Дизайн система", "UI/UX Pro Max Skill — Vibrant & Block-based"),
    twoColRow("GitHub репо", GITHUB_URL),
  ],
});

// ── DOCUMENT ──────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: "1C1C1C" } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [

      // ── ЗАГЛАВИЕ
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 160 },
        children: [new TextRun({ text: "Практическа изпитна задача по", bold: true, size: 32, font: "Arial" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: '"Workshop: Claude Agents for Efficient Work"', bold: true, size: 36, font: "Arial" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        children: [new TextRun({ text: "Редовен изпит — юни 2026", size: 22, font: "Arial", color: "888888" })],
      }),

      para("Настоящото задание има за цел да провери дали можете самостоятелно да използвате Claude за формулиране на реален проблем, генериране на приложимо решение и създаване на проста уебстраница, която представя решението по ясен и убедителен начин."),

      spacer(),
      studentTable,
      spacer(),

      redNote("Уверете се, че линковете към вашите резултати са със споделен достъп, позволяващ на проверяващите да ги отворят. Ако линкът не е достъпен, задачата няма да може да бъде оценена коректно."),

      divider(),

      // ── СЕКЦИЯ 1
      h2("1. Избор на сфера и формулиране на проблем"),
      para("Изберете реален проблем от бизнес, организация, екип, общност или ваша професионална среда. Проблемът трябва да е конкретен и да може да бъде обяснен в 1-2 изречения."),
      spacer(40, 80),
      topicTable,

      divider(),

      // ── СЕКЦИЯ 2
      h2("2. Генериране на решение с AI асистент"),
      para("Описах проблема на Claude и поисках предложение за просто, приложимо решение. В промпта включих кой има проблема, какво му пречи и какъв тип резултат очаквам."),
      para('Използван промпт формат: „Работя по проблем в сферата на [сфера]. Проблемът е: [описание]. Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения."', { italic: true }),
      spacer(40, 80),
      solutionTable,

      divider(),

      // ── СЕКЦИЯ 3
      h2("3. Създаване на уебстраница с AI"),
      para("Използвах Claude за генериране на HTML страница, представяща решението. Страницата е разбираема за потенциален потребител дори без технически познания."),
      spacer(40, 40),
      para("Изпълнени минимални изисквания:", { color: "1C1C1C" }),
      bullet("Заглавие с кратко и ясно послание (\"Тренирай без компромис\")"),
      bullet("Кратко обяснение на проблема и решението"),
      bullet("4 ключови услуги и 4 ценови плана в EUR"),
      bullet("Форма за запис — призив за действие"),
      bullet("Четим дизайн: Vibrant & Block-based, тъмен фон, оранжев акцент"),
      bullet("Мобилна оптимизация с hamburger меню"),
      bullet("Google Maps embed с реалната локация"),
      spacer(80, 80),
      htmlTable,
      spacer(40, 80),
      techTable,

      divider(),

      // ── КРИТЕРИИ
      h2("4. Критерии за оценяване"),
      bullet("Ясно формулиран реален проблем и целева аудитория"),
      bullet("Адекватно и приложимо решение, генерирано и доразвито с AI"),
      bullet("Качество на промптовете и видим процес на итерация"),
      bullet("Завършена HTML страница с нужните секции"),
      bullet("Яснота, подреденост и споделен достъп до финалния резултат"),
      spacer(80, 80),
      para("Няма изискване страницата да бъде сложна или технически перфектна. Важни са ясната идея, правилното използване на AI асистента и завършеното представяне на решението.", { italic: true }),

      divider(),

      // ── ПРЕДАВАНЕ
      h2("Предаване"),
      para("Проектът се качва като PDF или DOCX в Google Drive и се споделя линк в сайта, секция Редовен / Поправителен изпит."),
      spacer(60, 40),
      new Paragraph({
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "Редовен изпит: ", bold: true, font: "Arial", size: 22 }),
          new TextRun({ text: "до 14 юни 2026 г., 21:59 ч.", font: "Arial", size: 22 }),
        ],
      }),
      new Paragraph({
        spacing: { before: 40, after: 80 },
        children: [
          new TextRun({ text: "Поправителен изпит: ", bold: true, font: "Arial", size: 22 }),
          new TextRun({ text: "до 21 юни 2026 г., 21:59 ч.", font: "Arial", size: 22 }),
        ],
      }),

      spacer(160, 80),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 0 },
        children: [new TextRun({ text: "© SoftUni – https://softuni.bg", font: "Arial", size: 18, color: "AAAAAA" })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('Exam-Claude-Webpage.docx', buffer);
  console.log('Done');
});
