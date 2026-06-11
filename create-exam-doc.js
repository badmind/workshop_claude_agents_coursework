const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  LevelFormat, UnderlineType
} = require('docx');
const fs = require('fs');

const borderThin = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const borders = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

function cell(text, opts = {}) {
  const { bold=false, color="000000", bg="FFFFFF", width=4513, fontSize=22, italic=false, align=AlignmentType.LEFT } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, color, size: fontSize, font: "Arial", italics: italic })] })],
  });
}

function labelCell(text) { return cell(text, { bold: true, bg: "F2F2F2", width: 2268 }); }
function valueCell(text, width = 7088) { return cell(text, { width }); }
function twoColRow(label, value) { return new TableRow({ children: [labelCell(label), valueCell(value)] }); }

function heading(text, level = 1) {
  const sizes = { 1: 36, 2: 28, 3: 24 };
  return new Paragraph({
    spacing: { before: level === 1 ? 360 : 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: sizes[level] || 24, font: "Arial", underline: level <= 2 ? { type: UnderlineType.SINGLE } : undefined })],
  });
}

function para(text, opts = {}) {
  const { italic=false, color="000000", spacing=160 } = opts;
  return new Paragraph({ spacing: { before: 60, after: spacing }, children: [new TextRun({ text, font: "Arial", size: 22, italics: italic, color })] });
}

function redPara(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({ text: "Важно: ", bold: true, color: "CC0000", font: "Arial", size: 22 }),
      new TextRun({ text, color: "CC0000", font: "Arial", size: 22 }),
    ],
  });
}

function spacer() { return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] }); }

// ── CONTENT ────────────────────────────────────────────────────────────────

// UPDATED: цени в евро
const PROMPT_TEXT =
  'Работя по проблем в сферата на малкия бизнес / услуги. ' +
  'Проблемът е: Имам малък фитнес клуб в кв. Люлин, София ' +
  '(бул. Царица Йоанна 47), работещ вече 7 години, ' +
  'но без никакво онлайн присъствие — нямам уебсайт, намират ме само ' +
  'хора от квартала по препоръка. Искам потенциалните клиенти да могат ' +
  'да намерят информация за услугите ми (фитнес зала, кардио, групови ' +
  'тренировки, личен треньор), цените (от 5 EUR/посещение до 306 EUR/год.), ' +
  'работното време и да се запишат за безплатна пробна тренировка. ' +
  'Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения.';

const AI_SOLUTION =
  'Предложеното решение е създаването на проста едностранична уебстраница (landing page) ' +
  'за FitZone, която съдържа меню с услуги, ценоразпис в EUR, седмична програма на груповите ' +
  'тренировки, форма за запис за безплатна пробна тренировка, интерактивна Google Maps карта ' +
  'и информация за локация с работно време. Страницата помага на хора от квартала и около него, ' +
  'които търсят фитнес клуб в Google — те ще могат да намерят клуба, да разберат какво се предлага ' +
  'и на каква цена в евро, и да се запишат за пробна тренировка без да се налага да се обаждат. ' +
  'Страницата е статична (само HTML/CSS/JS), хоствана безплатно в GitHub Pages, ' +
  'лесна за поддържане и може да се надгради с онлайн резервации при нужда. ' +
  'Редизайнът следва UI/UX Pro Max Skill — Vibrant & Block-based стил с оранжево-зелена ' +
  'цветова схема (#F97316 / #22C55E) на тъмен фон, подходящ за спортни брандове. ' +
  'Решението е приложимо веднага, без техническа поддръжка и без месечен разход.';

const WEBPAGE_URL = 'https://badmind.github.io/workshop_claude_agents_coursework/';

// UPDATED: Пълно име и имейл
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
    twoColRow("Проблем", "Фитнес клуб FitZone (кв. Люлин, София) — 7 години работа, нулево онлайн присъствие. Клиентите се намират само по препоръка. Нужна е проста уебстраница с услуги, цени в EUR и форма за запис."),
  ],
});

const solutionTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Промпт", PROMPT_TEXT),
    twoColRow("Получено решение", AI_SOLUTION),
    twoColRow("Линк или screenshot", WEBPAGE_URL),
  ],
});

const webpageTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2268, 7092],
  rows: [
    twoColRow("Описание", "Едностранична HTML уебстраница за FitZone - фитнес клуб в кв. Люлин. Секции: Hero, Услуги (4 карти), Цени в EUR (4 плана), Групови тренировки (6 занятия), Форма за запис, Google Maps карта, Локация. Редизайн по UI/UX Pro Max Skill."),
    twoColRow("Технология", "HTML5 / CSS3 / JavaScript — без фреймуъркове. Хостван безплатно в GitHub Pages. Google Maps Embed API за карта."),
    twoColRow("Линк към страницата", WEBPAGE_URL),
    twoColRow("Линк към GitHub", "https://github.com/badmind/workshop_claude_agents_coursework"),
  ],
});

const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 22, color: "1C1C1C" } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [new TextRun({ text: 'Практическа изпитна задача по', bold: true, size: 32, font: "Arial" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        children: [new TextRun({ text: '"Workshop: Claude Agents for Efficient Work"', bold: true, size: 36, font: "Arial" })],
      }),
      para('Настоящото задание за изпит има за цел да провери дали можете самостоятелно да използвате Claude за формулиране на реален проблем, генериране на приложимо решение и създаване на проста уебстраница, която представя решението по ясен и убедителен начин.'),
      spacer(),
      studentTable,
      spacer(),
      redPara('Уверете се, че линковете към вашите резултати са със споделен достъп, позволяващ на проверяващите да ги отворят. Ако линкът не е достъпен, задачата няма да може да бъде оценена коректно.'),
      spacer(),
      heading('1. Избор на тема'),
      para('Изберете реален проблем от вашия живот или работа, за който искате да намерите решение с помощта на AI.'),
      para('Можете да изберете и собствена тема.', { italic: true }),
      para('Попълнете таблицата:'),
      topicTable,
      spacer(),
      heading('2. Генериране на решение с AI асистент'),
      para('Опишете проблема на AI асистента и поискайте предложение за просто, приложимо решение. В промпта включете кой има проблема, какво му пречи и какъв тип резултат очаквате.'),
      para('Примерен промпт: „Работя по проблем в сферата на [сфера]. Проблемът е: [описание]. Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения."', { italic: true }),
      para('Попълнете таблицата (поставете реалния промпт и резултата):'),
      solutionTable,
      spacer(),
      heading('3. Създаване на уебстраница с AI'),
      para('Използвайте Claude, за да генерирате проста HTML уебстраница, която представя решението. Страницата трябва да съдържа: заглавие, описание на проблема и решението, и поне един призив за действие (бутон, форма или линк).'),
      para('Публикувайте страницата (GitHub Pages, Netlify и т.н.) и поставете линка по-долу:'),
      webpageTable,
      spacer(),
      heading('Предаване'),
      para('Проектът се качва като PDF или DOCX в Drive и се споделя линк в сайта, секция „Редовен / Поправителен изпит".'),
      new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [
          new TextRun({ text: 'Редовен изпит: ', bold: true, font: "Arial", size: 22 }),
          new TextRun({ text: 'до 14 юни 2026 г., 21:59 ч.', font: "Arial", size: 22 }),
        ],
      }),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [
          new TextRun({ text: 'Поправителен изпит: ', bold: true, font: "Arial", size: 22 }),
          new TextRun({ text: 'до 21 юни 2026 г., 21:59 ч.', font: "Arial", size: 22 }),
        ],
      }),
      spacer(),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('Exam-Claude-Webpage.docx', buffer);
  console.log('Done');
});
