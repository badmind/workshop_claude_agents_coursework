const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  UnderlineType, PageBreak
} = require('docx');
const fs = require('fs');

const bLine  = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const bAll   = { top: bLine, bottom: bLine, left: bLine, right: bLine };

function run(text, opts = {}) {
  return new TextRun({ text, font: "Calibri", size: 22, ...opts });
}
function p(children, opts = {}) {
  if (typeof children === 'string') children = [run(children)];
  return new Paragraph({ spacing: { before: 60, after: 120 }, children, ...opts });
}
function pItalic(text) {
  return p([run(text, { italics: true, color: "555555" })]);
}
function pBold(text) {
  return p([run(text, { bold: true })]);
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 160 },
    children: [run(text, { bold: true, size: 28, color: "1F4E79" })],
  });
}
function redNote(text) {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    children: [
      run("Важно: ", { bold: true, color: "CC0000", size: 22 }),
      run(text, { bold: true, color: "CC0000", size: 22 }),
    ],
  });
}
function bullet(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 360 },
    children: [run("• ", { color: "1F4E79" }), run(text)],
  });
}
function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [run("")] });
}
function tRow(label, value, valueBold = false) {
  return new TableRow({ children: [
    new TableCell({
      borders: bAll,
      width: { size: 2500, type: WidthType.DXA },
      shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [p([run(label, { bold: true })])],
    }),
    new TableCell({
      borders: bAll,
      width: { size: 6860, type: WidthType.DXA },
      shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [p([run(value, { bold: valueBold })])],
    }),
  ]});
}
function table(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2500, 6860],
    rows: rows.map(r => tRow(r[0], r[1], r[2] || false)),
  });
}

// ── ДАННИ ────────────────────────────────────────────────────────

const PAGE_URL   = 'https://badmind.github.io/workshop_claude_agents_coursework/';
const GITHUB_URL = 'https://github.com/badmind/workshop_claude_agents_coursework';

const PROMPT_TEXT =
  'Работя по проблем в сферата на малкия бизнес и фитнес услуги. ' +
  'Проблемът е: имам малък фитнес клуб FitZone в кв. Люлин, София (бул. Царица Йоанна 47), ' +
  'работещ вече 7 години, без никакво онлайн присъствие — нямам уебсайт и клиентите ме намират само по препоръка. ' +
  'Искам потенциалните клиенти да намират информация за услугите ' +
  '(фитнес зала, кардио, групови тренировки, личен треньор), ' +
  'цените в EUR и да се записват за безплатна пробна тренировка онлайн. ' +
  'Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения.';

const AI_SOLUTION =
  'Решението е едностранична уебстраница (landing page) за FitZone. ' +
  'Страницата съдържа: Hero секция с послание "Тренирай без компромис" и 4 статистики, ' +
  'блок с 4 услуги, ценоразпис в EUR (4 плана: 5 EUR/посещение до 306 EUR/год.), ' +
  'седмична програма на груповите тренировки, форма за безплатна пробна тренировка, ' +
  'Google Maps embed с реалната локация и контакти. ' +
  'Страницата помага на хора от квартала, търсещи фитнес клуб в Google — ' +
  'виждат всичко необходимо и се записват без да се налага да се обаждат. ' +
  'Хоствана безплатно в GitHub Pages — без месечен разход.';

const HTML_PROMPT =
  'Направи ми красива, responsive едностранична HTML страница за фитнес клуб FitZone. ' +
  'Тъмен фон #0F172A, оранжев акцент #F97316, зелен CTA #22C55E, шрифт Barlow Condensed + Barlow от Google Fonts. ' +
  'Секции: Hero с "Тренирай без компромис" + статистики, Услуги (4 карти), ' +
  'Цени в EUR (4 плана), Групови тренировки (6 занятия), Форма за пробна тренировка, ' +
  'Google Maps embed, Footer. Мобилна оптимизация с hamburger меню. ' +
  'Следвай UI/UX Pro Max Skill — стил Vibrant & Block-based.';

const REDESIGN_PROMPT =
  'Направи редизайн на страницата спрямо UI/UX Pro Max Skill ' +
  '(https://github.com/nextlevelbuilder/ui-ux-pro-max-skill). ' +
  'Скилът препоръча за фитнес продукт: Vibrant & Block-based стил, ' +
  'цветова схема Energy Orange #F97316 + Success Green #22C55E на тъмен фон, ' +
  'шрифт Barlow Condensed (потвърден best-for: sports, fitness, gyms), ' +
  '48px+ block gaps, animated diagonal pattern, bold hover color-shift. ' +
  'Приложи всички правила за accessibility (WCAG AA), touch targets 44px, ' +
  'prefers-reduced-motion, aria-labels, видими focus states.';

const ITERATION_PROMPT =
  'Оптимизирай страницата за мобилни устройства: ' +
  'добави hamburger меню с animated drawer, ' +
  'поправи overflow на heading текстове (word-break), ' +
  'направи всички бутони full-width на мобилен, ' +
  'добави breakpoints 375/480/768/1024px, ' +
  'exam banner да се наредва вертикално на мобилен.';

const RATING_PROMPT =
  'Добави бутон "Оцени" до бутоните за документа. ' +
  'При клик отваря modal с: оценки по SoftUni скала (2.00/3.00/4.00/5.00/5.50/6.00), ' +
  'оценки >= 5.00 маркирани в зелено (дават право на сертификат), ' +
  'поле за коментар, поле за CC имейл, анимация с хвърчащи писма при изпращане, ' +
  'прогрес бар и success съобщение. ' +
  'Свържи с EmailJS (service_0bdu17s / template_i7idtwm) за реални имейли до pachovski@gmail.com. ' +
  'Темплейтът да е с дизайна на страницата: тъмен фон, оранжево/зелено, ' +
  'описание на оценката, критерии за оценяване, cert badge.';

// ── ДОКУМЕНТ ─────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [{
      id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
      run: { size: 28, bold: true, color: "1F4E79", font: "Calibri" },
      paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 },
    }],
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
        spacing: { before: 0, after: 120 },
        children: [run("Практическа изпитна задача по", { bold: true, size: 30 })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [run('"Workshop: Claude Agents for Efficient Work"', { bold: true, size: 34 })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 320 },
        children: [run("Редовен изпит — юни 2026", { size: 20, color: "888888" })],
      }),

      p("Настоящото задание за изпит има за цел да провери дали можете самостоятелно да използвате Claude за формулиране на реален проблем, генериране на приложимо решение и създаване на проста уебстраница, която представя решението по ясен и убедителен начин."),
      spacer(),

      table([
        ["Име",      "Божидар Пачовски"],
        ["Username", "badmind"],
        ["Email",    "pachovski@gmail.com"],
      ]),
      spacer(),

      redNote("Уверете се, че линковете към вашите резултати са със споделен достъп, позволяващ на проверяващите да ги отворят. Ако линкът не е достъпен, задачата няма да може да бъде оценена коректно."),
      spacer(),

      // ── СЕКЦИЯ 1
      h2("1. Избор на сфера и формулиране на проблем"),
      p("Изберете реален проблем от бизнес, организация, екип, общност или ваша професионална среда. Проблемът трябва да е конкретен и да може да бъде обяснен в 1-2 изречения."),
      p("Можете да използвате една от примерните теми: местен бизнес без онлайн присъствие, организиране на събитие, онбординг на нов служител, често задавани въпроси от клиенти, промотиране на нов продукт или услуга, здравословен навик в екипа, местна общност/квартал. Можете да изберете и собствена тема."),
      p("Попълнете таблицата:"),
      spacer(),
      table([
        ["Избрана сфера", "Малък бизнес / Фитнес услуги"],
        ["Проблем",
          "Фитнес клуб FitZone (кв. Люлин, София) работи 7 години без никакво онлайн присъствие. " +
          "Клиентите го намират само по препоръка от квартала. " +
          "Необходима е уебстраница с услуги, цени в EUR, форма за запис и Google Maps локация."],
      ]),

      // ── СЕКЦИЯ 2
      h2("2. Генериране на решение с AI асистент"),
      p("Описах проблема на Claude и поисках предложение за просто, приложимо решение."),
      pItalic('Промпт формат: „Работя по проблем в сферата на [сфера]. Проблемът е: [описание]. Предложи ми просто решение, обясни на кого помага и дай идеята в 5-6 изречения."'),
      p("Попълнете таблицата (поставете реалния промпт и резултата):"),
      spacer(),
      table([
        ["Промпт",             PROMPT_TEXT],
        ["Получено решение",   AI_SOLUTION],
        ["Линк или screenshot", PAGE_URL],
      ]),

      // ── СЕКЦИЯ 3
      h2("3. Създаване на уебстраница с AI"),
      p("Използвах Claude за итеративно генериране и подобряване на HTML страница. Процесът включваше множество промпта и подобрения."),
      pBold("Минимални изисквания към страницата:"),
      bullet("заглавие с кратко и ясно послание;"),
      bullet("кратко обяснение на проблема и решението;"),
      bullet("3 ключови ползи;"),
      bullet("бутон или текст с призив за действие;"),
      bullet("четим дизайн с подходящи цветове и подредба."),
      spacer(),

      pBold("Изпълнени изисквания и допълнителни функционалности:"),
      bullet("Hero секция: \"Тренирай без компромис\" + 4 статистики (7г, 340+ члена, 06:00, 4 треньора)"),
      bullet("4 услуги (тежести, кардио, групови, личен треньор) + 4 ценови плана в EUR"),
      bullet("Седмична програма на групови тренировки (6 занятия)"),
      bullet("Форма за безплатна пробна тренировка с validation и success state"),
      bullet("Google Maps embed с реалната локация (бул. Царица Йоанна 47, Люлин)"),
      bullet("Редизайн по UI/UX Pro Max Skill (Vibrant & Block-based) с оценка от скила"),
      bullet("Пълна мобилна оптимизация: hamburger меню, breakpoints 375/480/768/1024px"),
      bullet("Бутон \"Оцени\" с modal по SoftUни скала (2.00-6.00), cert badge за >=5.00"),
      bullet("EmailJS интеграция за реални имейли с HTML темплейт в дизайна на страницата"),
      bullet("Изпитен документ достъпен директно от страницата (Свали DOCX / Виж в GitHub)"),
      spacer(),

      pItalic('Промпт за HTML: „' + HTML_PROMPT + '"'),
      spacer(),
      pItalic('Промпт за редизайн: „' + REDESIGN_PROMPT + '"'),
      spacer(),
      pItalic('Промпт за мобилна оптимизация: „' + ITERATION_PROMPT + '"'),
      spacer(),
      pItalic('Промпт за оценъчен бутон: „' + RATING_PROMPT + '"'),
      spacer(),

      p("Попълнете таблицата:"),
      spacer(),
      table([
        ["Промпт за HTML",                             HTML_PROMPT],
        ["Линк към страница / HTML файл / screenshot", PAGE_URL],
      ]),
      spacer(),
      table([
        ["GitHub репо",      GITHUB_URL],
        ["Технология",       "HTML5 / CSS3 / JavaScript — без фреймуъркове, single-file"],
        ["Хостинг",          "GitHub Pages (безплатно, автоматично при push към main)"],
        ["Дизайн система",   "UI/UX Pro Max Skill — Vibrant & Block-based (nextlevelbuilder/ui-ux-pro-max-skill)"],
        ["Цени",             "EUR — еднократно €5, месечен €28, месечен Plus €40, годишен €306"],
        ["Имейл интеграция", "EmailJS (service_0bdu17s) — реални имейли при оценяване"],
        ["Оценъчна система", "SoftUni скала 2.00-6.00, cert badge за >=5.00, HTML email темплейт"],
      ]),

      // ── СЕКЦИЯ 4
      h2("4. Критерии за оценяване"),
      p("Задачата ще бъде оценена по следните критерии:"),
      bullet("ясно формулиран реален проблем и целева аудитория;"),
      bullet("адекватно и приложимо решение, генерирано и доразвито с AI;"),
      bullet("качество на промптовете и видим процес на итерация;"),
      bullet("завършена HTML страница с нужните секции;"),
      bullet("яснота, подреденост и споделен достъп до финалния резултат."),
      spacer(),
      pItalic("Няма изискване страницата да бъде сложна или технически перфектна. Важни са ясната идея, правилното използване на AI асистента и завършеното представяне на решението."),
      spacer(),

      // ── FOOTER
      new Paragraph({
        spacing: { before: 400, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" } },
        children: [
          run("Последвайте ни: ", { color: "888888", size: 20 }),
          run("© SoftUni – https://softuni.bg", { color: "888888", size: 20 }),
          run("  |  Защитен документ. Моля, не копирайте без разрешение.", { color: "AAAAAA", size: 18 }),
        ],
      }),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Exam-Claude-Webpage.docx', buf);
  console.log('Done');
});
