# CLAUDE.md

Контекст за Claude Code при работа по този проект.

## Какво е това

Курсова работа за **SoftUni Workshop: "Claude Agents for Efficient Work"** (редовен изпит, юни 2026).

Задачата: с помощта на AI се формулира реален проблем, генерира се решение и се прави проста уебстраница, която го представя.

**Студент:** Божидар Пачовски · **Username:** badmind · **Email:** pachovski@gmail.com

## Тема на проекта

**FitZone** — измислен малък фитнес клуб в кв. Люлин, София (бул. Царица Йоанна 47), работещ 7 години без онлайн присъствие. Решението е едностранична landing page с услуги, цени, програма, форма за запис и Google Maps.

## Структура на файловете

```
.
├── index.html              # Цялата уебстраница (single-file: HTML + CSS + JS inline)
├── Exam-Claude-Webpage.docx # Изпитният документ за предаване
├── create-exam-doc.js      # Node скрипт който генерира .docx (docx npm пакет)
├── README.md               # Описание на проекта
├── .nojekyll               # Казва на GitHub Pages да сервира HTML директно
└── CLAUDE.md               # Този файл
```

`index.html` е единственият source файл за страницата — всичко (стилове, скриптове) е inline, без build стъпка.

## Дизайн система

Страницата следва **UI/UX Pro Max Skill** (nextlevelbuilder/ui-ux-pro-max-skill), стил **Vibrant & Block-based** за фитнес/спорт:

- **Цветове:** оранжево `#F97316` (primary) + зелено `#22C55E` (CTA) на тъмен фон `#0F172A`
- **Шрифт:** Barlow Condensed (заглавия) + Barlow (текст) от Google Fonts
- **Лилаво** `#a855f7` се ползва само за бутона/модала "Оцени"
- **Layout:** блокове с 2px гутери, 64px секционни gaps, scroll-snap
- **Icons:** само inline SVG (Heroicons-стил) — без emoji
- **Мобилна оптимизация:** hamburger меню под 768px, breakpoints на 480/768/1024px

При промени по дизайна — спазвай тези токени, не въвеждай нови цветове.

## CSS променливи (в `:root`)

Всички цветове са CSS променливи. Не hardcode-вай hex стойности в компонентите — ползвай `var(--color-primary)`, `var(--color-accent)` и т.н.

## Ключови секции в index.html

1. Nav (с hamburger за мобилен) + mobile drawer меню
2. Hero — H1 "Тренирай без компромис" + 4 статистики
3. **Exam banner** — лента с бутони "Свали DOCX" / "Виж в GitHub" / "Оцени"
4. Услуги (4 карти) · Цени (4 плана в EUR) · Програма (6 занятия)
5. Форма за пробна тренировка
6. Google Maps embed (бул. Царица Йоанна 47)
7. Footer

## Бутон "Оцени" (rate modal)

Лилав бутон в exam banner-а отваря modal с:
- Избор на оценка по SoftUni скала: 2 (Слаб) → 6 (Отличен+)
- Поле за коментар + поле за CC имейл
- При изпращане: анимация с хвърчащи писма → прогрес бар → success екран
- Отваря mailto: до pachovski@gmail.com с форматирано тяло

**Важно при редакция на JS:** модалът HTML трябва да е **преди** `<script>` тага в DOM, иначе `getElementById` връща null. Деклариране на променливи (`const navEl`) трябва да е **преди** употребата им.

## Често срещани грешки (научени)

- **Кирилица в template literals** в `create-exam-doc.js`: внимавай с кавички — българските „кавички" чупят JS string-овете. Ползвай обикновени `"` или ги избягвай.
- **JS грешка спира целия script** — ако модалът спре да работи, провери конзолата за `ReferenceError`/`SyntaxError` някъде по-горе в скрипта.
- **GitHub Pages кешира** — след push изчакай 1-2 мин и hard refresh (Ctrl+Shift+R).

## Команди

```bash
# Регенериране на изпитния .docx след промяна на create-exam-doc.js
npm install docx          # първоначално
node create-exam-doc.js   # създава Exam-Claude-Webpage.docx

# Локален преглед на страницата
python3 -m http.server 8000   # после отвори http://localhost:8000

# Деплой — просто push към main, GitHub Pages се обновява автоматично
git add -A && git commit -m "..." && git push origin main
```

## Деплой

- **Хостинг:** GitHub Pages, branch `main`, root path
- **Live URL:** https://badmind.github.io/workshop_claude_agents_coursework/
- **Repo:** https://github.com/badmind/workshop_claude_agents_coursework
- Pages се обновява автоматично при всеки push към `main` (~1-2 мин)

## Стил на работа (предпочитания на Божидар)

- Не е програмист — предпочита пълни файлове и ясни инструкции какво къде се сменя
- Комуникация на български
- Кратки, стъпка-по-стъпка обяснения
- При проблем — проактивен подход, провери конзолата, тествай преди да кажеш "готово"
