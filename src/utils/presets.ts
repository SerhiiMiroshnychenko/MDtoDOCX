export interface MarkdownPreset {
  id: string;
  name: string;
  emoji: string;
  category: string;
  content: string;
  title: string;
  subtitle: string;
}

export const DOCUMENT_PRESETS: MarkdownPreset[] = [
  {
    id: 'business_proposal',
    name: 'Комерційна пропозиція (Business Proposal)',
    emoji: '💼',
    category: 'Бізнес',
    title: 'КОМЕРЦІЙНА ПРОПОЗИЦІЯ',
    subtitle: 'Впровадження автоматизованих рішень документообігу',
    content: `# Комерційна пропозиція: Системи автоматизації

Дякуємо за інтерес до наших послуг. Нижче наведено детальний план впровадження інноваційних рішень документообігу для вашої організації.

---

## 1. Огляд проекту

Наша мета — оптимізувати процес формування бізнес-документів для прискорення закриття угод та підвищення командної продуктивності на **35%**.

### Ключові цілі
* **Конвертація документів**: Переведення Markdown у хмарний DOCX за 1 клік.
* **Централізація шаблонів**: Зберігання фірмових шпаргалок.
* **Повна автоматизація**: Запобігання помилкам ручного копіювання.

---

## 2. Календарний план впровадження

Наступний графік демонструє етапи розробки та навчання:

1. **Фаза 1: Аналіз бізнес-процесів**
   * Термін: Тижні 1-2.
   * Очікуваний результат: Карта шляху користувача та концепт-архітектура.
2. **Фаза 2: Розробка та Конфігурація**
   * Термін: Тижні 3-5.
   * Очікуваний результат: Інтеграційний клієнт-сервер.
3. **Фаза 3: Тестування та Реліз**
   * Термін: Тиждень 6.
   * Очікуваний результат: Повна стабільність системи.

---

## 3. Розрахунок вартості послуг

Ми підготували гнучкі тарифні плани на основі потреб вашого бізнесу:

| Назва Етапу / Сервісу | Тривалість | Вартість | Статус дії |
| :--- | :---: | :---: | :---: |
| Консалтинг та проектування | 14 днів | $1,200 | Заплановано |
| Frontend-модулі та Налаштування | 21 день | $2,800 | Узгоджено |
| Тестування та Навчання персоналу | 7 днів | $800 | Безкоштовно |
| **Разом** | **42 дні** | **$4,800** | **Діє спецціна** |

> **Примітка:** Вказана вартість є фіксованою та покриває 12 місяців технічної підтримки та оновлень рішень.

---

## 4. Наступні кроки

Для старту спільного проекту необхідно:
1. Затвердити склад робочої групи.
2. Підписати договір про нерозголошення конфіденціальних даних (NDA).
3. Провести первинну координаційну зустріч.

Будемо раді співпраці з вами!
`
  },
  {
    id: 'meeting_minutes',
    name: 'Протокол робочої зустрічі (Meeting Minutes)',
    emoji: '📝',
    category: 'Команда',
    title: 'ПРОТОКОЛ НАРАДИ КОМАНДИ',
    subtitle: 'Планування релізу продукту та узгодження дорожньої карти',
    content: `# Протокол робочої зустрічі: Реліз v2.4

**Дата проведення**: 24 травня 2026 року  
**Час**: 14:00 - 15:30 (EEST)  
**Учасники**: Андрійченко П., Мирошниченко С., Коваленко О.

---

## 1. Порядок денний

* Аналіз відгуків тестувальників про стабільність завантаження файлів.
* Огляд вимог до експорту таблиць у Microsoft Word.
* Затвердження фінальних термінів деплою на GitHub Pages.

---

## 2. Ухвалені рішення

Під час обговорення було погоджено наступні ключові аспекти:

* **Альтернатива серверу**: Конвертація повинна відбуватися виключно на стороні клієнта за допомогою \`docx.js\` та \`marked.js\`. Це гарантує нульовий бюджет на підтримку бекенду та робить деплой на GitHub Pages миттєвим.
* **Стилістичний профіль**: Додати вибір кольорових акцентів та шрифтових пар прямо в інтерфейс, щоб користувачі мали кастомні варіації Word.
* **Адаптивний екран**: Зробити інтерфейс розділеним (split-screen) для комфортної одночасної роботи з кодом та перегляду результатів.

---

## 3. Таблиця завдань та відповідальності

| Завдання | Відповідальний | Термін | Пріорітет |
| :--- | :--- | :---: | :---: |
| Написати конвеєр переробки markdown токенів у Docx | Мирошниченко С. | 25 травня | **Високий** |
| Створити адаптивний UI-дизайн та тему Slate | Коваленко О. | 26 травня | Середній |
| Провести фінальне тестування експорту таблиць | Андрійченко П. | 28 травня | Низький |

---

## 4. Зауваження з тестування білдів

> Наполегливо рекомендується уникати завантаження важких картинок типу банерів 4K всередину DOCX, оскільки це збільшує розмір готового Word файлу та може знизити швидкість генерації в браузері. Обмеження картинкам в 2-4 МБ — оптимальний стандарт.

Зустріч оголошено закритою. Наступна координація відбудеться у середу.
`
  },
  {
    id: 'project_readme',
    name: 'README проєкту (Software Documentation)',
    emoji: '💻',
    category: 'Розробка',
    title: 'TECHNICAL DOCUMENTATION',
    subtitle: 'Markdown To Word Engine Guide & Specification',
    content: `# Developer README: Markdown to DOCX Engine

Welcome to the client-side conversion codebase repository. This project compiles and bundles markdown tokens directly into standard Microsoft Word Office Open XML formats.

---

## Technical Features

* **Zero Backend Overhead**: Compiles directly in the sandboxed browser execution cycle.
* **Modular Parsing Model**: Maps standard Markdown specifications to \`docx-js\` native blocks.
* **Theme Styling Profiles**: Seamlessly merges margins, cover sheets, and headers/footers.

---

## Quick Start Configuration

To run and debug the environment locally:

\`\`\`bash
# 1. Clone the static codebase
git clone https://github.com/user/markdown-to-docx.git

# 2. Install modern workspace dependencies
npm install

# 3. Launch Vite on secure dev port 3000
npm run dev
\`\`\`

---

## Project Structure Overview

Below is the directory architecture for static compilation:

| Directory Path | Purpose | Extensible |
| :--- | :--- | :---: |
| \`/src/utils/markdownToDocx.ts\` | Main translation logic between marked token tree & docx runs | **Yes** |
| \`/src/components/*\` | Modular layout widgets, forms, drag-drop cards | Yes |
| \`/src/App.tsx\` | Orchestration context, configuration panels & template loaders | No |

---

> This utility serves as a robust offline-first productivity tool. All files are manipulated completely on your computer in local memory, guaranteeing total privacy and immediate speed.
`
  }
];
