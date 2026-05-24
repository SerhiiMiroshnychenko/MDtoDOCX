import React, { useState, useRef, useEffect } from 'react';
import {
  FileText,
  Download,
  Plus,
  RefreshCw,
  Eye,
  Settings,
  HelpCircle,
  CheckCircle2,
  Trash2,
  Bold,
  Italic,
  Heading1,
  Heading2,
  List,
  ListOrdered,
  Quote,
  Table as TableIcon,
  Code2,
  Link,
  ChevronRight,
  Info,
  Layers,
  Sparkles,
  BookOpen,
  Check,
  Search,
  BookMarked
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { marked } from 'marked';
import {
  convertMarkdownToDocx,
  DocxStyleConfig,
  THEME_PRESETS
} from './utils/markdownToDocx';
import { DOCUMENT_PRESETS, MarkdownPreset } from './utils/presets';
import BrutalistSidebar from './components/BrutalistSidebar';
import BrutalistRightPanel from './components/BrutalistRightPanel';

export default function App() {
  // Main Markdown markdown text state
  const [markdown, setMarkdown] = useState<string>(DOCUMENT_PRESETS[0].content);
  
  // Selection / active preset to keep track
  const [activePresetId, setActivePresetId] = useState<string>(DOCUMENT_PRESETS[0].id);

  // Active workspace panels
  const [editorTab, setEditorTab] = useState<'edit' | 'mods'>('edit');
  const [styleTab, setStyleTab] = useState<'theme' | 'layout' | 'cover'>('theme');
  const [previewMode, setPreviewMode] = useState<'paper' | 'fluid'>('paper');

  // Input Ref
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Customizer styling options state
  const [styleTheme, setStyleTheme] = useState<string>('slate');
  const [fontSizeBody, setFontSizeBody] = useState<number>(11);
  const [fontSizeH1, setFontSizeH1] = useState<number>(24);
  const [fontSizeH2, setFontSizeH2] = useState<number>(18);
  const [fontSizeH3, setFontSizeH3] = useState<number>(14);
  const [spacingLineHeight, setSpacingLineHeight] = useState<number>(1.25);
  const [marginSize, setMarginSize] = useState<'normal' | 'narrow' | 'wide'>('normal');
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait');
  
  // Checkboxes
  const [includeToc, setIncludeToc] = useState<boolean>(true);
  const [pageNumbers, setPageNumbers] = useState<boolean>(true);
  const [titlePage, setTitlePage] = useState<boolean>(true);

  // Headers and Footers
  const [headerText, setHeaderText] = useState<string>('Звіт з документообігу');
  const [footerText, setFooterText] = useState<string>('Конфіденційно');

  // Cover Page Fields
  const [titlePageTitle, setTitlePageTitle] = useState<string>(DOCUMENT_PRESETS[0].title);
  const [titlePageSubtitle, setTitlePageSubtitle] = useState<string>(DOCUMENT_PRESETS[0].subtitle);
  const [titlePageAuthor, setTitlePageAuthor] = useState<string>('Сергій Мирошниченко');
  const [titlePageOrg, setTitlePageOrg] = useState<string>('Інноваційні Технології ТОВ');
  const [titlePageDate, setTitlePageDate] = useState<string>(new Date().toLocaleDateString('uk-UA'));

  // Console log feed for neoclassical design draft
  const [logs, setLogs] = useState<string[]>([]);
  const addLog = (msg: string) => {
    const time = new Date().toLocaleTimeString('uk-UA', { hour12: false });
    setLogs((prev) => [`[${time}] ${msg}`, ...prev].slice(0, 30));
  };

  // Pre-seed logs on mount
  useEffect(() => {
    const time = new Date().toLocaleTimeString('uk-UA', { hour12: false });
    setLogs([
      `[${time}] Очікування нових файлів...`,
      `[${time}] Усі модулі готові. Конвертер 100% клієнтський.`,
      `[${time}] Завантажено розширення Word стилів (DOCX v9).`,
      `[${time}] Ініціалізація WASM середовища... успішно`
    ]);
  }, []);

  // Export properties
  const [fileName, setFileName] = useState<string>('документ.docx');
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [notice, setNotice] = useState<{ text: string; type: 'success' | 'info' | 'error' } | null>(null);

  // Modifiers state
  const [findText, setFindText] = useState<string>('');
  const [replaceText, setReplaceText] = useState<string>('');

  // Auto-sync cover sheet titles when they are selected
  const handlePresetSelect = (preset: MarkdownPreset) => {
    setMarkdown(preset.content);
    setActivePresetId(preset.id);
    setTitlePageTitle(preset.title);
    setTitlePageSubtitle(preset.subtitle);
    // Custom set filename
    const sanitizedTitle = preset.name.split(' (')[0].toLowerCase().replace(/\s+/g, '_');
    setFileName(`${sanitizedTitle}.docx`);
    showNotice(`Завантажено шаблон: ${preset.name}`, 'success');
    addLog(`Завантажено файл шаблону: "${preset.name}"`);
  };

  // Helper for displaying snackbars / notices
  const showNotice = (text: string, type: 'success' | 'info' | 'error' = 'info') => {
    setNotice({ text, type });
    setTimeout(() => {
      setNotice(null);
    }, 4000);
  };

  // File selection parser
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result;
      if (typeof content === 'string') {
        setMarkdown(content);
        setActivePresetId('');
        
        // Auto extract possible H1 as the cover page title!
        const h1Match = content.match(/^#\s+(.*)$/m);
        if (h1Match) {
          setTitlePageTitle(h1Match[1].trim());
        } else {
          setTitlePageTitle(file.name.replace(/\.[^/.]+$/, ""));
        }
        
        // Set filename based on uploaded file
        const newDocxName = file.name.replace(/\.[^/.]+$/, ".docx");
        setFileName(newDocxName);
        showNotice(`Файл "${file.name}" успішно імпортовано!`, 'success');
        addLog(`Імпортовано локальний файл: "${file.name}" (${(file.size / 1024).toFixed(1)} kb)`);
      }
    };
    reader.readAsText(file);
  };

  // Drag and Drop files handlers
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && file.name.endsWith('.md')) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result;
        if (typeof content === 'string') {
          setMarkdown(content);
          setActivePresetId('');
          // Auto guess name
          const h1Match = content.match(/^#\s+(.*)$/m);
          setTitlePageTitle(h1Match ? h1Match[1].trim() : file.name.replace(/\.[^/.]+$/, ""));
          setFileName(file.name.replace(/\.[^/.]+$/, ".docx"));
          showNotice(`Успішно імпортовано перетягнутий файл!`, 'success');
          addLog(`Перетягнуто та імпортовано файл: "${file.name}" (${(file.size / 1024).toFixed(1)} kb)`);
        }
      };
      reader.readAsText(file);
    } else {
      showNotice("Будь ласка, перетягуйте тільки файли з розширенням .md", "error");
      addLog(`Помилка імпортування: файл не є .md сумісним`);
    }
  };

  // Reset document to empty space
  const handleClearEditor = () => {
    if (window.confirm("Очистити поле редактора? Всі незбережені зміни буде втрачено.")) {
      setMarkdown('');
      setActivePresetId('');
      setTitlePageTitle('Новий документ');
      setTitlePageSubtitle('');
      showNotice("Редактор очищено", "info");
      addLog("Вміст редактору повністю очищено.");
    }
  };

  // Insertion helper for Markdown buttons formatting
  const insertMarkdownHelper = (prefix: string, suffix: string = '') => {
    if (!textareaRef.current) return;
    const textarea = textareaRef.current;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const selectedText = text.substring(start, end);
    const replacement = prefix + selectedText + suffix;
    const newVal = text.substring(0, start) + replacement + text.substring(end);
    setMarkdown(newVal);

    // Reset cursor focus and select highlight back to selection
    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(
        start + prefix.length,
        start + prefix.length + selectedText.length
      );
    }, 10);
  };

  // Modifier Actions Suite
  const handleFindReplace = () => {
    if (!findText) {
      showNotice("Будь ласка, вкажіть слово для пошуку.", "error");
      return;
    }
    const escapedSearch = findText.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const regex = new RegExp(escapedSearch, 'g');
    const occurrencesCount = (markdown.match(regex) || []).length;
    
    if (occurrencesCount === 0) {
      showNotice(`Слово "${findText}" не знайдено в документі.`, "info");
      addLog(`Пошук: слово "${findText}" не знайдено для заміни.`);
      return;
    }
    
    const nextMarkdown = markdown.replace(regex, replaceText);
    setMarkdown(nextMarkdown);
    showNotice(`Успішно замінено ${occurrencesCount} входжень!`, "success");
    addLog(`Заміна: замінено "${findText}" -> "${replaceText}" (${occurrencesCount} роз)`);
    setFindText('');
    setReplaceText('');
  };

  const handleBeautify = () => {
    let text = markdown;
    // 1. Cleans triple spacing newlines to double newlines
    text = text.replace(/\n{3,}/g, '\n\n');
    // 2. Spaces after lists hyphens
    text = text.replace(/^-\s*([^\s\-])/gm, '- $1');
    // 3. Spaces after headers hashes
    text = text.replace(/^(#{1,6})\s*([^\s#])/gm, '$1 $2');
    // 4. Clean trailing whitespace in rows
    text = text.replace(/[ \t]+$/gm, '');
    
    setMarkdown(text);
    showNotice("Документ автоматично відформатовано та виправлено структуру!", "success");
    addLog("Авто-вирівнювання: покращено форматування та виправлено структуру Markdown");
  };

  const handleInsertToc = () => {
    const lines = markdown.split('\n');
    const headings = lines
      .map(line => {
        const match = line.match(/^(#{1,6})\s*(.*)$/);
        if (match) {
          const depth = match[1].length;
          const text = match[2].trim();
          return { depth, text };
        }
        return null;
      })
      .filter((h): h is { depth: number; text: string } => h !== null);

    if (headings.length === 0) {
      showNotice("Не знайдено жодного Markdown заголовка (# або ##) для генерації Змісту.", "error");
      addLog("Помилка генерації Змісту: у тексті немає жодних Markdown заголовків (#, ##)");
      return;
    }

    let tocMd = '## Зміст документа\n\n';
    headings.forEach((h) => {
      const indent = '  '.repeat(Math.max(0, h.depth - 1));
      tocMd += `${indent}* [${h.text}](#${h.text.toLowerCase().replace(/[^a-z0-9а-яёієїґ'-]+/g, '-')})\n`;
    });
    tocMd += '\n---\n';

    setMarkdown(tocMd + '\n' + markdown);
    showNotice("Згенований Зміст додано на самий початок документа!", "success");
    addLog(`Генератор Змісту: успішно додано ${headings.length} розділів у TOC.`);
  };

  const handleCaseChange = (mode: 'upper' | 'lower' | 'title') => {
    if (!textareaRef.current) return;
    const textarea = textareaRef.current;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const selection = markdown.substring(start, end);

    if (!selection) {
      // If nothing selected, process the whole text block
      let nextAllText = markdown;
      if (mode === 'upper') nextAllText = markdown.toUpperCase();
      else if (mode === 'lower') nextAllText = markdown.toLowerCase();
      else if (mode === 'title') {
        nextAllText = markdown.replace(/\b\w/g, c => c.toUpperCase());
      }
      setMarkdown(nextAllText);
      showNotice("Весь документ переведено у вибраний регістр!", "success");
      addLog(`Перетворення регістру: увесь текст переведено у стиль "${mode}"`);
      return;
    }

    let nextSelection = selection;
    if (mode === 'upper') nextSelection = selection.toUpperCase();
    else if (mode === 'lower') nextSelection = selection.toLowerCase();
    else if (mode === 'title') {
      nextSelection = selection.replace(/\b\w/g, c => c.toUpperCase());
    }

    const newVal = markdown.substring(0, start) + nextSelection + markdown.substring(end);
    setMarkdown(newVal);
    showNotice("Вибраний фрагмент тексту перетворено!", "success");
    addLog(`Перетворення регістру: фрагмент довжиною ${selection.length} симв. переведено у стиль "${mode}"`);
    
    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(start, start + nextSelection.length);
    }, 10);
  };

  // Core DOCX Compiler trigger
  const handleExportDocx = async () => {
    if (!markdown.trim()) {
      showNotice("Будь ласка, заповніть документ вмістом перед експортом", "error");
      return;
    }

    setIsExporting(true);
    showNotice("Почалася компіляція Word документа...", "info");
    addLog("Компілятор: розпочато збір структури документа DOCX...");

    try {
      // Map active preset to configuration object
      const activeThemePreset = THEME_PRESETS[styleTheme] || THEME_PRESETS.slate;
      
      const config: DocxStyleConfig = {
        fontBody: activeThemePreset.fontBody,
        fontHeading: activeThemePreset.fontHeading,
        primaryColor: activeThemePreset.primaryColor,
        accentColor: activeThemePreset.accentColor,
        fontSizeBody,
        fontSizeH1,
        fontSizeH2,
        fontSizeH3,
        spacingLineHeight,
        marginSize,
        orientation,
        includeToc,
        pageNumbers,
        headerText: headerText,
        footerText: footerText,
        titlePage,
        titlePageTitle,
        titlePageSubtitle,
        titlePageAuthor,
        titlePageOrg,
        titlePageDate
      };

      addLog(`Стилізація: шрифт-заголовків: "${activeThemePreset.fontHeading}", шрифт-тіла: "${activeThemePreset.fontBody}", колір: #${activeThemePreset.primaryColor}`);

      // Trigger compiler
      const blob = await convertMarkdownToDocx(markdown, config);

      // Create download trigger browser-side
      const downloadUrl = window.URL.createObjectURL(blob);
      const tempLink = document.createElement('a');
      tempLink.href = downloadUrl;
      
      // Sanitized filename check
      let safeFilename = fileName.trim();
      if (!safeFilename.endsWith('.docx')) {
        safeFilename += '.docx';
      }
      tempLink.download = safeFilename;
      
      document.body.appendChild(tempLink);
      tempLink.click();
      
      // Cleanup after download runs
      setTimeout(() => {
        document.body.removeChild(tempLink);
        window.URL.revokeObjectURL(downloadUrl);
      }, 100);

      showNotice(`Файл успішно сконвертовано! Завантажено "${safeFilename}"!`, "success");
      addLog(`Успішно сконвертовано! Файл завантажено: "${safeFilename}"`);
    } catch (err: any) {
      console.error(err);
      showNotice(`Помилка генерації Word файлу: ${err?.message || err}`, "error");
      addLog(`Помилка генератора: ${err?.message || 'Невідомий збій рендерера'}`);
    } finally {
      setIsExporting(false);
    }
  };

  // Fetch HTML rendering from marked parser for display preview
  const [renderedPreviewHtml, setRenderedPreviewHtml] = useState<string>('');
  useEffect(() => {
    try {
      // Synchronous parse using marked
      const html = marked.parse(markdown || '*Документ пустий. Напишіть щось або завантажте шаблон.*');
      setRenderedPreviewHtml(html as string);
    } catch (e) {
      console.error(e);
      setRenderedPreviewHtml('<p className="text-red-555">Помилка рендеру розмітки</p>');
    }
  }, [markdown]);

  // Current theme config for styling preview container inputs
  const currentThemePreset = THEME_PRESETS[styleTheme] || THEME_PRESETS.slate;

  return (
    <div className="min-h-screen lg:h-screen w-full bg-[#E4E3E0] text-[#141414] font-sans overflow-x-hidden lg:overflow-hidden flex flex-col border-4 lg:border-8 border-[#141414] select-none">
      
      {/* Toast notifications */}
      <AnimatePresence>
        {notice && (
          <motion.div
            initial={{ opacity: 0, y: -50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="fixed top-4 left-1/2 -translate-x-1/2 z-50 flex items-center gap-3 px-5 py-3 border-2 border-[#141414] bg-white text-[#141414] font-mono text-xs uppercase font-bold shadow-[3px_3px_0px_0px_#141414] max-w-lg min-w-[320px]"
          >
            <span className="w-2 h-2 bg-green-500 rounded-full inline-block animate-ping shrink-0" />
            <p>{notice.text}</p>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Main Header */}
      <header className="flex flex-col sm:flex-row items-center justify-between px-6 py-4 border-b-4 border-[#141414] bg-white gap-4 shrink-0 selection:bg-neutral-200">
        <div className="flex items-center space-x-3">
          <div className="w-8 h-8 bg-[#141414] flex items-center justify-center shrink-0">
            <span className="text-[#E4E3E0] font-black text-xs">MD</span>
          </div>
          <div>
            <h1 className="text-xl font-black uppercase tracking-tighter italic font-serif text-[#141414]">
              Convert.io // MD to DOCX Converter
            </h1>
            <p className="text-[10px] uppercase font-bold tracking-wider opacity-60">
              Усі конвертації відбуваються безпечно та локально в браузері
            </p>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2.5">
          <div className="flex items-center space-x-1.5 text-[9px] uppercase font-mono font-bold text-[#141414] bg-[#E4E3E0] border border-[#141414] px-2.5 py-1">
            <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse font-bold"></div>
            <span>STATUS: ACTIVE & SECURE CLIENT</span>
          </div>
          <a
            href="https://pages.github.com"
            target="_blank"
            referrerPolicy="no-referrer"
            className="px-4 py-1.5 border-2 border-[#141414] text-[10px] font-black uppercase bg-white hover:bg-[#141414] hover:text-[#E4E3E0] text-[#141414] transition-all shadow-[2px_2px_0px_0px_#141414] hover:shadow-none"
          >
            Деплой на Pages ↗
          </a>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <main className="flex-1 flex flex-col lg:flex-row overflow-hidden selection:bg-[#141414] selection:text-[#E4E3E0]">
        
        {/* Modular Left Sidebar - DOCX options */}
        <BrutalistSidebar
          styleTab={styleTab}
          setStyleTab={setStyleTab}
          styleTheme={styleTheme}
          setStyleTheme={setStyleTheme}
          fontSizeBody={fontSizeBody}
          setFontSizeBody={setFontSizeBody}
          spacingLineHeight={spacingLineHeight}
          setSpacingLineHeight={setSpacingLineHeight}
          orientation={orientation}
          setOrientation={setOrientation}
          marginSize={marginSize}
          setMarginSize={setMarginSize}
          includeToc={includeToc}
          setIncludeToc={setIncludeToc}
          pageNumbers={pageNumbers}
          setPageNumbers={setPageNumbers}
          titlePage={titlePage}
          setTitlePage={setTitlePage}
          headerText={headerText}
          setHeaderText={setHeaderText}
          footerText={footerText}
          setFooterText={setFooterText}
          titlePageTitle={titlePageTitle}
          setTitlePageTitle={setTitlePageTitle}
          titlePageSubtitle={titlePageSubtitle}
          setTitlePageSubtitle={setTitlePageSubtitle}
          titlePageAuthor={titlePageAuthor}
          setTitlePageAuthor={setTitlePageAuthor}
          titlePageOrg={titlePageOrg}
          setTitlePageOrg={setTitlePageOrg}
          titlePageDate={titlePageDate}
          setTitlePageDate={setTitlePageDate}
          addLog={addLog}
          showNotice={showNotice}
        />

        {/* Center Panel (Editable Workspace & Enhancements Toolbar) */}
        <section className="flex-1 bg-white border-b-4 lg:border-b-0 lg:border-r-4 border-[#141414] flex flex-col overflow-hidden min-h-[500px] lg:min-h-0">
          
          {/* Section Sub-tabs selector for the center column */}
          <div className="flex border-b-2 border-[#141414] bg-[#D8D7D3]/40 select-none shrink-0">
            <button
              onClick={() => {
                setEditorTab('edit');
                addLog("Редактор Markdown: активний.");
              }}
              className={`px-5 py-2.5 border-r-2 border-[#141414] text-[10px] font-black uppercase transition-all tracking-wider cursor-pointer ${
                editorTab === 'edit'
                  ? 'bg-white text-[#141414]'
                  : 'text-neutral-600 hover:bg-white/40'
              }`}
            >
              📝 Редактор Markdown
            </button>
            <button
              onClick={() => {
                setEditorTab('mods');
                addLog("Інструменти-модифікатори: активні.");
              }}
              className={`px-5 py-2.5 border-r-2 border-[#141414] text-[10px] font-black uppercase transition-all tracking-wider cursor-pointer ${
                editorTab === 'mods'
                  ? 'bg-white text-[#141414]'
                  : 'text-neutral-600 hover:bg-white/40'
              }`}
            >
              🛠️ Модифікатори
            </button>
            
            <div className="flex-1" />
            
            {markdown.trim() && (
              <button
                onClick={handleClearEditor}
                className="px-4 py-2 text-[10px] font-extrabold uppercase text-neutral-500 hover:text-red-100 flex items-center gap-1 transition-all cursor-pointer"
              >
                <Trash2 className="h-3 w-3" />
                <span>Очистити</span>
              </button>
            )}
          </div>

          {/* Tab Pane Rendering */}
          {editorTab === 'edit' ? (
            <div className="flex-1 flex flex-col overflow-hidden">
              {/* Text formatting action icons in a strict, brutalist theme row */}
              <div className="flex items-center gap-1 p-2 bg-[#D8D7D3]/20 border-b-2 border-[#141414] overflow-x-auto select-none shrink-0 scrollbar-none text-[#141414]">
                <button
                  onClick={() => insertMarkdownHelper('**', '**')}
                  className="p-1 px-2 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Жирний тип"
                >
                  Bold
                </button>
                <button
                  onClick={() => insertMarkdownHelper('*', '*')}
                  className="p-1 px-2 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Курсивний тип"
                >
                  Ital
                </button>
                <div className="h-5 w-0.5 bg-[#141414] mx-1 shrink-0"></div>
                <button
                  onClick={() => insertMarkdownHelper('# ')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Заголовок H1"
                >
                  H1
                </button>
                <button
                  onClick={() => insertMarkdownHelper('## ')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Заголовок H2"
                >
                  H2
                </button>
                <div className="h-5 w-0.5 bg-[#141414] mx-1 shrink-0"></div>
                <button
                  onClick={() => insertMarkdownHelper('- ')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Маркований"
                >
                  List
                </button>
                <button
                  onClick={() => insertMarkdownHelper('1. ')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Нумерований"
                >
                  Num
                </button>
                <button
                  onClick={() => insertMarkdownHelper('> ')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Цитата"
                >
                  Quote
                </button>
                <button
                  onClick={() => insertMarkdownHelper('```\n', '\n```')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Блок коду"
                >
                  Code
                </button>
                <button
                  onClick={() => insertMarkdownHelper('| Стовпець | Значення |\n| :--- | :--- |\n| Текст | Опис |\n')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Створити таблицю"
                >
                  Table
                </button>
                <button
                  onClick={() => insertMarkdownHelper('[', '](https://example.com)')}
                  className="p-1 px-1.5 border-2 border-[#141414] text-[9px] font-black uppercase bg-white hover:bg-[#141414] hover:text-white transition-all cursor-pointer"
                  title="Додати лінк"
                >
                  Link
                </button>
              </div>

              {/* Central Textarea Form Fields */}
              <div className="flex-1 relative flex flex-col overflow-hidden bg-[#FAF9F5]">
                <textarea
                  ref={textareaRef}
                  id="primary-markdown-editor"
                  className="flex-1 w-full h-full p-5 text-sm font-mono text-[#141414] bg-white border-0 outline-none resize-none leading-relaxed placeholder:text-neutral-400 overflow-y-auto"
                  placeholder="Введіть або перетягніть Markdown контент сюди та перетворіть..."
                  value={markdown}
                  onChange={(e) => setMarkdown(e.target.value)}
                />
              </div>

              {/* Active document counts stats */}
              <div className="bg-[#D8D7D3] border-t-2 border-[#141414] px-4 py-2 flex items-center justify-between text-[9px] text-[#141414] font-mono select-none shrink-0 font-bold uppercase">
                <div>
                  Символів: <span>{markdown.length}</span> | Рядків: <span>{markdown.split('\n').length}</span>
                </div>
                <div>
                  Слів: <span>{markdown.trim() ? markdown.trim().split(/\s+/).length : 0}</span>
                </div>
              </div>
            </div>
          ) : (
            /* Layout of the Modifiers Suites */
            <div className="flex-1 overflow-y-auto p-5 space-y-5 bg-[#FAF9F5]">
              <div>
                <h3 className="text-xs font-black text-[#141414] uppercase tracking-wider mb-1">
                  Модифікатори структури
                </h3>
                <p className="text-[10px] text-neutral-600 uppercase font-bold leading-normal">
                  Обробіть та перетворіть ваш текст до початку експорту в документ.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pb-2">
                <div className="p-4 border-2 border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414] flex flex-col justify-between">
                  <div>
                    <h4 className="text-xs font-black uppercase text-[#141414] mb-1">Очищення структури</h4>
                    <p className="text-[9px] text-neutral-500 uppercase tracking-tight leading-relaxed mb-3">
                      Цей модифікатор вирівняє заголовки, списки та прибере зайві зазори.
                    </p>
                  </div>
                  <button
                    onClick={handleBeautify}
                    className="w-full text-[10px] font-black uppercase bg-white border-2 border-[#141414] py-2 hover:bg-[#141414] hover:text-white transition-colors cursor-pointer animate-none"
                  >
                    🪄 Очистити Markdown
                  </button>
                </div>

                <div className="p-4 border-2 border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414] flex flex-col justify-between">
                  <div>
                    <h4 className="text-xs font-black uppercase text-[#141414] mb-1">Генерація Змісту</h4>
                    <p className="text-[9px] text-neutral-500 uppercase tracking-tight leading-relaxed mb-3">
                      Автоматично конструює інтерактивний Зміст та додає в початок.
                    </p>
                  </div>
                  <button
                    onClick={handleInsertToc}
                    className="w-full text-[10px] font-black uppercase bg-white border-2 border-[#141414] py-2 hover:bg-[#141414] hover:text-white transition-colors cursor-pointer text-center"
                  >
                    📖 Сформувати Зміст
                  </button>
                </div>
              </div>

              {/* Shifty Find and Replace Block */}
              <div className="p-4 border-2 border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414] space-y-3">
                <h4 className="text-xs font-black uppercase text-[#141414]">Швидка заміна по тексту</h4>
                <div className="grid grid-cols-2 gap-3.5">
                  <div>
                    <label className="block text-[8px] font-mono font-bold uppercase text-neutral-500 mb-1">Шукати:</label>
                    <input
                      type="text"
                      className="w-full text-xs px-2.5 py-1.5 border-2 border-[#141414] bg-white font-mono outline-none"
                      placeholder="..."
                      value={findText}
                      onChange={(e) => setFindText(e.target.value)}
                    />
                  </div>
                  <div>
                    <label className="block text-[8px] font-mono font-bold uppercase text-neutral-500 mb-1">Замінити:</label>
                    <input
                      type="text"
                      className="w-full text-xs px-2.5 py-1.5 border-2 border-[#141414] bg-white font-mono outline-none"
                      placeholder="..."
                      value={replaceText}
                      onChange={(e) => setReplaceText(e.target.value)}
                    />
                  </div>
                </div>
                <button
                  onClick={handleFindReplace}
                  className="w-full py-1.5 text-[10px] font-black uppercase bg-[#141414] hover:bg-neutral-800 text-[#E4E3E0] transition-colors cursor-pointer"
                >
                  🔄 Виконати автозаміну
                </button>
              </div>

              {/* Register Letter Modifiers */}
              <div className="p-4 border-2 border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414] space-y-3">
                <h4 className="text-xs font-black uppercase text-[#141414]">Регістр Літер (Case Modification)</h4>
                <p className="text-[9px] text-neutral-500 uppercase tracking-tight mb-2">
                  Виділіть фрагмент тексту, щоб змінити його регістр. Або перетворіть увесь документ цілком.
                </p>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    onClick={() => handleCaseChange('upper')}
                    className="py-1.5 text-[9px] font-black uppercase border-2 border-[#141414] hover:bg-neutral-100 cursor-pointer"
                  >
                    АА (ВЕЛИКИЙ)
                  </button>
                  <button
                    onClick={() => handleCaseChange('lower')}
                    className="py-1.5 text-[9px] font-black uppercase border-2 border-[#141414] hover:bg-neutral-100 cursor-pointer"
                  >
                    аа (малий)
                  </button>
                  <button
                    onClick={() => handleCaseChange('title')}
                    className="py-1.5 text-[9px] font-black uppercase border-2 border-[#141414] hover:bg-neutral-100 cursor-pointer"
                  >
                    Аа (Титульний)
                  </button>
                </div>
              </div>

              {/* Smart Template Insertion Block */}
              <div className="p-4 border-2 border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414] space-y-3">
                <h4 className="text-xs font-black uppercase text-[#141414]">Генерація Швидких Блоків</h4>
                <div className="flex flex-wrap gap-2 pt-1 select-none">
                  <button
                    onClick={() => {
                      const stamp = `\n\n*Документ сформовано: ${new Date().toLocaleDateString('uk-UA')} о ${new Date().toLocaleTimeString('uk-UA')}*\n`;
                      setMarkdown(markdown + stamp);
                      showNotice("Додано часовий штамп клієнта", "info");
                      addLog("Вставка: додано часовий штамп в кінець.");
                    }}
                    className="px-3 py-1.5 bg-[#FAF9F5] hover:bg-neutral-100 border-2 border-[#141414] text-[9px] font-black uppercase cursor-pointer"
                  >
                    🕒 Часовий штамп
                  </button>
                  <button
                    onClick={() => {
                      const contacts = `\n\n---\n**Контакти:**\n* **Автор:** [Ваше Ім'я]\n* **Контактна пошта:** email@example.com\n* **Орган коорд.:** Керівництво\n`;
                      setMarkdown(markdown + contacts);
                      showNotice("Додано блок відповідальних контактів", "info");
                      addLog("Вставка: додано контакти в кінець.");
                    }}
                    className="px-3 py-1.5 bg-[#FAF9F5] hover:bg-neutral-100 border-2 border-[#141414] text-[9px] font-black uppercase cursor-pointer"
                  >
                    📞 Блок Контактів
                  </button>
                  <button
                    onClick={() => {
                      const table = `\n\n| Параметр | Пояснення / Специфікація | Статус |\n| :--- | :--- | :--- |\n| Задача | Сконвертувати Markdown | Завершено |\n| Деплой | Встановити GitHub Pages | Заплановано |\n`;
                      setMarkdown(markdown + table);
                      showNotice("Вставлено нову таблицю специфікації", "info");
                      addLog("Вставка: додано зразок специфікації у кінець.");
                    }}
                    className="px-3 py-1.5 bg-[#FAF9F5] hover:bg-neutral-100 border-2 border-[#141414] text-[9px] font-black uppercase cursor-pointer"
                  >
                    📊 Зразок Таблиці
                  </button>
                </div>
              </div>
            </div>
          )}
        </section>

        {/* Modular Right Sidebar Panel - Queue and Previews */}
        <BrutalistRightPanel
          previewMode={previewMode}
          setPreviewMode={setPreviewMode}
          titlePage={titlePage}
          headerText={headerText}
          titlePageTitle={titlePageTitle}
          titlePageSubtitle={titlePageSubtitle}
          titlePageAuthor={titlePageAuthor}
          titlePageOrg={titlePageOrg}
          titlePageDate={titlePageDate}
          includeToc={includeToc}
          pageNumbers={pageNumbers}
          footerText={footerText}
          currentThemePreset={currentThemePreset}
          fontSizeBody={fontSizeBody}
          fontSizeH1={fontSizeH1}
          fontSizeH2={fontSizeH2}
          fontSizeH3={fontSizeH3}
          spacingLineHeight={spacingLineHeight}
          renderedPreviewHtml={renderedPreviewHtml}
          fileName={fileName}
          setFileName={setFileName}
          handleExportDocx={handleExportDocx}
          isExporting={isExporting}
          markdown={markdown}
          logs={logs}
          activePresetId={activePresetId}
          handlePresetSelect={handlePresetSelect}
          fileInputRef={fileInputRef}
          handleFileUpload={handleFileUpload}
        />

      </main>

      {/* Retro Neoclassical Footer Row */}
      <footer className="h-8 bg-[#141414] text-[#E4E3E0] flex items-center px-6 justify-between text-[10px] uppercase font-bold tracking-tight shrink-0 select-none border-t-2 border-[#141414]">
        <div className="flex items-center space-x-4">
          <span>SYSTEM: OK</span>
          <span className="hidden sm:inline">ENGINE: LOCAL WASM WORD GENERATOR v2</span>
          <span className="hidden md:inline">HOSTING: GitHub Pages Friendly</span>
        </div>
        <div>Open source tool • © 2026</div>
      </footer>

    </div>
  );
}
