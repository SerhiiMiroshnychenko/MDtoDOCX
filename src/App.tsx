import React, { useState, useRef } from 'react';
import { Download, Upload, FileText, Loader } from 'lucide-react';
import { AnimatePresence, motion } from 'motion/react';
import {
  convertMarkdownToDocx,
  DocxStyleConfig,
  THEME_PRESETS
} from './utils/markdownToDocx';
import BrutalistSidebar from './components/BrutalistSidebar';

export default function App() {
  const [markdown, setMarkdown] = useState<string>('');
  const [fileName, setFileName] = useState<string>('document.docx');
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [notice, setNotice] = useState<{ text: string; type: 'success' | 'info' | 'error' } | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const [styleTab, setStyleTab] = useState<'theme' | 'layout' | 'cover'>('theme');
  const [styleTheme, setStyleTheme] = useState<string>('slate');
  const [fontSizeBody, setFontSizeBody] = useState<number>(11);
  const [fontSizeH1, setFontSizeH1] = useState<number>(24);
  const [fontSizeH2, setFontSizeH2] = useState<number>(18);
  const [fontSizeH3, setFontSizeH3] = useState<number>(14);
  const [spacingLineHeight, setSpacingLineHeight] = useState<number>(1.25);
  const [marginSize, setMarginSize] = useState<'normal' | 'narrow' | 'wide'>('normal');
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [includeToc, setIncludeToc] = useState<boolean>(true);
  const [pageNumbers, setPageNumbers] = useState<boolean>(true);
  const [titlePage, setTitlePage] = useState<boolean>(true);
  const [headerText, setHeaderText] = useState<string>('');
  const [footerText, setFooterText] = useState<string>('');
  const [titlePageTitle, setTitlePageTitle] = useState<string>('');
  const [titlePageSubtitle, setTitlePageSubtitle] = useState<string>('');
  const [titlePageAuthor, setTitlePageAuthor] = useState<string>('');
  const [titlePageOrg, setTitlePageOrg] = useState<string>('');
  const [titlePageDate, setTitlePageDate] = useState<string>(new Date().toLocaleDateString('uk-UA'));

  const addLog = () => {};

  const showNotice = (text: string, type: 'success' | 'info' | 'error' = 'info') => {
    setNotice({ text, type });
    setTimeout(() => setNotice(null), 4000);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setProgress(15);
    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result;
      if (typeof content === 'string') {
        setMarkdown(content);
        const h1Match = content.match(/^#\s+(.*)$/m);
        setTitlePageTitle(h1Match ? h1Match[1].trim() : file.name.replace(/\.[^/.]+$/, ''));
        setFileName(file.name.replace(/\.[^/.]+$/, '.docx'));
        setProgress(100);
        setTimeout(() => setProgress(0), 600);
        showNotice(`Файл "${file.name}" завантажено`, 'success');
      }
    };
    reader.readAsText(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && file.name.endsWith('.md')) {
      setProgress(15);
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result;
        if (typeof content === 'string') {
          setMarkdown(content);
          const h1Match = content.match(/^#\s+(.*)$/m);
          setTitlePageTitle(h1Match ? h1Match[1].trim() : file.name.replace(/\.[^/.]+$/, ''));
          setFileName(file.name.replace(/\.[^/.]+$/, '.docx'));
          setProgress(100);
          setTimeout(() => setProgress(0), 600);
          showNotice(`Файл "${file.name}" завантажено`, 'success');
        }
      };
      reader.readAsText(file);
    } else {
      showNotice('Перетягніть тільки .md файл', 'error');
    }
  };

  const handleExportDocx = async () => {
    if (!markdown.trim()) {
      showNotice('Завантажте .md файл спочатку', 'error');
      return;
    }

    setIsExporting(true);
    setProgress(10);

    try {
      const activeThemePreset = THEME_PRESETS[styleTheme] || THEME_PRESETS.slate;

      const config: DocxStyleConfig = {
        fontBody: activeThemePreset.fontBody,
        fontHeading: activeThemePreset.fontHeading,
        primaryColor: activeThemePreset.primaryColor,
        accentColor: activeThemePreset.accentColor,
        fontSizeBody, fontSizeH1, fontSizeH2, fontSizeH3,
        spacingLineHeight, marginSize, orientation,
        includeToc, pageNumbers, headerText, footerText,
        titlePage, titlePageTitle, titlePageSubtitle,
        titlePageAuthor, titlePageOrg, titlePageDate,
      };

      setProgress(40);
      const blob = await convertMarkdownToDocx(markdown, config);
      setProgress(80);

      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      let safeName = fileName.trim();
      if (!safeName.endsWith('.docx')) safeName += '.docx';
      link.download = safeName;
      document.body.appendChild(link);
      link.click();
      setTimeout(() => {
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
      }, 100);

      setProgress(100);
      setTimeout(() => setProgress(0), 800);
      showNotice(`Файл "${safeName}" готовий`, 'success');
    } catch (err: any) {
      showNotice(`Помилка: ${err?.message || err}`, 'error');
    } finally {
      setIsExporting(false);
    }
  };

  const currentThemePreset = THEME_PRESETS[styleTheme] || THEME_PRESETS.slate;

  return (
    <div className="min-h-screen lg:h-screen w-full bg-[#E4E3E0] text-[#141414] font-sans overflow-x-hidden lg:overflow-hidden flex flex-col border-4 lg:border-8 border-[#141414] select-none">
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

      <header className="flex items-center justify-between px-6 py-4 border-b-4 border-[#141414] bg-white gap-4 shrink-0">
        <div className="flex items-center space-x-3">
          <div className="w-8 h-8 bg-[#141414] flex items-center justify-center shrink-0">
            <span className="text-[#E4E3E0] font-black text-xs">MD</span>
          </div>
          <h1 className="text-xl font-black uppercase tracking-tighter italic font-serif text-[#141414]">
            Convert.io // MD to DOCX Converter
          </h1>
        </div>
        <div className="flex items-center gap-2.5">
          <div className="flex items-center space-x-1.5 text-[9px] uppercase font-mono font-bold text-[#141414] bg-[#E4E3E0] border border-[#141414] px-2.5 py-1">
            <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse font-bold" />
            <span>STATUS: ACTIVE & SECURE CLIENT</span>
          </div>
        </div>
      </header>

      <main className="flex-1 flex flex-col lg:flex-row overflow-hidden">
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

        <section className="flex-1 bg-white flex flex-col items-center justify-center p-10 border-b-4 lg:border-b-0 lg:border-r-4 border-[#141414]">
          {!markdown ? (
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`w-full max-w-lg flex flex-col items-center justify-center p-16 border-4 border-dashed transition-colors ${
                isDragging ? 'border-[#141414] bg-neutral-100' : 'border-neutral-300 bg-neutral-50'
              }`}
            >
              <Upload className={`w-16 h-16 mb-6 ${isDragging ? 'text-[#141414]' : 'text-neutral-300'}`} />
              <h2 className="text-2xl font-black uppercase tracking-tight mb-2 text-center">
                {isDragging ? 'Відпустіть файл' : 'Завантажте .md файл'}
              </h2>
              <p className="text-sm text-neutral-500 mb-8 text-center">
                або натисніть кнопку, щоб вибрати файл
              </p>
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".md"
                hidden
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-8 py-3 bg-[#141414] text-white font-black uppercase text-sm tracking-wider hover:bg-neutral-800 transition-colors cursor-pointer"
              >
                Вибрати файл
              </button>
            </div>
          ) : (
            <div className="w-full max-w-md space-y-6">
              {progress > 0 && (
                <div className="w-full bg-neutral-200 h-2 border border-[#141414]">
                  <div
                    className="h-full bg-[#141414] transition-all duration-500 ease-out"
                    style={{ width: `${progress}%` }}
                  />
                </div>
              )}

              <div className="flex items-center gap-3 p-4 border-2 border-[#141414] bg-neutral-50">
                <FileText className="w-8 h-8 text-[#141414] shrink-0" />
                <div className="min-w-0">
                  <p className="font-bold text-sm truncate">{fileName}</p>
                  <p className="text-xs text-neutral-500">
                    {markdown.length} символів, {markdown.split('\n').length} рядків
                  </p>
                </div>
                <button
                  onClick={() => { setMarkdown(''); setFileName('document.docx'); }}
                  className="ml-auto shrink-0 text-[10px] font-bold uppercase text-neutral-400 hover:text-red-600 transition-colors cursor-pointer"
                >
                  Замінити
                </button>
              </div>

              <div>
                <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1">
                  Ім'я файлу:
                </label>
                <input
                  type="text"
                  value={fileName}
                  onChange={(e) => setFileName(e.target.value)}
                  className="w-full px-4 py-2 border-2 border-[#141414] font-mono text-sm outline-none bg-white"
                  placeholder="document.docx"
                />
              </div>

              <button
                onClick={handleExportDocx}
                disabled={isExporting}
                className="w-full py-3 bg-[#141414] text-white font-black uppercase text-sm tracking-wider hover:bg-neutral-800 transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2 cursor-pointer"
              >
                {isExporting ? (
                  <>
                    <Loader className="w-4 h-4 animate-spin" />
                    Конвертація...
                  </>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    Завантажити .docx
                  </>
                )}
              </button>
            </div>
          )}
        </section>
      </main>

      <footer className="h-8 bg-[#141414] text-[#E4E3E0] flex items-center px-6 justify-between text-[10px] uppercase font-bold tracking-tight shrink-0 select-none">
        <span>MDtoDOCX — конвертер Markdown у Word</span>
        <span>Open source tool</span>
      </footer>
    </div>
  );
}
