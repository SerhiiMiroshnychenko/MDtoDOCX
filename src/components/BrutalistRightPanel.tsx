import React from 'react';
import { Check } from 'lucide-react';
import { DOCUMENT_PRESETS, MarkdownPreset } from '../utils/presets';

interface BrutalistRightPanelProps {
  previewMode: 'paper' | 'fluid';
  setPreviewMode: (mode: 'paper' | 'fluid') => void;
  titlePage: boolean;
  headerText: string;
  titlePageTitle: string;
  titlePageSubtitle: string;
  titlePageAuthor: string;
  titlePageOrg: string;
  titlePageDate: string;
  includeToc: boolean;
  pageNumbers: boolean;
  footerText: string;
  currentThemePreset: {
    name: string;
    primaryColor: string;
    accentColor: string;
    fontHeading: string;
    fontBody: string;
  };
  fontSizeBody: number;
  fontSizeH1: number;
  fontSizeH2: number;
  fontSizeH3: number;
  spacingLineHeight: number;
  renderedPreviewHtml: string;
  fileName: string;
  setFileName: (n: string) => void;
  handleExportDocx: () => void;
  isExporting: boolean;
  markdown: string;
  logs: string[];
  activePresetId: string;
  handlePresetSelect: (preset: MarkdownPreset) => void;
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  handleFileUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
}

export default function BrutalistRightPanel({
  previewMode,
  setPreviewMode,
  titlePage,
  headerText,
  titlePageTitle,
  titlePageSubtitle,
  titlePageAuthor,
  titlePageOrg,
  titlePageDate,
  includeToc,
  pageNumbers,
  footerText,
  currentThemePreset,
  fontSizeBody,
  fontSizeH1,
  fontSizeH2,
  fontSizeH3,
  spacingLineHeight,
  renderedPreviewHtml,
  fileName,
  setFileName,
  handleExportDocx,
  isExporting,
  markdown,
  logs,
  activePresetId,
  handlePresetSelect,
  fileInputRef,
  handleFileUpload
}: BrutalistRightPanelProps) {
  return (
    <aside className="w-full lg:w-[465px] bg-[#E4E3E0] flex flex-col overflow-hidden shrink-0 border-t-4 lg:border-t-0 border-[#141414] min-h-[600px] lg:min-h-0 select-none">
      
      {/* File Queue / Templates Selector */}
      <div className="p-3 border-b-2 border-[#141414] bg-[#141414] text-[#E4E3E0] text-[10px] font-black uppercase tracking-widest flex items-center justify-between shrink-0">
        <span>Черга Файлів / Шаблони.md</span>
        <span className="bg-[#E4E3E0] text-[#141414] px-1.5 py-0.2 text-[8px] font-bold">READY</span>
      </div>

      <div className="p-3.5 bg-[#D8D7D3] border-b-2 border-[#141414] flex flex-col gap-2 shrink-0">
        <input
          type="file"
          id="right-md-uploader"
          className="hidden"
          ref={fileInputRef}
          onChange={handleFileUpload}
          accept=".md"
        />
        <button
          onClick={() => fileInputRef.current?.click()}
          className="w-full py-1.5 bg-white border-2 border-[#141414] text-[9px] font-black uppercase text-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors cursor-pointer"
        >
          + Своя Markdown-розмітка з ПК
        </button>
      </div>

      <div className="p-2 space-y-2 overflow-y-auto max-h-[140px] border-b-2 border-[#141414] bg-[#D8D7D3]/40 shrink-0 select-none">
        {DOCUMENT_PRESETS.map((preset) => {
          const isActive = activePresetId === preset.id;
          return (
            <div
              key={preset.id}
              onClick={() => handlePresetSelect(preset)}
              className={`p-2.5 bg-white border-2 border-[#141414] transition-all cursor-pointer flex flex-col ${
                isActive
                  ? 'shadow-[3px_3px_0px_0px_#141414]'
                  : 'opacity-65 grayscale-[30%] hover:opacity-100 hover:grayscale-0'
              }`}
            >
              <div className="flex justify-between items-start">
                <span className="text-[11px] font-black truncate pr-2 uppercase text-[#141414]">
                  {preset.emoji} {preset.name.split(' (')[0]}
                </span>
                <span className="text-[8px] font-mono opacity-60 font-bold uppercase">
                  {Math.round(preset.content.length / 100) / 10}kb
                </span>
              </div>
              <div className="mt-1.5 w-full bg-[#E4E3E0] h-1 border border-[#141414]/30">
                <div 
                  className={`h-full ${isActive ? 'bg-green-600' : 'bg-[#141414]'}`} 
                  style={{ width: isActive ? '100%' : '35%' }}
                />
              </div>
              <div className="mt-1 flex justify-between items-center text-[8px] font-black uppercase">
                <span className={isActive ? 'text-green-700 font-bold' : 'text-neutral-500'}>
                  {isActive ? 'Завантажено в редактор (LIVE)' : 'Готовий для редагування'}
                </span>
                {isActive && <Check className="h-2.5 w-2.5 text-green-700 inline ml-1" />}
              </div>
            </div>
          );
        })}
      </div>

      {/* Live Preview interactive header */}
      <div className="p-3 bg-[#D8D7D3] border-b-2 border-[#141414] flex items-center justify-between shrink-0">
        <span className="text-[10px] uppercase font-black text-[#141414]">Попередній Перегляд Word</span>
        <div className="flex gap-1 border-2 border-[#141414] bg-white p-0.5">
          <button
            onClick={() => setPreviewMode('paper')}
            className={`px-2.5 py-0.5 text-[8px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${
              previewMode === 'paper'
                ? 'bg-[#141414] text-[#E4E3E0]'
                : 'text-neutral-600 hover:bg-[#D8D7D3]/20'
            }`}
          >
            📄 Друк / Лист
          </button>
          <button
            onClick={() => setPreviewMode('fluid')}
            className={`px-2.5 py-0.5 text-[8px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${
              previewMode === 'fluid'
                ? 'bg-[#141414] text-[#E4E3E0]'
                : 'text-neutral-600 hover:bg-[#D8D7D3]/20'
            }`}
          >
            📱 Адаптив
          </button>
        </div>
      </div>

      {/* Preview Workspace */}
      <div className="flex-1 overflow-y-auto p-4 bg-[#D8D7D3]/20 border-b-2 border-[#141414] min-h-[220px]">
        {previewMode === 'paper' ? (
          <div className="space-y-4 max-w-full select-text mx-auto font-sans leading-normal">
            
            {titlePage && (
              <div className="bg-white border-2 border-[#141414] p-5.5 min-h-[340px] flex flex-col justify-between relative shadow-[3px_3px_0px_0px_#141414] overflow-hidden">
                <div className="text-[8px] font-mono text-neutral-500 text-right uppercase border-b border-neutral-200 pb-1 font-bold">
                  {headerText || 'MD TO DOCX DOCUMENTATION'}
                </div>
                <div className="my-auto space-y-2">
                  <div className="w-8 h-1 bg-[#141414]" />
                  <h1 className="text-base font-black leading-tight text-[#141414] uppercase tracking-tight font-serif">
                    {titlePageTitle || 'Новий документ'}
                  </h1>
                  {titlePageSubtitle && (
                    <p className="text-[10px] text-neutral-500 italic font-medium leading-relaxed max-w-xs">{titlePageSubtitle}</p>
                  )}
                </div>
                <div className="border-t border-neutral-200 pt-2 space-y-0.5 text-[#141414]">
                  {titlePageAuthor && <div className="text-[9px]"><span className="font-mono text-[7px] font-black opacity-50">AUTHOR:</span> {titlePageAuthor}</div>}
                  {titlePageOrg && <div className="text-[9px]"><span className="font-mono text-[7px] font-black opacity-50">ORG:</span> {titlePageOrg}</div>}
                  {titlePageDate && <div className="text-[9px] font-mono text-neutral-400">{titlePageDate}</div>}
                </div>
              </div>
            )}

            {includeToc && (
              <div className="bg-white border-2 border-[#141414] p-5.5 min-h-[160px] relative shadow-[3px_3px_0px_0px_#141414]">
                <h2 className="text-xs font-black border-b-2 border-[#141414] pb-1.5 mb-2.5 text-[#141414] uppercase">
                  Зміст (Table of Contents)
                </h2>
                <div className="space-y-1 select-none font-mono text-[9px]">
                  <div className="flex justify-between items-center text-[#141414] font-bold">
                    <span>1. ВСТУП ТА ПОСТАНОВКА ЗАДАЧІ</span>
                    <span className="border-b border-dotted border-black flex-1 mx-2"></span>
                    <span>ст. 2</span>
                  </div>
                  <div className="flex justify-between items-center text-neutral-600 pl-3">
                    <span>1.1 СТРУКТУРНА ДІАГРАМА ФАЗ</span>
                    <span className="border-b border-dotted border-neutral-300 flex-1 mx-2"></span>
                    <span>ст. 2</span>
                  </div>
                  <div className="flex justify-between items-center text-[#141414] font-bold">
                    <span>2. БЮДЖЕТ ТА РОЗРАХУНКИ</span>
                    <span className="border-b border-dotted border-black flex-1 mx-2"></span>
                    <span>ст. 3</span>
                  </div>
                </div>
              </div>
            )}

            <div className="bg-white border-2 border-[#141414] p-5.5 min-h-[340px] relative shadow-[3px_3px_0px_0px_#141414] break-words">
              <div className="text-[8px] font-mono text-neutral-400 font-bold border-b border-neutral-200 pb-1 mb-3.5 flex justify-between select-none uppercase">
                <span>{headerText ? 'КОЛОНТИТУЛ' : ''}</span>
                <span>{headerText || ''}</span>
              </div>

              <div
                style={{
                  '--primary': `#${currentThemePreset.primaryColor}`,
                  '--accent': `#${currentThemePreset.accentColor}`,
                  '--font-heading': currentThemePreset.fontHeading,
                  '--font-body': currentThemePreset.fontBody,
                  '--font-size-body': `${fontSizeBody}px`,
                  '--font-size-h1': `${fontSizeH1}px`,
                  '--font-size-h2': `${fontSizeH2}px`,
                  '--font-size-h3': `${fontSizeH3}px`,
                  '--line-height': spacingLineHeight
                } as React.CSSProperties}
                className="markdown-preview select-text text-slate-900"
                dangerouslySetInnerHTML={{ __html: renderedPreviewHtml }}
              />

              <div className="text-[8px] font-mono text-neutral-400 font-bold border-t border-neutral-200 pt-1.5 mt-5 flex justify-between select-none uppercase">
                <span>{footerText || ''}</span>
                {pageNumbers && <span>СТОРІНКА 3</span>}
              </div>
            </div>
          </div>
        ) : (
          <div
            style={{
              '--primary': `#${currentThemePreset.primaryColor}`,
              '--accent': `#${currentThemePreset.accentColor}`,
              '--font-heading': currentThemePreset.fontHeading,
              '--font-body': currentThemePreset.fontBody,
              '--font-size-body': `${fontSizeBody}px`,
              '--font-size-h1': `${fontSizeH1}px`,
              '--font-size-h2': `${fontSizeH2}px`,
              '--font-size-h3': `${fontSizeH3}px`,
              '--line-height': spacingLineHeight
            } as React.CSSProperties}
            className="markdown-preview bg-white p-5.5 border-2 border-[#141414] shadow-[3px_3px_0px_0px_#141414] max-w-full select-text leading-relaxed text-slate-900 break-words"
            dangerouslySetInnerHTML={{ __html: renderedPreviewHtml }}
          />
        )}
      </div>

      {/* Downloader File Name Form */}
      <div className="p-3 bg-white border-b-2 border-[#141414] flex items-center justify-between gap-2 shrink-0">
        <span className="text-[9px] font-mono font-black uppercase text-[#141414]">Ім'я файлу DOCX:</span>
        <input
          type="text"
          placeholder="документ.docx"
          className="flex-1 bg-[#E4E3E0] text-xs font-mono font-bold px-2 py-1 border border-[#141414] text-[#141414] outline-none"
          value={fileName}
          onChange={(e) => setFileName(e.target.value)}
        />
      </div>

      {/* Action Download / Export Button */}
      <div className="p-3.5 bg-white border-b-2 border-[#141414] shrink-0">
        <button
          onClick={handleExportDocx}
          disabled={isExporting || !markdown.trim()}
          className="w-full py-2.5 px-4 font-black text-[11px] uppercase tracking-wider text-[#E4E3E0] select-none transition-all cursor-pointer border-2 border-[#141414] shadow-[3px_3px_0px_0px_#141414] hover:shadow-none bg-[#141414] hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {isExporting ? 'компліюємо та будуємо...' : 'Конвертувати у .DOCX'}
        </button>
      </div>

      {/* Local System Console Logs Feed */}
      <div className="p-3 bg-white shrink-0">
        <h4 className="text-[9px] font-black uppercase mb-1 text-[#141414] tracking-wider select-none">Лог консолі локального клієнта</h4>
        <div className="font-mono text-[8px] h-14 overflow-y-auto space-y-0.5 opacity-80 text-[#141414] select-none">
          {logs.map((log, i) => (
            <div key={i} className="leading-tight truncate">{log}</div>
          ))}
        </div>
      </div>
    </aside>
  );
}
