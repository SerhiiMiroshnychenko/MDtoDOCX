import React from 'react';
import { THEME_PRESETS } from '../utils/markdownToDocx';

interface BrutalistSidebarProps {
  styleTab: 'theme' | 'layout' | 'cover';
  setStyleTab: (tab: 'theme' | 'layout' | 'cover') => void;
  styleTheme: string;
  setStyleTheme: (theme: string) => void;
  fontSizeBody: number;
  setFontSizeBody: (size: number) => void;
  spacingLineHeight: number;
  setSpacingLineHeight: (spacing: number) => void;
  orientation: 'portrait' | 'landscape';
  setOrientation: (o: 'portrait' | 'landscape') => void;
  marginSize: 'normal' | 'narrow' | 'wide';
  setMarginSize: (m: 'normal' | 'narrow' | 'wide') => void;
  includeToc: boolean;
  setIncludeToc: (toc: boolean) => void;
  pageNumbers: boolean;
  setPageNumbers: (p: boolean) => void;
  titlePage: boolean;
  setTitlePage: (title: boolean) => void;
  headerText: string;
  setHeaderText: (text: string) => void;
  footerText: string;
  setFooterText: (text: string) => void;
  titlePageTitle: string;
  setTitlePageTitle: (t: string) => void;
  titlePageSubtitle: string;
  setTitlePageSubtitle: (s: string) => void;
  titlePageAuthor: string;
  setTitlePageAuthor: (a: string) => void;
  titlePageOrg: string;
  setTitlePageOrg: (o: string) => void;
  titlePageDate: string;
  setTitlePageDate: (d: string) => void;
  addLog: (msg: string) => void;
  showNotice: (text: string, type?: 'success' | 'info' | 'error') => void;
}

export default function BrutalistSidebar({
  styleTab,
  setStyleTab,
  styleTheme,
  setStyleTheme,
  fontSizeBody,
  setFontSizeBody,
  spacingLineHeight,
  setSpacingLineHeight,
  orientation,
  setOrientation,
  marginSize,
  setMarginSize,
  includeToc,
  setIncludeToc,
  pageNumbers,
  setPageNumbers,
  titlePage,
  setTitlePage,
  headerText,
  setHeaderText,
  footerText,
  setFooterText,
  titlePageTitle,
  setTitlePageTitle,
  titlePageSubtitle,
  setTitlePageSubtitle,
  titlePageAuthor,
  setTitlePageAuthor,
  titlePageOrg,
  setTitlePageOrg,
  titlePageDate,
  setTitlePageDate,
  addLog,
  showNotice
}: BrutalistSidebarProps) {
  return (
    <aside className="w-full lg:w-80 border-b-4 lg:border-b-0 lg:border-r-4 border-[#141414] bg-[#D8D7D3] p-4 flex flex-col space-y-4 overflow-y-auto shrink-0 select-none">
      
      {/* Settings Sub-Tabs */}
      <h3 className="text-[10px] uppercase font-black tracking-widest text-[#141414] mb-1">Налаштування</h3>
      
      <div className="flex border-2 border-[#141414] bg-white p-0.5 select-none">
        {(['theme', 'layout', 'cover'] as const).map((tab) => (
          <button
            key={tab}
            onClick={() => setStyleTab(tab)}
            className={`flex-1 text-center py-1.5 text-[9px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${
              styleTab === tab
                ? 'bg-[#141414] text-[#E4E3E0]'
                : 'text-neutral-600 hover:bg-[#D8D7D3]/50'
            }`}
          >
            {tab === 'theme' ? '🎨 Тема' : tab === 'layout' ? '📄 Параметри' : '🏷️ Обкладинка'}
          </button>
        ))}
      </div>

      {styleTab === 'theme' && (
        <div className="space-y-4">
          <div>
            <h4 className="text-[10px] uppercase font-black mb-2 tracking-widest text-[#141414]">Палітра та шрифти готового файлу</h4>
            <div className="grid grid-cols-2 gap-2">
              {Object.entries(THEME_PRESETS).map(([key, preset]) => {
                const isSelected = styleTheme === key;
                return (
                  <button
                    key={key}
                    onClick={() => {
                      setStyleTheme(key);
                      showNotice(`Вибрано оформлення: ${preset.name}`, 'info');
                      addLog(`Тема змінена: ${preset.name}`);
                    }}
                    className={`p-2.5 border-2 text-left flex flex-col justify-between h-20 transition-all cursor-pointer ${
                      isSelected
                        ? 'border-[#141414] bg-white shadow-[2px_2px_0px_0px_#141414]'
                        : 'border-[#141414]/30 bg-white/50 hover:bg-white hover:border-[#141414]'
                    }`}
                  >
                    <div className="flex items-center gap-1">
                      <div
                        className="w-3 h-3 rounded-full border border-[#141414]"
                        style={{ backgroundColor: `#${preset.primaryColor}` }}
                      />
                      <div
                        className="w-3 h-3 rounded-full border border-[#141414] -ml-1.5"
                        style={{ backgroundColor: `#${preset.accentColor}` }}
                      />
                      <span className="text-[9px] font-black text-[#141414] leading-none truncate ml-1">{preset.name.split(' ')[0]}</span>
                    </div>
                    <div className="text-[8px] font-mono text-neutral-500 leading-tight">
                      <span className="font-bold block text-[#141414]">{preset.fontHeading}</span>
                      <span className="block text-[7px] truncate">{preset.fontBody}</span>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="space-y-2 pt-2 border-t-2 border-[#141414]/10">
            <div>
              <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1">
                Розмір шрифту тіла ({fontSizeBody}pt)
              </label>
              <input
                type="range"
                min="10"
                max="16"
                step="1"
                value={fontSizeBody}
                onChange={(e) => setFontSizeBody(Number(e.target.value))}
                className="w-full accent-[#141414] h-1.5 cursor-pointer border border-[#141414] rounded-sm bg-white"
              />
            </div>
            <div>
              <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1">
                Інтервал рядка
              </label>
              <select
                className="w-full text-xs p-1.5 border-2 border-[#141414] bg-white font-mono uppercase text-[#141414] outline-none"
                value={spacingLineHeight}
                onChange={(e) => setSpacingLineHeight(Number(e.target.value))}
              >
                <option value="1.0">Одинарний (1.0)</option>
                <option value="1.15">Компактний (1.15)</option>
                <option value="1.25">Комфортний (1.25)</option>
                <option value="1.5">Півторачний (1.5)</option>
                <option value="2.0">Подвійний (2.0)</option>
              </select>
            </div>
          </div>
        </div>
      )}

      {styleTab === 'layout' && (
        <div className="space-y-4">
          <div>
            <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1.5">
              Орієнтація аркуша:
            </label>
            <div className="grid grid-cols-2 gap-1.5">
              <button
                onClick={() => {
                  setOrientation('portrait');
                  addLog("Орієнтація змінена: Книжкова");
                }}
                className={`px-3 py-1.5 text-[10px] font-extrabold uppercase border-2 transition-all cursor-pointer ${
                  orientation === 'portrait'
                    ? 'bg-[#141414] text-white border-[#141414]'
                    : 'bg-white text-[#141414] border-[#141414] hover:bg-neutral-100'
                }`}
              >
                ↕ Книжкова
              </button>
              <button
                onClick={() => {
                  setOrientation('landscape');
                  addLog("Орієнтація змінена: Альбомна");
                }}
                className={`px-3 py-1.5 text-[10px] font-extrabold uppercase border-2 transition-all cursor-pointer ${
                  orientation === 'landscape'
                    ? 'bg-[#141414] text-white border-[#141414]'
                    : 'bg-white text-[#141414] border-[#141414] hover:bg-neutral-100'
                }`}
              >
                ↔ Альбомна
              </button>
            </div>
          </div>

          <div>
            <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1.5">Поля сторінки:</label>
            <div className="grid grid-cols-3 gap-1">
              {(['normal', 'narrow', 'wide'] as const).map((size) => (
                <button
                  key={size}
                  onClick={() => {
                    setMarginSize(size);
                    addLog(`Поля змінені: ${size}`);
                  }}
                  className={`py-1 bg-white text-[#141414] border-2 text-[9px] font-extrabold uppercase transition-all cursor-pointer block text-center ${
                    marginSize === size
                      ? 'bg-[#141414] text-white border-[#141414]'
                      : 'border-[#141414] hover:bg-neutral-100'
                  }`}
                >
                  {size === 'normal' ? 'Звичайні' : size === 'narrow' ? 'Вузькі' : 'Широкі'}
                </button>
              ))}
            </div>
          </div>

          <div className="space-y-2 pt-2 border-t border-[#141414]/10">
            <button
              onClick={() => {
                setIncludeToc(!includeToc);
                addLog(`Зміст: ${!includeToc ? 'Включено' : 'Вимкнено'}`);
              }}
              className="flex items-center justify-between w-full p-2 bg-white border-2 border-[#141414] text-[10px] font-bold uppercase transition-all hover:bg-neutral-100 cursor-pointer"
            >
              <span>Авто-Зміст (TOC)</span>
              <span className={`w-8 h-4 border-2 border-[#141414] flex items-center p-0.5 ${includeToc ? 'bg-[#141414] justify-end' : 'bg-white justify-start'}`}>
                <span className={`w-2 h-2 ${includeToc ? 'bg-white' : 'bg-[#141414]'}`} />
              </span>
            </button>

            <button
              onClick={() => {
                setPageNumbers(!pageNumbers);
                addLog(`Нумерація: ${!pageNumbers ? 'Включено' : 'Вимкнено'}`);
              }}
              className="flex items-center justify-between w-full p-2 bg-white border-2 border-[#141414] text-[10px] font-bold uppercase transition-all hover:bg-neutral-100 cursor-pointer"
            >
              <span>Нумерація сторінок</span>
              <span className={`w-8 h-4 border-2 border-[#141414] flex items-center p-0.5 ${pageNumbers ? 'bg-[#141414] justify-end' : 'bg-white justify-start'}`}>
                <span className={`w-2 h-2 ${pageNumbers ? 'bg-white' : 'bg-[#141414]'}`} />
              </span>
            </button>

            <button
              onClick={() => {
                setTitlePage(!titlePage);
                addLog(`Титульний лист: ${!titlePage ? 'Включено' : 'Вимкнено'}`);
              }}
              className="flex items-center justify-between w-full p-2 bg-white border-2 border-[#141414] text-[10px] font-bold uppercase transition-all hover:bg-neutral-100 cursor-pointer"
            >
              <span>Титульний лист</span>
              <span className={`w-8 h-4 border-2 border-[#141414] flex items-center p-0.5 ${titlePage ? 'bg-[#141414] justify-end' : 'bg-white justify-start'}`}>
                <span className={`w-2 h-2 ${titlePage ? 'bg-white' : 'bg-[#141414]'}`} />
              </span>
            </button>
          </div>

          <div className="space-y-2 pt-2 border-t border-[#141414]/10">
            <div>
              <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1">Верхній колонтитул:</label>
              <input
                type="text"
                className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] font-mono outline-none"
                placeholder="Введіть текст..."
                value={headerText}
                onChange={(e) => setHeaderText(e.target.value)}
              />
            </div>
            <div>
              <label className="block text-[9px] uppercase font-black text-[#141414] opacity-70 mb-1">Нижній колонтитул:</label>
              <input
                type="text"
                className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] font-mono outline-none"
                placeholder="Введіть текст..."
                value={footerText}
                onChange={(e) => setFooterText(e.target.value)}
              />
            </div>
          </div>
        </div>
      )}

      {styleTab === 'cover' && (
        <div className="space-y-3">
          <div className="p-2.5 bg-white border-2 border-[#141414] flex items-center justify-between">
            <span className="text-[10px] font-black uppercase text-[#141414]">Обкладинка є</span>
            <button
              onClick={() => {
                setTitlePage(!titlePage);
                addLog(`Обкладинка перемкнута: ${!titlePage}`);
              }}
              className={`w-10 h-5 border-2 border-[#141414] flex items-center p-0.5 cursor-pointer ${titlePage ? 'bg-[#141414] justify-end' : 'bg-white justify-start'}`}
            >
              <span className={`w-3 h-3 ${titlePage ? 'bg-white' : 'bg-[#141414]'}`} />
            </button>
          </div>

          {titlePage ? (
            <div className="space-y-2">
              <div>
                <label className="block text-[9px] uppercase font-bold text-[#141414] opacity-70 mb-1">Заголовок обкладинки:</label>
                <input
                  type="text"
                  value={titlePageTitle}
                  onChange={(e) => setTitlePageTitle(e.target.value)}
                  className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] outline-none font-mono"
                  placeholder="..."
                />
              </div>
              <div>
                <label className="block text-[9px] uppercase font-bold text-[#141414] opacity-70 mb-1">Підзаголовок:</label>
                <input
                  type="text"
                  value={titlePageSubtitle}
                  onChange={(e) => setTitlePageSubtitle(e.target.value)}
                  className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] outline-none font-mono"
                  placeholder="..."
                />
              </div>
              <div>
                <label className="block text-[9px] uppercase font-bold text-[#141414] opacity-70 mb-1">Виконавець (Автор):</label>
                <input
                  type="text"
                  value={titlePageAuthor}
                  onChange={(e) => setTitlePageAuthor(e.target.value)}
                  className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] outline-none font-mono"
                  placeholder="..."
                />
              </div>
              <div>
                <label className="block text-[9px] uppercase font-bold text-[#141414] opacity-70 mb-1">Організація:</label>
                <input
                  type="text"
                  value={titlePageOrg}
                  onChange={(e) => setTitlePageOrg(e.target.value)}
                  className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] outline-none font-mono"
                  placeholder="..."
                />
              </div>
              <div>
                <label className="block text-[9px] uppercase font-bold text-[#141414] opacity-70 mb-1">Дата документу:</label>
                <input
                  type="text"
                  value={titlePageDate}
                  onChange={(e) => setTitlePageDate(e.target.value)}
                  className="w-full text-xs px-2 py-1 bg-white border-2 border-[#141414] text-[#141414] outline-none font-mono"
                />
              </div>
            </div>
          ) : (
            <div className="p-4 border-2 border-dashed border-[#141414]/30 text-center text-xs text-neutral-500 uppercase tracking-wide">
              Титульний лист вимкнено.<br/>Ввімкніть його вище, щоб заповнити.
            </div>
          )}
        </div>
      )}
    </aside>
  );
}
