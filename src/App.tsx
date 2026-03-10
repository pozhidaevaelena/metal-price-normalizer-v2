/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx-js-style';
import JSZip from 'jszip';
import { 
  Upload, 
  Play, 
  Download, 
  Settings, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  Trash2,
  Layers,
  FileSpreadsheet,
  Globe
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface LogEntry {
  id: string;
  message: string;
  type: 'info' | 'success' | 'error' | 'warning';
  timestamp: Date;
}

interface ProcessingOptions {
  removeHeader: boolean;
  mergeColumns: boolean;
  splitTabs: boolean;
  outputFilename: string;
}

export default function App() {
  const [zipFile, setZipFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [resultBlob, setResultBlob] = useState<Blob | null>(null);
  const [progress, setProgress] = useState(0);
  const [manualUrl, setManualUrl] = useState('');
  const [showManualInput, setShowManualInput] = useState(false);
  
  const [options, setOptions] = useState<ProcessingOptions>({
    removeHeader: true,
    mergeColumns: true,
    splitTabs: true,
    outputFilename: 'metal_prices_normalized.xlsx'
  });
  const [showAdvanced, setShowAdvanced] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const logEndRef = useRef<HTMLDivElement>(null);

  const addLog = useCallback((message: string, type: LogEntry['type'] = 'info') => {
    setLogs(prev => [
      ...prev,
      {
        id: Math.random().toString(36).substring(7),
        message,
        type,
        timestamp: new Date()
      }
    ]);
  }, []);

  const scrollToBottom = useCallback(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, []);

  const getCellValue = (cell: any) => {
    if (!cell) return '';
    
    // Сначала пробуем получить форматированное значение (с переносами)
    if (cell.w) {
      return cell.w;  // w - это отформатированный текст, сохраняет переносы
    }
    
    // Если нет форматированного, берем сырое значение
    if (cell.v !== undefined && cell.v !== null) {
      return String(cell.v);
    }
    
    return '';
  };

  const isCategory = (text: string, rowValues: any[], isLeft: boolean) => {
    if (!text) return false;
    const val = text.toString().trim();
    if (!val || val.length < 3) return false;
    
    // Проверяем, пусты ли колонки характеристик и цены для этой группы
    const otherIndices = isLeft ? [1, 2, 3] : [6, 7, 8];
    let hasOtherData = false;
    for (const idx of otherIndices) {
      const cellVal = rowValues[idx];
      if (cellVal !== undefined && cellVal !== null && cellVal.toString().trim() !== '') {
        hasOtherData = true;
        break;
      }
    }

    // Если есть текст в названии, но нет данных в характеристиках и цене - это категория
    // Также исключаем чисто технические строки (короткие или состоящие из спецсимволов)
    if (!hasOtherData) {
      // Дополнительная проверка: заголовок обычно не состоит только из цифр
      if (/^\d+$/.test(val.replace(/[\s.,/-]/g, ''))) return false;
      return true;
    }

    return false;
  };

  React.useEffect(() => {
    scrollToBottom();
  }, [logs, scrollToBottom]);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && file.name.toLowerCase().endsWith('.zip')) {
      setZipFile(file);
      setResultBlob(null);
      setLogs([]);
      addLog(`Выбран файл: ${file.name}`, 'info');
    } else if (file) {
      addLog('Пожалуйста, выберите ZIP-архив', 'error');
    }
  };

  const downloadPriceFromMcRu = async () => {
    setIsProcessing(true);
    setResultBlob(null);
    setLogs([]);
    setShowManualInput(false);
    addLog("Начинаю загрузку прайса с mc.ru...", "info");
    
    // ТОЧНАЯ ссылка, которая гарантированно работает
    const directUrl = 'https://mc.ru/prices/metserv.zip';
    
    try {
      addLog(`Скачиваю по прямой ссылке...`, "info");
      
      // Используем прокси для обхода CORS
      const proxyUrl = `/api/proxy?url=${encodeURIComponent(directUrl)}`;
      const response = await fetch(proxyUrl);
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`Ошибка прокси (${response.status}): ${errorData.error || response.statusText}`);
      }
      
      const contentType = response.headers.get('content-type');
      if (contentType && !contentType.includes('zip') && !contentType.includes('octet-stream') && !contentType.includes('application/x-zip-compressed')) {
        console.warn('Unexpected content type:', contentType);
        // Мы все равно попробуем, но предупредим
      }

      const blob = await response.blob();
      
      if (blob.size < 1000) {
        // Слишком маленький файл для ZIP с прайсами, скорее всего это ошибка в виде текста
        const text = await blob.text();
        if (text.includes('error') || text.includes('<!DOCTYPE')) {
          throw new Error('Получен некорректный файл (возможно, ошибка сервера или блокировка)');
        }
      }

      const fileName = 'metserv.zip';
      const zipFile = new File([blob], fileName, { type: "application/zip" });
      
      setZipFile(zipFile);
      addLog(`✅ Файл загружен (${(blob.size / 1024 / 1024).toFixed(2)} MB)`, "success");
      addLog(`Начинаю обработку...`, "info");
      
      await processZip(zipFile);
      
    } catch (error) {
      console.error(error);
      addLog(`❌ Ошибка при загрузке: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, "error");
      addLog("Попробуй скачать файл вручную и выбрать его кнопкой выше или вставь ссылку ниже.", "warning");
      setShowManualInput(true);
      setIsProcessing(false);
    }
  };

  const downloadFromManualUrl = async () => {
    if (!manualUrl.trim()) {
      addLog("Введите ссылку", "error");
      return;
    }
    
    setIsProcessing(true);
    setResultBlob(null);
    addLog(`Скачиваю по ручной ссылке: ${manualUrl}`, "info");
    
    try {
      const zipProxyUrl = `/api/proxy?url=${encodeURIComponent(manualUrl.trim())}`;
      const response = await fetch(zipProxyUrl);
      if (!response.ok) throw new Error(`Ошибка ${response.status}`);
      
      const blob = await response.blob();
      const fileName = manualUrl.split('/').pop()?.split('?')[0] || 'metserv.zip';
      const zipFile = new File([blob], fileName, { type: "application/zip" });
      
      setZipFile(zipFile);
      addLog(`✅ Файл загружен (${(blob.size / 1024 / 1024).toFixed(2)} MB)`, "success");
      await processZip(zipFile);
      setShowManualInput(false);
    } catch (error) {
      addLog(`❌ Ошибка при скачивании: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, "error");
      setIsProcessing(false);
    }
  };

  const processZip = async (fileToProcess?: File) => {
    const targetFile = fileToProcess || zipFile;
    if (!targetFile) return;

    setIsProcessing(true);
    setResultBlob(null);
    setProgress(0);
    if (!fileToProcess) setLogs([]); 
    addLog('Начало обработки ZIP-архива...', 'info');

    try {
      const zip = new JSZip();
      const content = await zip.loadAsync(targetFile);
      const excelFiles = Object.keys(content.files).filter(name => 
        (name.toLowerCase().endsWith('.xlsx') || name.toLowerCase().endsWith('.xls')) && 
        !name.includes('__MACOSX')
      );

      if (excelFiles.length === 0) {
        throw new Error('В архиве не найдено Excel-файлов (.xls, .xlsx)');
      }

      addLog(`Найдено файлов: ${excelFiles.length}`, 'info');

      const workbook = XLSX.utils.book_new();
      const allDataByCategory: Record<string, { title: string; data: any[][] }> = {};

      for (let i = 0; i < excelFiles.length; i++) {
        const fileName = excelFiles[i];
        try {
          const fileData = await content.files[fileName].async('arraybuffer');
          addLog(`Обработка файла: ${fileName}...`, 'info');
          
          const categoryName = fileName.split('/').pop()?.split('.')[0] || 'Без категории';
          const formattedCategory = categoryName.charAt(0).toUpperCase() + categoryName.slice(1);

          const sourceWb = XLSX.read(fileData, { 
            type: 'array',
            cellText: true,
            cellDates: true,
            raw: true,
            cellNF: true,
            sheetStubs: true,
          });
          
          const firstSheetName = sourceWb.SheetNames[0];
          const worksheet = sourceWb.Sheets[firstSheetName];
          
          // 2. Читаем строки ВРУЧНУЮ по диапазону
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          let rows: any[][] = [];
          
          for (let R = range.s.r; R <= range.e.r; ++R) {
            const row: any[] = [];
            for (let C = 0; C <= 9; C++) {
              const cell_address = { c: C, r: R };
              const cell_ref = XLSX.utils.encode_cell(cell_address);
              const cell = worksheet[cell_ref];
              row.push(getCellValue(cell));
            }
            rows.push(row);
          }

          // 3. Удаляем шапку
          if (options.removeHeader && rows.length > 4) {
            rows = rows.slice(4);
          }

          // 4. Разбиваем на страницы по пустым строкам
          const pages: { rowNum: number; values: any[] }[][] = [];
          let currentPage: { rowNum: number; values: any[] }[] = [];
          
          for (let j = 0; j < rows.length; j++) {
            const row = rows[j];
            
            // Проверяем, пустая ли строка (ВСЕ колонки от A до J пустые)
            let isEmpty = true;
            for (let k = 0; k <= 9; k++) {
              const val = row[k];
              if (val !== undefined && val !== null && String(val).trim() !== '') {
                isEmpty = false;
                break;
              }
            }
            
            if (isEmpty) {
              if (currentPage.length > 0) {
                pages.push(currentPage);
                currentPage = [];
              }
            } else {
              currentPage.push({ rowNum: j + 5, values: row });
            }
          }
          if (currentPage.length > 0) pages.push(currentPage);

          addLog(`Файл "${fileName}": найдено страниц - ${pages.length}`, 'info');
          pages.forEach((p, idx) => {
            if (idx < 3 || idx === pages.length - 1) {
               addLog(`  Стр ${idx + 1}: ${p.length} строк`, 'info');
            } else if (idx === 3) {
               addLog(`  ...`, 'info');
            }
          });

          // 5. Обрабатываем страницы
          const normalizedRows: any[][] = [];
          let currentCategory = 'Без категории';
          let sheetTitle = '';

          // Пытаемся достать заголовок из 3-й строки оригинала
          const titleCell = worksheet['A3'] || worksheet['B3'] || worksheet['C3'];
          sheetTitle = getCellValue(titleCell) || formattedCategory;

          for (let p = 0; p < pages.length; p++) {
            const page = pages[p];
            let leftCount = 0;
            let rightCount = 0;

            // ЭТАП 1: Обрабатываем левую колонку (A-D) текущей страницы
            for (let r = 0; r < page.length; r++) {
              const row = page[r];
              const valA = row.values[0] ? row.values[0].toString().trim() : '';
              const valB = row.values[1] ? row.values[1].toString().trim() : '';
              const valC = row.values[2] ? row.values[2].toString().trim() : '';
              const valD = row.values[3] ? row.values[3].toString().trim() : '';
              
              if (valA && isCategory(valA, row.values, true)) {
                const cleaned = valA.replace(/\s*\(продолжение\)\s*$/i, '').trim();
                if (cleaned.toLowerCase() !== currentCategory.toLowerCase()) {
                  currentCategory = cleaned;
                  normalizedRows.push([currentCategory, '', '', '', 'ПОДКАТЕГОРИЯ']);
                }
                continue;
              }
              
              // Если есть наименование (A) и хотя бы одно другое поле (B, C или D)
              if (valA && (valB || valC || valD)) {
                normalizedRows.push([valA, valB, valC, valD, currentCategory]);
                leftCount++;
              }
            }

            // ЭТАП 2: Обрабатываем правую колонку (F-I) текущей страницы
            if (options.mergeColumns) {
              for (let r = 0; r < page.length; r++) {
                const row = page[r];
                const valF = row.values[5] ? row.values[5].toString().trim() : '';
                const valG = row.values[6] ? row.values[6].toString().trim() : '';
                const valH = row.values[7] ? row.values[7].toString().trim() : '';
                const valI = row.values[8] ? row.values[8].toString().trim() : '';
                
                if (valF && isCategory(valF, row.values, false)) {
                  const cleaned = valF.replace(/\s*\(продолжение\)\s*$/i, '').trim();
                  if (cleaned.toLowerCase() !== currentCategory.toLowerCase()) {
                    currentCategory = cleaned;
                    normalizedRows.push([currentCategory, '', '', '', 'ПОДКАТЕГОРИЯ']);
                  }
                  continue;
                }
                
                if (valF && (valG || valH || valI)) {
                  // Данные из F-I попадают в A-D выходного файла
                  normalizedRows.push([valF, valG, valH, valI, currentCategory]);
                  rightCount++;
                }
              }
            }
            
            if (leftCount > 0 || rightCount > 0) {
              addLog(`  Стр ${p + 1}: добавлено ${leftCount} (лев) + ${rightCount} (прав) записей`, 'info');
            }
          }

          // Добавляем в общую структуру
          const targetSheet = options.splitTabs ? formattedCategory : 'Общий список';
          if (!allDataByCategory[targetSheet]) {
            allDataByCategory[targetSheet] = {
              title: sheetTitle,
              data: []
            };
          }
          allDataByCategory[targetSheet].data.push(...normalizedRows);

          addLog(`✅ Файл "${fileName}" успешно обработан.`, 'success');
        } catch (fileError) {
          addLog(`❌ Ошибка в файле "${fileName}": ${fileError instanceof Error ? fileError.message : 'Неизвестная ошибка'}`, 'error');
        }
        setProgress(Math.round(((i + 1) / excelFiles.length) * 100));
      }

      // Финальная сборка XLSX
      if (Object.keys(allDataByCategory).length === 0) {
        throw new Error('Нет данных для сохранения');
      }

      Object.entries(allDataByCategory).forEach(([sheetName, sheetInfo]: [string, any]) => {
        const { title, data } = sheetInfo;
        
        // Определяем заголовок для Характеристики 1 в зависимости от названия вкладки
        const getChar1Header = (name: string) => {
          const n = name.toLowerCase();
          if (n.includes('cvetmet')) return "Диаметр";
          if (n.includes('engineering')) return "Диаметр";
          if (n.includes('kachestvst')) return "Диаметр";
          if (n.includes('krepezh')) return "Длина";
          if (n.includes('listovojprokat')) return "Толщина";
          if (n.includes('metizy')) return "Диаметр";
          if (n.includes('nerzhaveika')) return "Толщина";
          if (n.includes('truby')) return "Стенка";
          return "Характеристика 1";
        };

        const char1 = getChar1Header(sheetName);
        const headers = ["Наименование", char1, "Ед. изм.", "Цена", "Подкатегория", "Количество", "Стоимость"];
        
        // Формируем массив строк: 1-я заголовок, 2-я шапка, далее данные
        const aoa = [
          [title, "", "", "", "", "", ""], // Строка 1 (будет объединена)
          headers                          // Строка 2
        ];
        
        data.forEach((row: any[]) => {
          aoa.push([...row, "", ""]); // Добавляем пустые ячейки для Кол-во и Стоимость
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        
        // Стили
        const borderStyle = {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        };

        const titleStyle = {
          font: { bold: true, size: 14 },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: "F3F4F6" } }
        };

        const headerStyle = {
          fill: { fgColor: { rgb: "4F46E5" } },
          font: { bold: true, color: { rgb: "FFFFFF" } },
          border: borderStyle,
          alignment: { horizontal: "center", vertical: "center", wrapText: true }
        };

        const categoryRowStyle = {
          fill: { fgColor: { rgb: "EEF2FF" } },
          font: { bold: true, italic: true, color: { rgb: "1E40AF" } },
          border: borderStyle,
          alignment: { vertical: "center", wrapText: true }
        };

        const dataCellStyle = {
          border: borderStyle,
          alignment: { vertical: "center", wrapText: true }
        };

        const priceStyle = {
          border: borderStyle,
          numFmt: "#,##0.00",
          alignment: { vertical: "center" }
        };

        // Объединяем первую строку
        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];

        const wsRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let R = wsRange.s.r; R <= wsRange.e.r; ++R) {
          for (let C = wsRange.s.c; C <= wsRange.e.c; ++C) {
            const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
            const cell = ws[cell_ref];
            if (!cell) continue;

            if (R === 0) {
              cell.s = titleStyle;
            } else if (R === 1) {
              cell.s = headerStyle;
            } else {
              const rowData = data[R - 2];
              const isSubcategory = rowData && rowData[4] === 'ПОДКАТЕГОРИЯ';
              
              if (isSubcategory) {
                cell.s = categoryRowStyle;
              } else {
                cell.s = (C === 3 || C === 6) ? priceStyle : dataCellStyle;
                
                // Добавляем формулу в колонку G (Стоимость) = D * F с защитой от ошибок #VALUE!
                if (C === 6 && !isSubcategory) {
                  const rowNum = R + 1;
                  cell.f = `IFERROR(IF(F${rowNum}="", "", D${rowNum}*F${rowNum}), "")`;
                  cell.t = 'n';
                }
              }
            }
          }
        }

        // Закрепление первых двух строк
        ws['!views'] = [{ state: 'frozen', ySplit: 2 }];
        
        // Фильтры на вторую строку
        ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 1, c: 0 }, e: { r: wsRange.e.r, c: 6 } }) };

        ws['!cols'] = [
          { wch: 50 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 25 }, { wch: 12 }, { wch: 15 }
        ];

        XLSX.utils.book_append_sheet(workbook, ws, sheetName.substring(0, 31));
      });

      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      setResultBlob(new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      addLog('Все файлы обработаны. Итоговый файл готов!', 'success');

    } catch (error) {
      addLog(`Критическая ошибка: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, 'error');
    } finally {
      setIsProcessing(false);
      setProgress(100);
    }
  };

  const downloadFile = () => {
    if (!resultBlob) return;
    const url = URL.createObjectURL(resultBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = options.outputFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-2 rounded-lg">
            <FileSpreadsheet className="text-white w-6 h-6" />
          </div>
          <h1 className="text-xl font-semibold tracking-tight">
            UniversalMetalPriceNormalizer <span className="text-gray-400 font-normal ml-2">| Нормализатор прайс-листов</span>
          </h1>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={downloadPriceFromMcRu}
            disabled={isProcessing}
            className={cn(
              "flex items-center gap-2 px-4 py-2 rounded-lg transition-all text-sm font-medium border",
              isProcessing
                ? "bg-gray-50 text-gray-400 border-gray-200 cursor-not-allowed"
                : "bg-white text-indigo-600 border-indigo-200 hover:bg-indigo-50 hover:border-indigo-300"
            )}
          >
            <Globe size={18} />
            Загрузить с mc.ru
          </button>
          {resultBlob && (
            <motion.button
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              onClick={downloadFile}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all text-sm font-semibold shadow-lg animate-pulse hover:animate-none"
            >
              <Download size={18} />
              Скачать результат
            </motion.button>
          )}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors text-sm font-medium"
          >
            <Upload size={18} />
            Выбрать ZIP
          </button>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileSelect}
            accept=".zip"
            className="hidden"
          />
          <button
            onClick={() => processZip()}
            disabled={!zipFile || isProcessing}
            className={cn(
              "flex items-center gap-2 px-6 py-2 rounded-lg transition-all text-sm font-semibold shadow-sm",
              !zipFile || isProcessing
                ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                : "bg-indigo-600 text-white hover:bg-indigo-700 active:scale-95"
            )}
          >
            {isProcessing ? <Loader2 className="animate-spin" size={18} /> : <Play size={18} />}
            Обработать
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Settings Panel */}
        <aside className="lg:col-span-4 space-y-6">
          <section className="bg-white p-6 rounded-2xl border border-gray-200 shadow-sm">
            <button 
              onClick={() => setShowAdvanced(!showAdvanced)}
              className="flex items-center justify-between w-full gap-2 mb-2"
            >
              <div className="flex items-center gap-2">
                <Settings size={20} className="text-indigo-600" />
                <h2 className="font-semibold text-lg">Настройки</h2>
              </div>
              <span className="text-xs text-indigo-600 hover:underline">
                {showAdvanced ? 'Скрыть' : 'Показать'}
              </span>
            </button>

            <AnimatePresence>
              {showAdvanced && (
                <motion.div 
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  className="overflow-hidden"
                >
                  <div className="space-y-4 pt-4">
                    <label className="flex items-start gap-3 cursor-pointer group">
                      <div className="relative flex items-center">
                        <input
                          type="checkbox"
                          checked={options.removeHeader}
                          onChange={e => setOptions(prev => ({ ...prev, removeHeader: e.target.checked }))}
                          className="peer h-5 w-5 cursor-pointer appearance-none rounded border border-gray-300 bg-white checked:bg-indigo-600 checked:border-indigo-600 transition-all"
                        />
                        <CheckCircle2 className="absolute w-5 h-5 text-white scale-0 peer-checked:scale-75 transition-transform left-0" />
                      </div>
                      <div className="flex flex-col">
                        <span className="text-sm font-medium text-gray-700 group-hover:text-indigo-600 transition-colors">Удалить шапку</span>
                        <span className="text-xs text-gray-400">Пропускает первые 4 строки файла</span>
                      </div>
                    </label>

                    <label className="flex items-start gap-3 cursor-pointer group">
                      <div className="relative flex items-center">
                        <input
                          type="checkbox"
                          checked={options.mergeColumns}
                          onChange={e => setOptions(prev => ({ ...prev, mergeColumns: e.target.checked }))}
                          className="peer h-5 w-5 cursor-pointer appearance-none rounded border border-gray-300 bg-white checked:bg-indigo-600 checked:border-indigo-600 transition-all"
                        />
                        <CheckCircle2 className="absolute w-5 h-5 text-white scale-0 peer-checked:scale-75 transition-transform left-0" />
                      </div>
                      <div className="flex flex-col">
                        <span className="text-sm font-medium text-gray-700 group-hover:text-indigo-600 transition-colors">Объединить G-J в A-D</span>
                        <span className="text-xs text-gray-400">Переносит данные из правой части таблицы</span>
                      </div>
                    </label>

                    <label className="flex items-start gap-3 cursor-pointer group">
                      <div className="relative flex items-center">
                        <input
                          type="checkbox"
                          checked={options.splitTabs}
                          onChange={e => setOptions(prev => ({ ...prev, splitTabs: e.target.checked }))}
                          className="peer h-5 w-5 cursor-pointer appearance-none rounded border border-gray-300 bg-white checked:bg-indigo-600 checked:border-indigo-600 transition-all"
                        />
                        <CheckCircle2 className="absolute w-5 h-5 text-white scale-0 peer-checked:scale-75 transition-transform left-0" />
                      </div>
                      <div className="flex flex-col">
                        <span className="text-sm font-medium text-gray-700 group-hover:text-indigo-600 transition-colors">Разделить на вкладки</span>
                        <span className="text-xs text-gray-400">Создает отдельный лист для каждого файла</span>
                      </div>
                    </label>

                    <div className="pt-4 border-t border-gray-100">
                      <label className="block text-sm font-medium text-gray-700 mb-2">Название итогового файла</label>
                      <div className="relative">
                        <input
                          type="text"
                          value={options.outputFilename}
                          onChange={e => setOptions(prev => ({ ...prev, outputFilename: e.target.value }))}
                          className="w-full pl-3 pr-10 py-2 bg-gray-50 border border-gray-200 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                        />
                        <FileText className="absolute right-3 top-2.5 text-gray-400" size={16} />
                      </div>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </section>

          {resultBlob && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-indigo-50 p-6 rounded-2xl border border-indigo-100 shadow-sm"
            >
              <h3 className="text-indigo-900 font-semibold mb-2 flex items-center gap-2">
                <CheckCircle2 size={18} />
                Результат готов
              </h3>
              <p className="text-indigo-700 text-sm mb-4">
                Все файлы успешно обработаны и объединены в один документ.
              </p>
              <button
                onClick={downloadFile}
                className="w-full flex items-center justify-center gap-2 py-3 bg-indigo-600 text-white rounded-xl font-semibold hover:bg-indigo-700 transition-all shadow-md active:scale-95"
              >
                <Download size={20} />
                Скачать XLSX
              </button>
            </motion.div>
          )}
        </aside>

        {/* Log Area */}
        <section className="lg:col-span-8 flex flex-col h-[calc(100vh-180px)]">
          <div className="bg-white rounded-2xl border border-gray-200 shadow-sm flex flex-col h-full overflow-hidden">
            <button 
              onClick={() => setShowAdvanced(!showAdvanced)}
              className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50 w-full"
            >
              <div className="flex items-center gap-2">
                <Layers size={18} className="text-gray-500" />
                <h2 className="font-semibold text-gray-700">Лог обработки</h2>
              </div>
              <div className="flex items-center gap-4">
                {logs.length > 0 && (
                  <span 
                    onClick={(e) => { e.stopPropagation(); setLogs([]); }}
                    className="text-xs text-gray-400 hover:text-red-500 flex items-center gap-1 transition-colors"
                  >
                    <Trash2 size={14} />
                    Очистить
                  </span>
                )}
                <span className="text-xs text-indigo-600 hover:underline">
                  {showAdvanced ? 'Скрыть' : 'Показать'}
                </span>
              </div>
            </button>

            <AnimatePresence>
              {showAdvanced && (
                <motion.div 
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  className="flex-1 overflow-y-auto p-6 space-y-3 font-mono text-sm bg-[#FCFCFD] min-h-[200px]"
                >
                  {logs.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-gray-400 space-y-4 py-10">
                      <div className="p-4 bg-gray-50 rounded-full">
                        <Upload size={32} />
                      </div>
                      <p className="text-center max-w-xs">
                        Логи появятся здесь после начала обработки.
                      </p>
                    </div>
                  ) : (
                    logs.map((log) => (
                      <motion.div
                        key={log.id}
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        className={cn(
                          "flex gap-3 p-3 rounded-lg border",
                          log.type === 'info' && "bg-blue-50/50 border-blue-100 text-blue-800",
                          log.type === 'success' && "bg-emerald-50/50 border-emerald-100 text-emerald-800",
                          log.type === 'error' && "bg-red-50/50 border-red-100 text-red-800",
                          log.type === 'warning' && "bg-amber-50/50 border-amber-100 text-amber-800"
                        )}
                      >
                        <div className="mt-0.5">
                          {log.type === 'info' && <FileText size={16} />}
                          {log.type === 'success' && <CheckCircle2 size={16} />}
                          {log.type === 'error' && <AlertCircle size={16} />}
                          {log.type === 'warning' && <AlertCircle size={16} />}
                        </div>
                        <div className="flex-1">
                          <div className="flex justify-between items-start mb-1">
                            <span className="font-bold opacity-50 text-[10px] uppercase tracking-wider">
                              {log.timestamp.toLocaleTimeString([], { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                            </span>
                          </div>
                          <p className="leading-relaxed">{log.message}</p>
                        </div>
                      </motion.div>
                    ))
                  )}
                  <div ref={logEndRef} />
                </motion.div>
              )}
            </AnimatePresence>

            {!showAdvanced && logs.length > 0 && (
              <div className="px-6 py-3 bg-gray-50 border-t border-gray-100 flex items-center justify-between">
                <span className="text-xs text-gray-500 truncate max-w-[80%]">
                  Последнее: {logs[logs.length - 1].message}
                </span>
                {isProcessing && <Loader2 className="animate-spin text-indigo-600" size={14} />}
              </div>
            )}

            {isProcessing && (
              <div className="px-6 py-4 bg-white border-t border-gray-100">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-semibold text-indigo-600 uppercase tracking-wider">Прогресс обработки</span>
                  <span className="text-xs font-bold text-indigo-600">{progress}%</span>
                </div>
                <div className="w-full bg-gray-100 rounded-full h-2 overflow-hidden">
                  <motion.div 
                    className="bg-indigo-600 h-full"
                    initial={{ width: 0 }}
                    animate={{ width: `${progress}%` }}
                    transition={{ type: 'spring', bounce: 0, duration: 0.5 }}
                  />
                </div>
              </div>
            )}

            {showManualInput && (
              <motion.div 
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                className="px-6 py-4 bg-gray-50 border-t border-gray-100 space-y-3"
              >
                <p className="text-sm font-semibold text-gray-700">Введите прямую ссылку на metserv.zip:</p>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={manualUrl}
                    onChange={e => setManualUrl(e.target.value)}
                    placeholder="https://mc.ru/.../metserv.zip"
                    className="flex-1 px-3 py-2 bg-white border border-gray-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                  <button
                    onClick={downloadFromManualUrl}
                    className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-bold hover:bg-indigo-700 transition-all"
                  >
                    Скачать
                  </button>
                </div>
              </motion.div>
            )}

            {resultBlob && !isProcessing && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="px-6 py-4 bg-emerald-50 border-t border-emerald-100 flex items-center justify-between"
              >
                <div className="flex items-center gap-3 text-emerald-800">
                  <CheckCircle2 size={20} />
                  <span className="font-semibold text-sm">Обработка завершена!</span>
                </div>
                <button
                  onClick={downloadFile}
                  className="flex items-center gap-2 px-6 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all text-sm font-bold shadow-md"
                >
                  <Download size={18} />
                  Скачать итоговый файл
                </button>
              </motion.div>
            )}
          </div>
        </section>
      </main>

      {/* Footer */}
      <footer className="max-w-7xl mx-auto px-6 py-8 text-center text-gray-400 text-xs">
        <p>© 2024 UniversalMetalPriceNormalizer. Все права защищены.</p>
        <p className="mt-1">Инструмент для автоматизации обработки и загрузки прайс-листов металлопроката.</p>
      </footer>
    </div>
  );
}
