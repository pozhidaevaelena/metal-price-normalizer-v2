/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx-js-style';
import JSZip from 'jszip';
import { GoogleGenAI } from "@google/genai";
import { 
  Upload, 
  Play, 
  Download, 
  FileText, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  Trash2,
  Layers,
  FileSpreadsheet,
  Globe,
  Calculator as CalcIcon,
  Plus,
  History,
  Database,
  Package,
  TrendingUp,
  RefreshCw,
  Edit2,
  ChevronDown,
  ChevronUp,
  X,
  Image as ImageIcon,
  Scan,
  Send
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import Tesseract from 'tesseract.js';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Initialize Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || '' });
const model = "gemini-3-flash-preview";

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

interface PriceItem {
  name: string;
  param_raw: string;
  all_params: string[];
  param_name: string;
  price: number;
  category: string;
  subcategory: string;
  supplier: string;
  unit: string;
}

interface CartEntry {
  subcategory: string;
  parameter: string;
  name?: string;
  quantity: number;
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'normalizer' | 'calculator' | 'drawing-calc'>('calculator');
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
  const [supplierName, setSupplierName] = useState('mc.ru');
  const [singleExcelFile, setSingleExcelFile] = useState<File | null>(null);
  
  // Refs
  const fileInputRef = useRef<HTMLInputElement>(null);
  const singleFileInputRef = useRef<HTMLInputElement>(null);
  const logEndRef = useRef<HTMLDivElement>(null);
  
  // Calculator State
  const [diagnostics, setDiagnostics] = useState<any>(null);
  const [diagError, setDiagError] = useState<string | null>(null);
  const [calcSubcategory, setCalcSubcategory] = useState('');
  const [calcParameter, setCalcParameter] = useState('');
  const [calcName, setCalcName] = useState('');
  const [calcQuantity, setCalcQuantity] = useState<string>('');
  const [isCalculating, setIsCalculating] = useState(false);
  const [calcError, setCalcError] = useState<string | null>(null);
  const [editIndex, setEditIndex] = useState<number | null>(null);
  const [isLogExpanded, setIsLogExpanded] = useState(false);
  const [allSubcategories, setAllSubcategories] = useState<string[]>([]);
  const [allPriceItems, setAllPriceItems] = useState<PriceItem[]>([]);
  const [availableNames, setAvailableNames] = useState<string[]>([]);
  const [isSubcategoryDropdownOpen, setIsSubcategoryDropdownOpen] = useState(false);
  const [isNameDropdownOpen, setIsNameDropdownOpen] = useState(false);
  const [cart, setCart] = useState<any[]>([]);
  const [calcResults, setCalcResults] = useState<{ results_max: any[], results_avg: any[] } | null>(null);
  const [showResults, setShowResults] = useState(false);
  const [errors, setErrors] = useState<Record<string, boolean>>({});

  // Drawing Calc State
  const [drawingImage, setDrawingImage] = useState<File | null>(null);
  const [drawingPreview, setDrawingPreview] = useState<string | null>(null);
  const [isRecognizing, setIsRecognizing] = useState(false);
  const [recognizedPositions, setRecognizedPositions] = useState<any[]>([]);
  const [rawOcrText, setRawOcrText] = useState<string>('');
  const [showRawText, setShowRawText] = useState(false);
  const [ocrStatus, setOcrStatus] = useState<string>('');
  const [ocrError, setOcrError] = useState<string | null>(null);
  const drawingInputRef = useRef<HTMLInputElement>(null);

  const fetchDiagnostics = useCallback(async () => {
    try {
      const res = await fetch('/api/diagnostics');
      const data = await res.json();
      if (data.error) {
        setDiagError(data.error);
        setDiagnostics(null);
      } else {
        setDiagnostics(data);
        setDiagError(null);
        // Also fetch subcategories for the dropdown
        const subsRes = await fetch('/api/get_subcategories');
        if (subsRes.ok) {
          const subs = await subsRes.json();
          setAllSubcategories(subs);
        }
      }
    } catch (err) {
      setDiagError("Ошибка подключения к серверу");
    }
  }, []);

  const fetchNames = useCallback(async () => {
    if (!calcSubcategory || !calcParameter) return;
    
    // First, try to get from local state (Vercel compatibility)
    if (allPriceItems.length > 0) {
      const normalizeParam = (p: string) => String(p).replace(',', '.').trim();
      const targetParam = normalizeParam(calcParameter);
      
      const names = Array.from(new Set(
        allPriceItems
          .filter((i: any) => {
            if (i.subcategory !== calcSubcategory) return false;
            const itemParams = (i.all_params || (i.param_raw ? i.param_raw.split(';').map((p: string) => p.trim()) : []))
              .map(normalizeParam);
            return itemParams.includes(targetParam);
          })
          .map((i: any) => i.name)
      )).sort();
      
      if (names.length > 0) {
        setAvailableNames(names);
        return;
      }
    }

    try {
      const res = await fetch('/api/get_names', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ subcategory: calcSubcategory, param: calcParameter })
      });
      if (res.ok) {
        const names = await res.json();
        setAvailableNames(names);
      }
    } catch (err) {
      console.error('Failed to fetch names', err);
    }
  }, [calcSubcategory, calcParameter]);

  React.useEffect(() => {
    if (activeTab === 'calculator') {
      fetchDiagnostics();
    }
  }, [activeTab, fetchDiagnostics]);

  React.useEffect(() => {
    if (calcSubcategory && calcParameter) {
      fetchNames();
    } else {
      setAvailableNames([]);
      setCalcName('');
    }
  }, [calcSubcategory, calcParameter, fetchNames]);

  const handleAddPosition = () => {
    const newErrors: Record<string, boolean> = {};
    if (!calcSubcategory) newErrors.subcategory = true;
    if (!calcParameter) newErrors.parameter = true;
    
    const qty = parseFloat(calcQuantity.replace(',', '.'));
    if (isNaN(qty) || qty <= 0) newErrors.quantity = true;

    if (Object.keys(newErrors).length > 0) {
      setErrors(newErrors);
      return;
    }

    setErrors({});
    const newEntry = {
      subcategory: calcSubcategory,
      param: calcParameter,
      name: calcName || null,
      quantity: qty,
      unit: 'т' // Default, will be updated by server
    };

    if (editIndex !== null) {
      const newCart = [...cart];
      newCart[editIndex] = newEntry;
      setCart(newCart);
      setEditIndex(null);
    } else {
      setCart([...cart, newEntry]);
    }

    // Clear fields
    setCalcSubcategory('');
    setCalcParameter('');
    setCalcName('');
    setCalcQuantity('');
  };

  const handleEditPosition = (index: number) => {
    const item = cart[index];
    setCalcSubcategory(item.subcategory);
    setCalcParameter(item.param);
    setCalcName(item.name || '');
    setCalcQuantity(String(item.quantity));
    setEditIndex(index);
    // Scroll to form
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleDrawingUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setDrawingImage(file);
      const reader = new FileReader();
      reader.onloadend = () => {
        setDrawingPreview(reader.result as string);
      };
      reader.readAsDataURL(file);
      setRecognizedPositions([]);
    }
  };

  const recognizeDrawing = async () => {
    if (!drawingPreview) return;
    setIsRecognizing(true);
    setRawOcrText('');
    setOcrError(null);
    setOcrStatus('Анализ чертежа через ИИ (v2)...');
    console.log("Starting Gemini recognition...");
    
    try {
      if (!process.env.GEMINI_API_KEY) {
        throw new Error("API ключ Gemini не найден в системе.");
      }

      // 1. Prepare image for Gemini
      const base64Data = drawingPreview.split(',')[1];
      const mimeType = drawingPreview.split(';')[0].split(':')[1];

      // 2. Call Gemini for intelligent analysis
      const prompt = `
        Ты — эксперт по чтению технических чертежей металлоконструкций.
        Проанализируй это изображение (чертёж каркаса).
        
        ТВОЯ ЗАДАЧА:
        1. Найди все числовые размеры (длины деталей) в миллиметрах.
        2. Определи тип профиля. На чертеже обычно есть подпись (например, 'труба 25х25х2').
        3. Составь список всех уникальных длин и их количества.
        
        ПРАВИЛА ПОДСЧЕТА:
        - ВНИМАТЕЛЬНО считай количество одинаковых отрезков. Если это 3D каркас, помни про симметрию (передняя/задняя рама, левая/правая сторона).
        - Например, если ты видишь прямоугольное основание со стороной 950мм, их должно быть как минимум 2 (перед и зад).
        - Если материал указан один раз, он относится ко всем деталям на чертеже.
        - Размеры на чертежах ВСЕГДА в мм.
        - Игнорируй мелкие цифры, которые не являются размерами.
        
        ФОРМАТИРОВАНИЕ:
        - subcategory: Используй стандартные названия: 'Трубы', 'Лист', 'Швеллер', 'Уголок', 'Балка', 'Круг'.
        - param: Используй символ 'x' (английский) как разделитель (например, '25x25x2').
        
        ВЕРНИ ОТВЕТ СТРОГО В ФОРМАТЕ JSON (массив объектов):
        [
          {
            "subcategory": "Трубы",
            "param": "25x25x2",
            "rawSize": "950",
            "rawCount": 4,
            "type": "profile"
          }
        ]
        
        Если ничего не найдено, верни пустой массив [].
        Никаких пояснений, только JSON.
      `;

      const response = await ai.models.generateContent({
        model: model,
        contents: [{
          parts: [
            { text: prompt },
            { inlineData: { data: base64Data, mimeType: mimeType } }
          ]
        }],
        config: {
          responseMimeType: "application/json"
        }
      });

      const resultText = response.text || '[]';
      setRawOcrText(resultText);
      console.log("Gemini Response:", resultText);

      let parsedResults: any[] = [];
      try {
        parsedResults = JSON.parse(resultText);
      } catch (e) {
        const jsonMatch = resultText.match(/\[.*\]/s);
        if (jsonMatch) {
          parsedResults = JSON.parse(jsonMatch[0]);
        } else {
          throw new Error("ИИ вернул некорректный формат данных.");
        }
      }

      // 3. Calculate quantities and fetch units
      const finalPositions = await Promise.all(parsedResults.map(async (m: any) => {
        let quantity = 0;
        let unit = m.type === 'sheet' ? 'т' : 'м';

        const size = parseFloat(m.rawSize);
        const count = parseInt(m.rawCount) || 1;

        if (m.type === 'profile') {
          quantity = (size * count) / 1000;
        } else if (m.type === 'sheet') {
          const [l, w] = String(m.rawSize).split(/[xх*×]/).map(parseFloat);
          if (l && w) {
            const area = (l / 1000) * (w / 1000);
            quantity = area * count * parseFloat(m.param) * 7.85;
          } else {
            quantity = 1;
          }
        }

        // Fetch unit from diagnostics with robust matching
        if (diagnostics?.items) {
          const normParam = (p: string) => p.toLowerCase().replace(/[,]/g, '.').replace(/[х*×]/g, 'x').trim();
          const targetP = normParam(m.param);
          const targetS = m.subcategory.toLowerCase().replace(/[ыиа]$/, ''); // fuzzy subcategory

          const match = diagnostics.items.find((i: any) => {
            const itemS = i.subcategory.toLowerCase().replace(/[ыиа]$/, '');
            if (itemS !== targetS) return false;
            
            const itemParams = (i.all_params || (i.param_raw ? i.param_raw.split(';').map((p: string) => p.trim()) : []))
              .map(normParam);
            return itemParams.includes(targetP);
          });
          
          if (match) unit = match.unit || unit;
        }

        return {
          ...m,
          quantity,
          unit,
          description: `${m.subcategory} (${m.param}): ${m.rawSize} мм × ${m.rawCount} шт`
        };
      }));

      setRecognizedPositions(finalPositions);
      if (finalPositions.length > 0) {
        setOcrStatus(`ИИ успешно распознал ${finalPositions.length} позиций.`);
      } else {
        setOcrStatus('Детали на чертеже не обнаружены.');
      }

    } catch (err: any) {
      console.error("Recognition Error:", err);
      setOcrError(err.message || 'Неизвестная ошибка');
      setOcrStatus('Ошибка анализа.');
    } finally {
      setIsRecognizing(false);
    }
  };

  const updateRecognizedPosition = (idx: number, field: string, value: string) => {
    const newPositions = [...recognizedPositions];
    const pos = { ...newPositions[idx] };
    
    if (field === 'rawSize') pos.rawSize = value;
    if (field === 'rawCount') pos.rawCount = parseInt(value) || 0;
    
    // Recalculate quantity
    const size = parseFloat(pos.rawSize);
    const count = pos.rawCount;
    
    if (pos.type === 'profile') {
      pos.quantity = (size * count) / 1000;
    } else if (pos.type === 'sheet') {
      const [l, w] = String(pos.rawSize).split(/[xх*×]/).map(parseFloat);
      if (l && w) {
        const area = (l / 1000) * (w / 1000);
        pos.quantity = area * count * parseFloat(pos.param) * 7.85;
      }
    }
    
    newPositions[idx] = pos;
    setRecognizedPositions(newPositions);
  };

  const removeRecognizedPosition = (idx: number) => {
    setRecognizedPositions(recognizedPositions.filter((_, i) => i !== idx));
  };

  const sendToCalculator = () => {
    if (recognizedPositions.length === 0) return;
    
    // Aggregate positions by subcategory and param
    const aggregated: Record<string, any> = {};
    
    const normParam = (p: string) => p.toLowerCase().replace(/[,]/g, '.').replace(/[х*×]/g, 'x').trim();
    const normSubcat = (s: string) => s.toLowerCase().replace(/[ыиа]$/, '');

    recognizedPositions.forEach(p => {
      const key = `${normSubcat(p.subcategory)}_${normParam(p.param)}`;
      if (aggregated[key]) {
        aggregated[key].quantity += p.quantity;
      } else {
        aggregated[key] = {
          subcategory: p.subcategory,
          param: p.param,
          quantity: p.quantity,
          unit: p.unit,
          name: null
        };
      }
    });

    const newCart = Object.values(aggregated);
    setCart([...cart, ...newCart]);
    setActiveTab('calculator');
    // Scroll to cart
    setTimeout(() => {
      const cartEl = document.getElementById('calculator-cart');
      if (cartEl) cartEl.scrollIntoView({ behavior: 'smooth' });
    }, 100);
  };

  const handleCalculate = async () => {
    let currentCart = [...cart];
    setCalcError(null);

    // Auto-add logic: if cart is empty but form is filled, add it automatically
    if (currentCart.length === 0 && calcSubcategory && calcParameter) {
      const qty = parseFloat(calcQuantity.replace(',', '.'));
      if (!isNaN(qty) && qty > 0) {
        const newEntry = {
          subcategory: calcSubcategory,
          param: calcParameter,
          name: calcName || null,
          quantity: qty,
          unit: 'т'
        };
        currentCart = [newEntry];
      }
    }

    if (currentCart.length === 0) {
      setCalcError("Сначала добавьте позиции в список или заполните поля выше");
      return;
    }
    
    setIsCalculating(true);
    setShowResults(false);
    try {
      console.log("Starting calculation for:", currentCart);
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 30000); // 30s timeout

      const res = await fetch('/api/calculate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ positions: currentCart }),
        signal: controller.signal
      });
      
      clearTimeout(timeoutId);
      console.log("Response status:", res.status);
      
      if (res.ok) {
        const results = await res.json();
        console.log("Calculation results:", results);
        setCalcResults(results);
        setShowResults(true);
        
        setCalcSubcategory('');
        setCalcParameter('');
        setCalcName('');
        setCalcQuantity('');
        setEditIndex(null);
      } else {
        const err = await res.json();
        console.error("Calculation error:", err);
        setCalcError(err.error || "Ошибка расчета");
      }
    } catch (error) {
      console.error("Calculation failed", error);
      setCalcError("Ошибка связи с сервером");
    } finally {
      setIsCalculating(false);
    }
  };

  const getParamLabel = (sub: string) => {
    return "Параметр (толщина, диаметр...)";
  };

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

  const isCategory = (text: string, rowValues: any[], offset: number = 0) => {
    if (!text) return false;
    const val = text.toString().trim();
    if (!val || val.length < 2) return false;
    
    // Проверяем, пусты ли колонки характеристик и цены для этой группы
    // Для левой стороны (offset 0): B(1), C(2), D(3)
    // Для правой стороны (offset 4/5): F(5), G(6), H(7), I(8)
    const otherIndices = offset === 0 ? [1, 2, 3] : [5, 6, 7, 8];
    let hasOtherData = false;
    for (const idx of otherIndices) {
      const cellVal = rowValues[idx];
      if (cellVal !== undefined && cellVal !== null && cellVal.toString().trim() !== '') {
        hasOtherData = true;
        break;
      }
    }

    // Если есть текст в названии, но нет данных в характеристиках и цене - это категория
    if (!hasOtherData) {
      // Дополнительная проверка: заголовок обычно не состоит только из цифр
      if (/^\d+$/.test(val.replace(/[\s.,/-]/g, ''))) return false;
      
      const lower = val.toLowerCase();
      if (lower.includes('цена') || lower.includes('прайс') || lower.includes('наименование')) return false;
      
      return true;
    }

    return false;
  };

  React.useEffect(() => {
    scrollToBottom();
  }, [logs, scrollToBottom]);

  const handleSingleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setSingleExcelFile(file);
      addLog(`Выбран файл для импорта: ${file.name}`, 'info');
    }
  };

  const processSingleExcel = async () => {
    if (!singleExcelFile) return;
    setIsProcessing(true);
    setProgress(0);
    addLog(`Начинаю импорт файла: ${singleExcelFile.name}...`, 'info');

    try {
      const data = await singleExcelFile.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

      if (jsonData.length < 2) throw new Error('Файл пуст или содержит только заголовок');

      addLog(`Файл прочитан. Найдено строк: ${jsonData.length}`, 'info');

      // Simple mapping logic (assuming some common structure or first few columns)
      // We'll try to find columns by header names or use defaults
      const headers = jsonData[0].map((h: any) => String(h).toLowerCase().trim());
      const findCol = (names: string[]) => headers.findIndex(h => names.some(n => h.includes(n)));

      const nameIdx = findCol(['наименование', 'название', 'товар', 'name']) !== -1 ? findCol(['наименование', 'название', 'товар', 'name']) : 0;
      const priceIdx = findCol(['цена', 'стоимость', 'price', 'cost']) !== -1 ? findCol(['цена', 'стоимость', 'price', 'cost']) : 3;
      const param1Idx = findCol(['характеристика', 'параметр', 'размер', 'диаметр', 'толщина']) !== -1 ? findCol(['характеристика', 'параметр', 'размер', 'диаметр', 'толщина']) : 1;
      const catIdx = findCol(['категория', 'раздел', 'category']) !== -1 ? findCol(['категория', 'раздел', 'category']) : -1;

      const items: PriceItem[] = [];
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row[nameIdx]) continue;

        const price = parseFloat(String(row[priceIdx]).replace(/[^\d.]/g, ''));
        if (isNaN(price)) continue;

        items.push({
          name: String(row[nameIdx]),
          param_raw: row[param1Idx] ? String(row[param1Idx]) : '',
          all_params: row[param1Idx] ? String(row[param1Idx]).split(';').map(p => p.trim()).filter(p => p) : [],
          param_name: 'параметр',
          price: price,
          category: catIdx !== -1 ? String(row[catIdx]) : 'Импорт',
          subcategory: 'Ручной ввод',
          supplier: supplierName || 'Manual',
          unit: 'кг'
        });
        
        if (i % 100 === 0) setProgress(Math.round((i / jsonData.length) * 100));
      }

      if (items.length > 0) {
        addLog(`Сохраняю ${items.length} позиций в базу...`, 'info');
        const res = await fetch('/api/save-prices', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ items, supplier: supplierName || 'Manual' })
        });
        if (res.ok) {
          const data = await res.json();
          addLog(`✅ Данные поставщика "${supplierName}" успешно импортированы!`, 'success');
          if (data.subcategories) {
            setAllSubcategories(data.subcategories);
          }
          if (data.items) {
            setAllPriceItems(data.items);
          }
          fetchDiagnostics();
        } else {
          throw new Error('Ошибка при сохранении в базу');
        }
      } else {
        throw new Error('Не удалось извлечь данные из файла. Проверьте структуру.');
      }

    } catch (error) {
      addLog(`❌ Ошибка импорта: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, 'error');
    } finally {
      setIsProcessing(false);
      setProgress(100);
      setSingleExcelFile(null);
    }
  };

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
      
      const finalBlob = await processZip(zipFile, true);
      if (finalBlob) {
        addLog(`🚀 Запускаю скачивание результата...`, "success");
        triggerDownload(finalBlob, options.outputFilename);
      }
      
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
      await processZip(zipFile, true);
      setShowManualInput(false);
    } catch (error) {
      addLog(`❌ Ошибка при скачивании: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, "error");
      setIsProcessing(false);
    }
  };

  const processZip = async (fileToProcess?: File, saveToDb: boolean = false) => {
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
      const allDataForDb: PriceItem[] = [];

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
          
          // Проходим по ВСЕМ листам в файле
          for (let sIdx = 0; sIdx < sourceWb.SheetNames.length; sIdx++) {
            const sheetName = sourceWb.SheetNames[sIdx];
            const worksheet = sourceWb.Sheets[sheetName];
            
            addLog(`  Лист: ${sheetName}`, 'info');
            
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

            // 5. Обрабатываем страницы
            const normalizedRows: any[][] = [];
            let currentCategoryForExcel = 'Без категории';
            let sheetTitle = '';

            const titleCell = worksheet['A3'] || worksheet['B3'] || worksheet['C3'];
            sheetTitle = getCellValue(titleCell) || formattedCategory;

            for (let p = 0; p < pages.length; p++) {
              const page = pages[p];
              let leftCount = 0;
              let rightCount = 0;

              for (let r = 0; r < page.length; r++) {
                const row = page[r];
                const valA = row.values[0] ? row.values[0].toString().trim() : '';
                const valB = row.values[1] ? row.values[1].toString().trim() : '';
                const valC = row.values[2] ? row.values[2].toString().trim() : '';
                const valD = row.values[3] ? row.values[3].toString().trim() : '';
                const valE = row.values[4] ? row.values[4].toString().trim() : '';
                
                // --- ЛОГИКА ДЛЯ НОРМАЛИЗАТОРА (EXCEL) ---
                // Используем только первый лист и оригинальное определение (Колонка A)
                let isExcelSub = false;
                if (sIdx === 0) {
                  if (valA && isCategory(valA, row.values, 0)) {
                    const cleaned = valA.replace(/\s*\(продолжение\)\s*$/i, '').trim();
                    if (cleaned && cleaned.toLowerCase() !== currentCategoryForExcel.toLowerCase()) {
                      currentCategoryForExcel = cleaned;
                      normalizedRows.push([currentCategoryForExcel, '', '', '', 'ПОДКАТЕГОРИЯ']);
                    }
                    isExcelSub = true;
                  }
                }

                if (isExcelSub) continue;
                
                // Левая колонка (A-D)
                if (valA && (valB || valC || valD)) {
                  if (sIdx === 0) {
                    normalizedRows.push([valA, valB, valC, valD, currentCategoryForExcel]);
                    leftCount++;

                    // Сбор данных для JSON (только из тех же строк, что и Excel)
                    const price = parseFloat(valD.toString().replace(/[^\d.]/g, ''));
                    if (!isNaN(price)) {
                      const allParams = valB.includes(';') 
                        ? valB.split(';').map(p => p.trim()).filter(p => p)
                        : (valB ? [valB] : []);
                      
                      const paramName = categoryName.toLowerCase().includes('truby') ? 'стенка' : 
                                       categoryName.toLowerCase().includes('list') ? 'толщина' : 
                                       categoryName.toLowerCase().includes('krug') ? 'диаметр' : 'параметр';

                      allDataForDb.push({
                        name: valA,
                        param_raw: valB,
                        all_params: allParams,
                        param_name: paramName,
                        price: price,
                        category: categoryName,
                        subcategory: currentCategoryForExcel,
                        supplier: 'mc.ru',
                        unit: valC || 'кг'
                      });
                    }
                  }
                }

                // Правая колонка (F-I)
                const valF = row.values[5] ? row.values[5].toString().trim() : '';
                const valG = row.values[6] ? row.values[6].toString().trim() : '';
                const valH = row.values[7] ? row.values[7].toString().trim() : '';
                const valI = row.values[8] ? row.values[8].toString().trim() : '';

                if (valF && (valG || valH || valI)) {
                  if (sIdx === 0 && options.mergeColumns) {
                    normalizedRows.push([valF, valG, valH, valI, currentCategoryForExcel]);
                    rightCount++;

                    // Сбор данных для JSON
                    const price = parseFloat(valI.toString().replace(/[^\d.]/g, ''));
                    if (!isNaN(price)) {
                      const allParams = valG.includes(';') 
                        ? valG.split(';').map(p => p.trim()).filter(p => p)
                        : (valG ? [valG] : []);
                      
                      const paramName = categoryName.toLowerCase().includes('truby') ? 'стенка' : 
                                       categoryName.toLowerCase().includes('list') ? 'толщина' : 
                                       categoryName.toLowerCase().includes('krug') ? 'диаметр' : 'параметр';

                      allDataForDb.push({
                        name: valF,
                        param_raw: valG,
                        all_params: allParams,
                        param_name: paramName,
                        price: price,
                        category: categoryName,
                        subcategory: currentCategoryForExcel,
                        supplier: 'mc.ru',
                        unit: valH || 'кг'
                      });
                    }
                  }
                }
              }
              
              if (sIdx === 0 && (leftCount > 0 || rightCount > 0)) {
                addLog(`    Стр ${p + 1}: добавлено ${leftCount} (лев) + ${rightCount} (прав) записей`, 'info');
              }
            }

            // Добавляем в итоговый Excel ТОЛЬКО первый лист (как в оригинале)
            if (sIdx === 0) {
              const targetSheet = options.splitTabs ? formattedCategory : 'Общий список';
              if (!allDataByCategory[targetSheet]) {
                allDataByCategory[targetSheet] = {
                  title: sheetTitle,
                  data: []
                };
              }
              allDataByCategory[targetSheet].data.push(...normalizedRows);
            }
          }

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
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      setResultBlob(blob);
      addLog('Все файлы обработаны. Итоговый файл готов!', 'success');

      // Пост-обработка: создание JSON из готового Excel
      try {
        addLog('Запуск пост-обработки (создание JSON из Excel)...', 'info');
        const formData = new FormData();
        formData.append('file', blob, 'normalized_prices.xlsx');
        
        const postRes = await fetch('/api/post-process-excel', {
          method: 'POST',
          body: formData
        });
        
        if (postRes.ok) {
          const postData = await postRes.json();
          addLog(`📊 JSON СОЗДАН ИЗ КОПИИ EXCEL`, "success");
          addLog(`📈 Позиций: ${postData.rowCount}`, "info");
          addLog(`💾 Размер JSON: ${postData.sizeKb} KB`, "info");
          if (postData.sample) addLog(`🔍 Пример: ${postData.sample}`, "info");
          
          if (postData.subcategories) {
            setAllSubcategories(postData.subcategories);
          }
          if (postData.items) {
            setAllPriceItems(postData.items);
          }
          
          fetchDiagnostics(); // Обновляем диагностику
        } else {
          const err = await postRes.json();
          addLog(`❌ Ошибка пост-обработки: ${err.error}`, 'error');
        }
      } catch (e) {
        console.error("Post-processing fetch failed", e);
        addLog(`❌ Ошибка связи с сервером при пост-обработке`, 'error');
      }

      return blob;
    } catch (error) {
      addLog(`Критическая ошибка: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}`, 'error');
      return null;
    } finally {
      setIsProcessing(false);
      setProgress(100);
    }
  };

  const triggerDownload = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const downloadFile = () => {
    if (!resultBlob) return;
    triggerDownload(resultBlob, options.outputFilename);
  };

  const removeFromCart = (index: number) => {
    setCart(prev => prev.filter((_, i) => i !== index));
  };

  const updateCartQuantity = (index: number, delta: number) => {
    setCart(prev => prev.map((item, i) => {
      if (i === index) {
        const newQty = Math.max(1, item.quantity + delta);
        return { ...item, quantity: newQty };
      }
      return item;
    }));
  };

  const cartTotal = cart.length; // Just count items for now as prices are unknown until calculation

  return (
    <div className="flex h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans overflow-hidden">
      {/* Sidebar Navigation */}
      <aside className="w-20 lg:w-64 bg-white border-r border-gray-200 flex flex-col z-20">
        <div className="p-6 flex items-center gap-3 border-b border-gray-100">
          <div className="bg-indigo-600 p-2 rounded-xl shadow-lg shadow-indigo-200">
            <FileSpreadsheet className="text-white w-6 h-6" />
          </div>
          <span className="font-bold text-xl tracking-tight hidden lg:block">MetalApp</span>
        </div>

        <nav className="flex-1 p-4 space-y-2">
          <button
            onClick={() => setActiveTab('normalizer')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-4 rounded-2xl transition-all duration-200 group",
              activeTab === 'normalizer' 
                ? "bg-indigo-600 text-white shadow-lg shadow-indigo-200" 
                : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
            )}
          >
            <Layers className={cn("w-6 h-6 transition-transform group-hover:scale-110", activeTab === 'normalizer' ? "text-white" : "text-gray-400")} />
            <span className="font-bold hidden lg:block">Нормализатор</span>
          </button>

          <button
            onClick={() => setActiveTab('calculator')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-4 rounded-2xl transition-all duration-200 group",
              activeTab === 'calculator' 
                ? "bg-emerald-600 text-white shadow-lg shadow-emerald-200" 
                : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
            )}
          >
            <CalcIcon className={cn("w-6 h-6 transition-transform group-hover:scale-110", activeTab === 'calculator' ? "text-white" : "text-gray-400")} />
            <span className="font-bold hidden lg:block">Калькулятор</span>
          </button>

          <button
            onClick={() => setActiveTab('drawing-calc')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-4 rounded-2xl transition-all duration-200 group",
              activeTab === 'drawing-calc' 
                ? "bg-amber-600 text-white shadow-lg shadow-amber-200" 
                : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
            )}
          >
            <ImageIcon className={cn("w-6 h-6 transition-transform group-hover:scale-110", activeTab === 'drawing-calc' ? "text-white" : "text-gray-400")} />
            <span className="font-bold hidden lg:block">Расчёт по чертежу</span>
          </button>
        </nav>

        <div className="p-4 border-t border-gray-100 space-y-2">
          <div className="bg-gray-50 rounded-xl p-3 hidden lg:block">
            <div className="flex items-center gap-2 mb-1">
              <Database size={14} className="text-gray-400" />
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Статус базы</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
              <span className="text-xs font-medium text-gray-600">Подключено</span>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        {activeTab === 'normalizer' && (
          <>
            {/* Normalizer Header */}
            <header className="bg-white border-b border-gray-200 px-8 py-5 flex items-center justify-between sticky top-0 z-10">
              <div>
                <h1 className="text-2xl font-bold tracking-tight text-gray-900">Нормализатор цен</h1>
                <p className="text-sm text-gray-500 mt-1">Обработка прайс-листов и обновление базы данных</p>
              </div>
              
              <div className="flex items-center gap-3">
                <button
                  onClick={downloadPriceFromMcRu}
                  disabled={isProcessing}
                  className={cn(
                    "flex items-center gap-2 px-4 py-2.5 rounded-xl transition-all text-sm font-semibold border",
                    isProcessing
                      ? "bg-gray-50 text-gray-400 border-gray-200 cursor-not-allowed"
                      : "bg-white text-indigo-600 border-indigo-200 hover:bg-indigo-50 hover:border-indigo-300 active:scale-95"
                  )}
                >
                  <Globe size={18} />
                  Загрузить с mc.ru
                </button>
                
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="flex items-center gap-2 px-4 py-2.5 bg-white border border-gray-300 rounded-xl hover:bg-gray-50 transition-all text-sm font-semibold active:scale-95"
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
                    "flex items-center gap-2 px-6 py-2.5 rounded-xl transition-all text-sm font-bold shadow-lg shadow-indigo-100",
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

            <div className="flex-1 overflow-y-auto p-8">
              <div className="max-w-6xl mx-auto grid grid-cols-1 lg:grid-cols-12 gap-8">
                {/* Left Column: Manual Importer */}
                <div className="lg:col-span-4 space-y-6">
                  {/* Manual Importer Section */}
                  <section className="bg-white p-6 rounded-3xl border border-gray-200 shadow-sm">
                    <div className="flex items-center gap-3 mb-6">
                      <div className="bg-emerald-50 p-2 rounded-lg">
                        <Upload size={20} className="text-emerald-600" />
                      </div>
                      <h2 className="font-bold text-lg">Импорт иных файлов</h2>
                    </div>

                    <div className="space-y-4">
                      <div>
                        <label className="block text-xs font-bold text-gray-400 uppercase tracking-wider mb-2">Название поставщика</label>
                        <input
                          type="text"
                          value={supplierName}
                          onChange={e => setSupplierName(e.target.value)}
                          placeholder="Напр: Краски Мира"
                          className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-2xl text-sm font-medium focus:ring-2 focus:ring-emerald-500 outline-none transition-all"
                        />
                      </div>

                      <div 
                        onClick={() => singleFileInputRef.current?.click()}
                        className={cn(
                          "border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all",
                          singleExcelFile 
                            ? "border-emerald-500 bg-emerald-50/30" 
                            : "border-gray-200 hover:border-emerald-300 hover:bg-gray-50"
                        )}
                      >
                        <input
                          type="file"
                          ref={singleFileInputRef}
                          onChange={handleSingleFileSelect}
                          accept=".xlsx,.xls"
                          className="hidden"
                        />
                        {singleExcelFile ? (
                          <div className="space-y-2">
                            <FileSpreadsheet className="mx-auto text-emerald-600" size={32} />
                            <p className="text-sm font-bold text-emerald-900 truncate px-2">{singleExcelFile.name}</p>
                            <p className="text-[10px] text-emerald-600 uppercase font-black">Файл выбран</p>
                          </div>
                        ) : (
                          <div className="space-y-2">
                            <Upload className="mx-auto text-gray-300" size={32} />
                            <p className="text-sm font-bold text-gray-500">Нажмите для выбора Excel</p>
                            <p className="text-[10px] text-gray-400 uppercase font-black">XLSX или XLS</p>
                          </div>
                        )}
                      </div>

                      <button
                        onClick={processSingleExcel}
                        disabled={!singleExcelFile || isProcessing}
                        className={cn(
                          "w-full py-3.5 rounded-2xl font-bold text-sm transition-all shadow-lg active:scale-95",
                          !singleExcelFile || isProcessing
                            ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                            : "bg-emerald-600 text-white hover:bg-emerald-700 shadow-emerald-100"
                        )}
                      >
                        {isProcessing ? <Loader2 className="animate-spin mx-auto" size={20} /> : "Загрузить в базу"}
                      </button>
                    </div>
                  </section>
                </div>

                {/* Right Column: Logs & Progress */}
                <div className="lg:col-span-8 space-y-6">
                  <div className="bg-white rounded-3xl border border-gray-200 shadow-sm flex flex-col overflow-hidden min-h-[500px]">
                    <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/30 cursor-pointer" onClick={() => setIsLogExpanded(!isLogExpanded)}>
                      <div className="flex items-center gap-2">
                        <History size={18} className="text-gray-400" />
                        <h2 className="font-bold text-gray-700">Журнал событий</h2>
                        {isLogExpanded ? <ChevronUp size={16} className="text-gray-400" /> : <ChevronDown size={16} className="text-gray-400" />}
                      </div>
                      <div className="flex items-center gap-4">
                        {logs.length > 0 && (
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              setLogs([]);
                            }}
                            className="text-xs font-bold text-gray-400 hover:text-red-500 flex items-center gap-1.5 transition-colors"
                          >
                            <Trash2 size={14} />
                            Очистить
                          </button>
                        )}
                      </div>
                    </div>

                    <AnimatePresence>
                      {isLogExpanded && (
                        <motion.div 
                          initial={{ height: 0, opacity: 0 }}
                          animate={{ height: 'auto', opacity: 1 }}
                          exit={{ height: 0, opacity: 0 }}
                          className="overflow-hidden"
                        >
                          <div className="max-h-[500px] overflow-y-auto p-6 space-y-3 font-mono text-sm border-t border-gray-50">
                            {logs.length === 0 ? (
                              <div className="h-full flex flex-col items-center justify-center text-gray-400 space-y-4 py-20">
                                <div className="p-6 bg-gray-50 rounded-full">
                                  <Package size={40} className="text-gray-300" />
                                </div>
                                <p className="text-center max-w-xs font-medium">
                                  Ожидание действий... Загрузите файл или выберите из mc.ru
                                </p>
                              </div>
                            ) : (
                              logs.map((log) => (
                                <motion.div
                                  key={log.id}
                                  initial={{ opacity: 0, y: 5 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  className={cn(
                                    "flex gap-3 p-4 rounded-2xl border-l-4",
                                    log.type === 'info' && "bg-blue-50/30 border-blue-400 text-blue-900",
                                    log.type === 'success' && "bg-emerald-50/30 border-emerald-400 text-emerald-900",
                                    log.type === 'error' && "bg-red-50/30 border-red-400 text-red-900",
                                    log.type === 'warning' && "bg-amber-50/30 border-amber-400 text-amber-900"
                                  )}
                                >
                                  <div className="mt-0.5 shrink-0">
                                    {log.type === 'info' && <FileText size={16} className="text-blue-500" />}
                                    {log.type === 'success' && <CheckCircle2 size={16} className="text-emerald-500" />}
                                    {log.type === 'error' && <AlertCircle size={16} className="text-red-500" />}
                                    {log.type === 'warning' && <AlertCircle size={16} className="text-amber-500" />}
                                  </div>
                                  <div className="flex-1">
                                    <div className="flex justify-between items-center mb-1">
                                      <span className="font-bold opacity-40 text-[10px] uppercase tracking-widest">
                                        {log.timestamp.toLocaleTimeString([], { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' })}
                                      </span>
                                    </div>
                                    <p className="leading-relaxed font-medium">{log.message}</p>
                                  </div>
                                </motion.div>
                              ))
                            )}
                            <div ref={logEndRef} />
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>

                    {isProcessing && (
                      <div className="px-6 py-5 bg-white border-t border-gray-100">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-[10px] font-bold text-indigo-600 uppercase tracking-widest">Прогресс обработки</span>
                          <span className="text-sm font-black text-indigo-600">{progress}%</span>
                        </div>
                        <div className="w-full bg-gray-100 rounded-full h-3 overflow-hidden p-0.5">
                          <motion.div 
                            className="bg-indigo-600 h-full rounded-full shadow-sm"
                            initial={{ width: 0 }}
                            animate={{ width: `${progress}%` }}
                            transition={{ type: 'spring', bounce: 0, duration: 0.5 }}
                          />
                        </div>
                      </div>
                    )}

                    {showManualInput && (
                      <motion.div 
                        initial={{ opacity: 0, y: 20 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="px-6 py-6 bg-amber-50/50 border-t border-amber-100 space-y-4"
                      >
                        <div className="flex items-center gap-2 text-amber-800">
                          <AlertCircle size={18} />
                          <p className="text-sm font-bold">Ручной ввод ссылки:</p>
                        </div>
                        <div className="flex gap-3">
                          <input
                            type="text"
                            value={manualUrl}
                            onChange={e => setManualUrl(e.target.value)}
                            placeholder="https://mc.ru/.../metserv.zip"
                            className="flex-1 px-4 py-3 bg-white border border-amber-200 rounded-2xl text-sm outline-none focus:ring-2 focus:ring-amber-500"
                          />
                          <button
                            onClick={downloadFromManualUrl}
                            className="px-6 py-3 bg-amber-600 text-white rounded-2xl text-sm font-bold hover:bg-amber-700 transition-all active:scale-95 shadow-lg shadow-amber-100"
                          >
                            Скачать
                          </button>
                        </div>
                      </motion.div>
                    )}

                    {resultBlob && !isProcessing && (
                      <motion.div 
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        className="px-6 py-6 bg-emerald-50 border-t border-emerald-100 flex items-center justify-between"
                      >
                        <div className="flex items-center gap-4 text-emerald-900">
                          <div className="bg-emerald-500 p-2 rounded-xl">
                            <CheckCircle2 size={24} className="text-white" />
                          </div>
                          <div>
                            <span className="font-black text-lg block">Готово!</span>
                            <span className="text-xs font-medium opacity-70">Файл успешно сформирован и база обновлена</span>
                          </div>
                        </div>
                        <button
                          onClick={downloadFile}
                          className="flex items-center gap-3 px-8 py-4 bg-emerald-600 text-white rounded-2xl hover:bg-emerald-700 transition-all font-bold shadow-xl shadow-emerald-100 active:scale-95"
                        >
                          <Download size={20} />
                          Скачать XLSX
                        </button>
                      </motion.div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </>
        )}
        
        {activeTab === 'calculator' && (
          <div className="flex-1 overflow-y-auto bg-gray-50" id="calculator-cart">
            <div className="max-w-5xl mx-auto p-8 space-y-8">
              {/* Diagnostics Block */}
              <div className="bg-white p-6 rounded-3xl border border-gray-200 shadow-sm">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <div className="bg-indigo-50 p-3 rounded-2xl">
                      <Database size={24} className="text-indigo-600" />
                    </div>
                    <div>
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Прайс от</p>
                      <p className="text-lg font-black text-gray-900">
                        {diagnostics?.last_update || 'Загрузка...'}
                        {diagnostics?.rowCount && (
                          <span className="ml-2 text-xs font-bold text-indigo-400">
                            ({diagnostics.rowCount} поз.)
                          </span>
                        )}
                      </p>
                    </div>
                  </div>
                  
                  <button 
                    onClick={fetchDiagnostics} 
                    className="p-4 bg-gray-50 hover:bg-indigo-50 text-gray-400 hover:text-indigo-600 rounded-2xl transition-all active:scale-90"
                    title="Обновить базу"
                  >
                    <RefreshCw size={24} className={cn(diagError ? "text-red-400" : "")} />
                  </button>
                </div>
              </div>

              {/* Calculator Form */}
              <div className="bg-white p-8 rounded-[40px] border border-gray-200 shadow-sm">
                <div className="flex items-center justify-between mb-8">
                  <h1 className="text-3xl font-black tracking-tight text-gray-900 uppercase italic">КАЛЬКУЛЯТОР</h1>
                  <span className="text-xs font-bold text-gray-400 uppercase tracking-widest">
                    {new Date().toLocaleDateString()}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-12">
                  <div className="space-y-6">
                    <div className="space-y-4">
                      <div>
                        <label className="block text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] mb-2">Подкатегория</label>
                        <div className="relative">
                          <input
                            type="text"
                            value={calcSubcategory}
                            onChange={e => {
                              setCalcSubcategory(e.target.value);
                              setIsSubcategoryDropdownOpen(true);
                            }}
                            onFocus={() => setIsSubcategoryDropdownOpen(true)}
                            placeholder="Выберите из списка..."
                            className={cn(
                              "w-full px-5 py-4 bg-gray-50 border rounded-2xl text-sm font-bold focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all",
                              errors.subcategory ? "border-red-500 bg-red-50" : "border-gray-200"
                            )}
                          />
                          <AnimatePresence>
                            {isSubcategoryDropdownOpen && (
                              <>
                                <div className="fixed inset-0 z-10" onClick={() => setIsSubcategoryDropdownOpen(false)} />
                                <motion.div
                                  initial={{ opacity: 0, y: -10 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  exit={{ opacity: 0, y: -10 }}
                                  className="absolute left-0 right-0 top-full mt-2 bg-white border border-gray-100 rounded-2xl shadow-2xl z-20 max-h-60 overflow-y-auto py-2"
                                >
                                  {allSubcategories
                                    .filter(s => s.toLowerCase().includes(calcSubcategory.toLowerCase()))
                                    .map((s, idx) => (
                                      <button
                                        key={idx}
                                        onClick={() => {
                                          setCalcSubcategory(s);
                                          setIsSubcategoryDropdownOpen(false);
                                        }}
                                        className="w-full text-left px-5 py-3 text-sm font-bold hover:bg-indigo-50 hover:text-indigo-600 transition-colors"
                                      >
                                        {s}
                                      </button>
                                    ))}
                                </motion.div>
                              </>
                            )}
                          </AnimatePresence>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-4">
                        <div>
                          <label className="block text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] mb-2">
                            {getParamLabel(calcSubcategory)}
                          </label>
                          <input
                            type="text"
                            value={calcParameter}
                            onChange={e => setCalcParameter(e.target.value)}
                            placeholder="Напр: 2,5"
                            className={cn(
                              "w-full px-5 py-4 bg-gray-50 border rounded-2xl text-sm font-bold focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all",
                              errors.parameter ? "border-red-500 bg-red-50" : "border-gray-200"
                            )}
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] mb-2">Количество</label>
                          <input
                            type="text"
                            value={calcQuantity}
                            onFocus={(e) => e.target.select()}
                            onChange={e => {
                              let val = e.target.value.replace(',', '.');
                              if (val === '.') val = '0.';
                              if (val === '' || /^\d*\.?\d*$/.test(val)) {
                                setCalcQuantity(val);
                              }
                            }}
                            placeholder="0"
                            className={cn(
                              "w-full px-5 py-4 bg-gray-50 border rounded-2xl text-sm font-bold focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all",
                              errors.quantity ? "border-red-500 bg-red-50" : "border-gray-200"
                            )}
                          />
                        </div>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] mb-2">Наименование (необязательно)</label>
                        <div className="relative">
                          <input
                            type="text"
                            value={calcName}
                            onChange={e => {
                              setCalcName(e.target.value);
                              setIsNameDropdownOpen(true);
                            }}
                            onFocus={() => {
                              setIsNameDropdownOpen(true);
                              if (availableNames.length === 0) {
                                fetchNames();
                              }
                            }}
                            placeholder="Выберите конкретное имя..."
                            className="w-full px-5 py-4 bg-gray-50 border border-gray-200 rounded-2xl text-sm font-bold focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all"
                          />
                          <AnimatePresence>
                            {isNameDropdownOpen && (
                              <>
                                <div className="fixed inset-0 z-10" onClick={() => setIsNameDropdownOpen(false)} />
                                <motion.div
                                  initial={{ opacity: 0, y: -10 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  exit={{ opacity: 0, y: -10 }}
                                  className="absolute left-0 right-0 top-full mt-2 bg-white border border-gray-100 rounded-2xl shadow-2xl z-20 max-h-60 overflow-y-auto py-2"
                                >
                                  <button
                                    onClick={() => {
                                      setCalcName('');
                                      setIsNameDropdownOpen(false);
                                    }}
                                    className="w-full text-left px-5 py-3 text-sm font-bold text-indigo-600 hover:bg-indigo-50 transition-colors"
                                  >
                                    Все подходящие (расчет макс/сред)
                                  </button>
                                  {availableNames.length > 0 ? (
                                    availableNames
                                      .filter(n => n.toLowerCase().includes(calcName.toLowerCase()))
                                      .map((n, idx) => (
                                        <button
                                          key={idx}
                                          onClick={() => {
                                            setCalcName(n);
                                            setIsNameDropdownOpen(false);
                                          }}
                                          className="w-full text-left px-5 py-3 text-sm font-bold hover:bg-indigo-50 hover:text-indigo-600 transition-colors"
                                        >
                                          {n}
                                        </button>
                                      ))
                                  ) : (
                                    <div className="px-5 py-3 text-xs text-gray-400 italic">
                                      {calcSubcategory && calcParameter ? "Загрузка имен..." : "Выберите подкатегорию и параметр"}
                                    </div>
                                  )}
                                </motion.div>
                              </>
                            )}
                          </AnimatePresence>
                        </div>
                      </div>

                      <div className="flex gap-3">
                        <button
                          onClick={handleAddPosition}
                          className="flex-1 py-5 bg-indigo-600 text-white rounded-2xl font-black text-sm uppercase tracking-widest hover:bg-indigo-700 shadow-xl shadow-indigo-100 active:scale-[0.98] transition-all flex items-center justify-center gap-3"
                        >
                          {editIndex !== null ? <CheckCircle2 size={20} strokeWidth={3} /> : <Plus size={20} strokeWidth={3} />}
                          {editIndex !== null ? 'Сохранить' : 'Добавить позицию'}
                        </button>
                        {editIndex !== null && (
                          <button
                            onClick={() => {
                              setEditIndex(null);
                              setCalcSubcategory('');
                              setCalcParameter('');
                              setCalcQuantity('');
                              setCalcName('');
                            }}
                            className="px-6 py-5 bg-gray-100 text-gray-500 rounded-2xl font-black text-sm uppercase tracking-widest hover:bg-gray-200 transition-all active:scale-[0.98]"
                            title="Отменить редактирование"
                          >
                            <X size={20} strokeWidth={3} />
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-col">
                    <div className="flex items-center gap-2 text-gray-400 mb-4">
                      <h2 className="font-black text-sm uppercase tracking-widest">Выбранные позиции</h2>
                      <span className="bg-gray-100 text-gray-500 text-[10px] font-black px-2 py-0.5 rounded-full">{cart.length}</span>
                    </div>

                    <div className="flex-1 bg-gray-50 rounded-3xl border border-gray-100 p-4 space-y-3 overflow-y-auto max-h-[400px]">
                      {calcError && (
                        <motion.div 
                          initial={{ opacity: 0, y: -10 }}
                          animate={{ opacity: 1, y: 0 }}
                          className="p-4 bg-red-50 border border-red-100 rounded-2xl flex items-center gap-3 text-red-600 text-xs font-bold"
                        >
                          <AlertCircle size={16} />
                          {calcError}
                        </motion.div>
                      )}
                      {cart.length === 0 ? (
                        <div className="h-full flex flex-col items-center justify-center text-gray-300 space-y-4 py-10">
                          <Package size={48} strokeWidth={1} />
                          <p className="text-[10px] font-black uppercase tracking-widest">Список пуст</p>
                        </div>
                      ) : (
                        cart.map((entry, idx) => (
                          <div key={idx} className="bg-white p-4 rounded-2xl border border-gray-100 shadow-sm flex items-center justify-between group">
                            <div className="min-w-0 pr-4">
                              <h3 className="font-bold text-gray-900 truncate text-sm">
                                {entry.subcategory} {entry.name && `| ${entry.name}`}
                              </h3>
                              <p className="text-[10px] font-bold text-gray-400 uppercase mt-1">
                                {getParamLabel(entry.subcategory)} {entry.param} | {entry.quantity} {entry.unit}
                              </p>
                            </div>
                            <div className="flex items-center gap-2">
                              <button 
                                onClick={() => handleEditPosition(idx)}
                                className="p-2 hover:bg-indigo-50 text-indigo-500 rounded-lg transition-colors"
                                title="Редактировать"
                              >
                                <Edit2 size={16} />
                              </button>
                              <button 
                                onClick={() => setCart(prev => prev.filter((_, i) => i !== idx))}
                                className="p-2 hover:bg-red-50 text-red-500 rounded-lg transition-colors"
                                title="Удалить"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>

                    <button
                      disabled={isCalculating || (cart.length === 0 && (!calcSubcategory || !calcParameter))}
                      onClick={handleCalculate}
                      className={cn(
                        "mt-6 w-full py-5 rounded-2xl font-black text-sm uppercase tracking-widest transition-all shadow-xl active:scale-[0.98] flex items-center justify-center gap-3",
                        isCalculating || (cart.length === 0 && (!calcSubcategory || !calcParameter))
                          ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                          : "bg-gray-900 text-white hover:bg-black shadow-gray-200"
                      )}
                    >
                      {isCalculating ? (
                        <>
                          <Loader2 size={20} className="animate-spin" />
                          Рассчитываю {cart.length || 1} поз...
                        </>
                      ) : (
                        <>
                          🧮 Рассчитать {cart.length > 0 ? `(${cart.length})` : ''}
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'drawing-calc' && (
          <div className="flex-1 overflow-y-auto bg-gray-50">
            <div className="max-w-5xl mx-auto p-8 space-y-8">
              <div className="bg-white p-8 rounded-[40px] border border-gray-200 shadow-sm">
                <div className="flex items-center justify-between mb-8">
                  <h1 className="text-3xl font-black tracking-tight text-gray-900 uppercase italic">РАСЧЁТ ПО ЧЕРТЕЖУ</h1>
                  <div className="flex items-center gap-3">
                    <button
                      onClick={() => drawingInputRef.current?.click()}
                      className="flex items-center gap-2 px-6 py-3 bg-indigo-600 text-white rounded-2xl hover:bg-indigo-700 transition-all font-bold shadow-lg shadow-indigo-100 active:scale-95"
                    >
                      <Upload size={18} />
                      Загрузить чертёж
                    </button>
                    <input
                      type="file"
                      ref={drawingInputRef}
                      onChange={handleDrawingUpload}
                      accept="image/*"
                      className="hidden"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-12">
                  {/* Drawing Preview */}
                  <div className="space-y-6">
                    <div className="aspect-video bg-gray-100 rounded-3xl border-2 border-dashed border-gray-200 overflow-hidden flex items-center justify-center relative group">
                      {drawingPreview ? (
                        <img src={drawingPreview} alt="Drawing" className="w-full h-full object-contain" />
                      ) : (
                        <div className="text-center space-y-4">
                          <ImageIcon size={48} className="mx-auto text-gray-300" />
                          <p className="text-sm font-bold text-gray-400 uppercase tracking-widest">Чертёж не выбран</p>
                        </div>
                      )}
                    </div>
                    
                    <button
                      disabled={!drawingPreview || isRecognizing}
                      onClick={recognizeDrawing}
                      className={cn(
                        "w-full py-5 rounded-2xl font-black text-sm uppercase tracking-widest transition-all shadow-xl active:scale-[0.98] flex items-center justify-center gap-3",
                        !drawingPreview || isRecognizing
                          ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                          : "bg-gray-900 text-white hover:bg-black shadow-gray-200"
                      )}
                    >
                      {isRecognizing ? (
                        <>
                          <Loader2 size={20} className="animate-spin" />
                          {ocrStatus || 'Распознаю...'}
                        </>
                      ) : (
                        <>
                          <Scan size={20} />
                          Распознать
                        </>
                      )}
                    </button>
                    {ocrStatus && !isRecognizing && (
                      <p className={cn(
                        "text-[10px] font-bold text-center uppercase tracking-widest animate-pulse",
                        ocrError ? "text-red-500" : "text-indigo-600"
                      )}>
                        {ocrError ? `Ошибка: ${ocrError}` : ocrStatus}
                      </p>
                    )}
                  </div>

                  {/* Recognized Positions */}
                  <div className="space-y-6">
                    <div className="flex items-center justify-between text-gray-400">
                      <div className="flex items-center gap-2">
                        <h2 className="font-black text-sm uppercase tracking-widest">Распознанные позиции</h2>
                        <span className="bg-gray-100 text-gray-500 text-[10px] font-black px-2 py-0.5 rounded-full">{recognizedPositions.length}</span>
                      </div>
                      {rawOcrText && (
                        <button 
                          onClick={() => setShowRawText(!showRawText)}
                          className="text-[10px] font-bold hover:text-indigo-600 underline"
                        >
                          {showRawText ? 'Скрыть текст' : 'Показать сырой текст'}
                        </button>
                      )}
                    </div>

                    {showRawText && (
                      <div className="p-4 bg-gray-900 rounded-2xl text-[10px] font-mono text-gray-400 overflow-x-auto whitespace-pre-wrap max-h-40 overflow-y-auto border border-gray-800">
                        <p className="text-gray-500 mb-2 uppercase tracking-widest font-black">Результат OCR:</p>
                        {rawOcrText}
                      </div>
                    )}

                    <div className="bg-gray-50 rounded-3xl border border-gray-100 p-4 space-y-3 overflow-y-auto max-h-[500px] min-h-[300px]">
                      {recognizedPositions.length === 0 ? (
                        <div className="h-full flex flex-col items-center justify-center text-gray-300 space-y-4 py-20">
                          <Scan size={48} strokeWidth={1} />
                          <p className="text-[10px] font-black uppercase tracking-widest text-center max-w-[200px]">
                            Нажмите "Распознать" после загрузки чертежа
                          </p>
                        </div>
                      ) : (
                        recognizedPositions.map((pos, idx) => (
                          <div key={idx} className="bg-white p-5 rounded-2xl border border-gray-100 shadow-sm space-y-4 group relative">
                            <button 
                              onClick={() => removeRecognizedPosition(idx)}
                              className="absolute top-4 right-4 text-gray-300 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"
                            >
                              <Trash2 size={16} />
                            </button>

                            <div className="flex items-center gap-2">
                              <div className={cn(
                                "w-2 h-2 rounded-full",
                                pos.type === 'profile' ? "bg-blue-500" : "bg-emerald-500"
                              )} />
                              <h3 className="font-bold text-gray-900">{pos.subcategory} ({pos.param})</h3>
                            </div>

                            <div className="grid grid-cols-2 gap-4">
                              <div className="space-y-1">
                                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Размер (мм)</label>
                                <input 
                                  type="text"
                                  value={pos.rawSize}
                                  onChange={(e) => updateRecognizedPosition(idx, 'rawSize', e.target.value)}
                                  className="w-full bg-gray-50 border-none rounded-xl px-3 py-2 text-sm font-bold focus:ring-2 focus:ring-indigo-500/20 transition-all"
                                />
                              </div>
                              <div className="space-y-1">
                                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Кол-во (шт)</label>
                                <input 
                                  type="number"
                                  value={pos.rawCount}
                                  onChange={(e) => updateRecognizedPosition(idx, 'rawCount', e.target.value)}
                                  className="w-full bg-gray-50 border-none rounded-xl px-3 py-2 text-sm font-bold focus:ring-2 focus:ring-indigo-500/20 transition-all"
                                />
                              </div>
                            </div>

                            <div className="pt-4 border-t border-gray-50 flex items-center justify-between">
                              <div className="text-[10px] font-bold text-gray-400 uppercase">
                                Итого: <span className="text-indigo-600">{pos.quantity.toFixed(2)} {pos.unit}</span>
                              </div>
                              <div className="text-[10px] font-bold text-gray-400 uppercase">
                                {pos.unit} в прайсе
                              </div>
                            </div>
                          </div>
                        ))
                      )}
                    </div>

                    <button
                      disabled={recognizedPositions.length === 0}
                      onClick={sendToCalculator}
                      className={cn(
                        "w-full py-5 rounded-2xl font-black text-sm uppercase tracking-widest transition-all shadow-xl active:scale-[0.98] flex items-center justify-center gap-3",
                        recognizedPositions.length === 0
                          ? "bg-gray-100 text-gray-400 cursor-not-allowed"
                          : "bg-amber-600 text-white hover:bg-amber-700 shadow-amber-100"
                      )}
                    >
                      <Send size={20} />
                      Отправить в калькулятор
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
        
        {/* Results View - Outside tab conditional */}
        <AnimatePresence>
          {showResults && calcResults && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-white overflow-y-auto">
              <motion.div 
                initial={{ opacity: 0, y: 50 }}
                animate={{ opacity: 1, y: 0 }}
                className="w-full max-w-5xl py-12 space-y-12"
              >
                <div className="flex items-center justify-between">
                  <h2 className="text-4xl font-black uppercase tracking-tighter italic">Результаты расчета</h2>
                  <button 
                    onClick={() => {
                      setShowResults(false);
                      setCart([]); // Clear cart when user returns from results
                    }}
                    className="px-6 py-3 bg-gray-100 hover:bg-gray-200 rounded-2xl text-xs font-black uppercase tracking-widest transition-all"
                  >
                    Вернуться
                  </button>
                </div>

                <div className="grid grid-cols-1 gap-12">
                  {/* Max Table */}
                  <div className="space-y-6">
                    <div className="flex items-center gap-3 text-indigo-600 border-b-2 border-indigo-600 pb-2">
                      <TrendingUp size={24} />
                      <h3 className="text-2xl font-black uppercase tracking-tight">ВАРИАНТ: МАКСИМУМ</h3>
                    </div>
                    <div className="overflow-hidden rounded-3xl border border-gray-100 shadow-sm">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-gray-50">
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Наименование</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Параметр</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Цена</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Кол-во</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Сумма</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {calcResults.results_max.map((item, idx) => (
                            <tr key={idx} className="hover:bg-gray-50/50 transition-colors">
                              <td className="px-6 py-4">
                                <div className="font-bold text-gray-900">{item.name}</div>
                                <div className="text-[10px] font-bold text-gray-400 uppercase">{item.subcategory}</div>
                              </td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.param}</td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.price.toLocaleString()}</td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.quantity} {item.unit}</td>
                              <td className="px-6 py-4 text-right font-black text-indigo-600">{item.total.toLocaleString()}</td>
                            </tr>
                          ))}
                          <tr className="bg-indigo-50/30">
                            <td colSpan={4} className="px-6 py-6 text-right text-sm font-black uppercase tracking-widest text-indigo-900">Итого:</td>
                            <td className="px-6 py-6 text-right text-xl font-black text-indigo-600">
                              {calcResults.results_max.reduce((acc, item) => acc + item.total, 0).toLocaleString()} руб.
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Avg Table */}
                  <div className="space-y-6">
                    <div className="flex items-center gap-3 text-emerald-600 border-b-2 border-emerald-600 pb-2">
                      <TrendingUp size={24} className="rotate-45" />
                      <h3 className="text-2xl font-black uppercase tracking-tight">ВАРИАНТ: СРЕДНЕЕ</h3>
                    </div>
                    <div className="overflow-hidden rounded-3xl border border-gray-100 shadow-sm">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-gray-50">
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Наименование</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Параметр</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Цена</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Кол-во</th>
                            <th className="px-6 py-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Сумма</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {calcResults.results_avg.map((item, idx) => (
                            <tr key={idx} className="hover:bg-gray-50/50 transition-colors">
                              <td className="px-6 py-4">
                                <div className="font-bold text-gray-900">{item.name}</div>
                                <div className="text-[10px] font-bold text-gray-400 uppercase">{item.subcategory}</div>
                              </td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.param}</td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.price.toLocaleString()}</td>
                              <td className="px-6 py-4 text-right font-bold text-gray-900">{item.quantity} {item.unit}</td>
                              <td className="px-6 py-4 text-right font-black text-emerald-600">{item.total.toLocaleString()}</td>
                            </tr>
                          ))}
                          <tr className="bg-emerald-50/30">
                            <td colSpan={4} className="px-6 py-6 text-right text-sm font-black uppercase tracking-widest text-emerald-900">Итого:</td>
                            <td className="px-6 py-6 text-right text-xl font-black text-emerald-600">
                              {calcResults.results_avg.reduce((acc, item) => acc + item.total, 0).toLocaleString()} руб.
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                <div className="flex justify-center pt-8">
                  <button 
                    onClick={() => {
                      const wb = XLSX.utils.book_new();
                      
                      const data = [
                        ['ВАРИАНТ: МАКСИМУМ'],
                        ['Наименование', 'Подкатегория', 'Параметр', 'Цена', 'Кол-во', 'Ед.', 'Сумма'],
                        ...calcResults.results_max.map(item => [
                          item.name, item.subcategory, item.param, item.price, item.quantity, item.unit, item.total
                        ]),
                        ['', '', '', '', '', 'ИТОГО:', calcResults.results_max.reduce((acc, item) => acc + item.total, 0)],
                        [],
                        ['ВАРИАНТ: СРЕДНЕЕ'],
                        ['Наименование', 'Подкатегория', 'Параметр', 'Цена', 'Кол-во', 'Ед.', 'Сумма'],
                        ...calcResults.results_avg.map(item => [
                          item.name, item.subcategory, item.param, item.price, item.quantity, item.unit, item.total
                        ]),
                        ['', '', '', '', '', 'ИТОГО:', calcResults.results_avg.reduce((acc, item) => acc + item.total, 0)]
                      ];

                      const ws = XLSX.utils.aoa_to_sheet(data);

                      // Styles
                      const borderStyle = {
                        top: { style: "thin" },
                        bottom: { style: "thin" },
                        left: { style: "thin" },
                        right: { style: "thin" }
                      };

                      const headerStyle = {
                        fill: { fgColor: { rgb: "E5E7EB" } },
                        font: { bold: true },
                        border: borderStyle,
                        alignment: { horizontal: "center" }
                      };

                      const maxTitleStyle = {
                        font: { bold: true, size: 14, color: { rgb: "1E40AF" } },
                        alignment: { horizontal: "left" }
                      };

                      const avgTitleStyle = {
                        font: { bold: true, size: 14, color: { rgb: "059669" } },
                        alignment: { horizontal: "left" }
                      };

                      const cellStyle = {
                        border: borderStyle
                      };

                      const totalStyle = {
                        font: { bold: true },
                        border: borderStyle,
                        fill: { fgColor: { rgb: "F3F4F6" } }
                      };

                      // Apply styles to cells
                      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
                      const maxTableStart = 0;
                      const maxHeaderRow = 1;
                      const maxDataEnd = 1 + calcResults.results_max.length;
                      const maxTotalRow = 2 + calcResults.results_max.length;
                      
                      const avgTableStart = 4 + calcResults.results_max.length;
                      const avgHeaderRow = 5 + calcResults.results_max.length;
                      const avgDataEnd = 5 + calcResults.results_max.length + calcResults.results_avg.length;
                      const avgTotalRow = 6 + calcResults.results_max.length + calcResults.results_avg.length;

                      for (let R = range.s.r; R <= range.e.r; ++R) {
                        for (let C = range.s.c; C <= range.e.c; ++C) {
                          const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                          if (!ws[cellRef]) continue;

                          if (R === maxTableStart) {
                            ws[cellRef].s = maxTitleStyle;
                          } else if (R === avgTableStart) {
                            ws[cellRef].s = avgTitleStyle;
                          } else if (R === maxHeaderRow || R === avgHeaderRow) {
                            ws[cellRef].s = headerStyle;
                          } else if (R === maxTotalRow || R === avgTotalRow) {
                            ws[cellRef].s = totalStyle;
                          } else if (R > maxHeaderRow && R <= maxDataEnd) {
                            ws[cellRef].s = cellStyle;
                          } else if (R > avgHeaderRow && R <= avgDataEnd) {
                            ws[cellRef].s = cellStyle;
                          }
                        }
                      }

                      ws['!cols'] = [
                        { wch: 40 }, { wch: 25 }, { wch: 15 }, { wch: 12 }, { wch: 10 }, { wch: 8 }, { wch: 15 }
                      ];

                      XLSX.utils.book_append_sheet(wb, ws, 'Смета');
                      XLSX.writeFile(wb, 'metal_estimate.xlsx');
                    }}
                    className="flex items-center gap-3 px-12 py-6 bg-indigo-600 text-white rounded-[32px] font-black text-lg uppercase tracking-widest hover:bg-indigo-700 shadow-2xl shadow-indigo-100 transition-all active:scale-95"
                  >
                    <Download size={24} />
                    💾 Скачать Excel
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
