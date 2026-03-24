import express from "express";
import { createServer as createViteServer } from "vite";
import * as XLSX from "xlsx";
import AdmZip from "adm-zip";
import fs from "fs";
import path from "path";
import fetch from "node-fetch";
import multer from "multer";
import ExcelJS from "exceljs";

const app = express();
const PORT = 3000;

// Vercel compatibility: use /tmp for storage
const isVercel = !!process.env.VERCEL;
const baseDir = isVercel ? '/tmp' : process.cwd();

const DATA_DIR = path.join(baseDir, "data");
const EXPORTS_DIR = path.join(baseDir, "exports");
const UPLOADS_DIR = path.join(baseDir, "uploads");
const DB_PATH = path.join(DATA_DIR, "prices.json");

// Ensure directories exist
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(EXPORTS_DIR)) fs.mkdirSync(EXPORTS_DIR, { recursive: true });
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

const upload = multer({ dest: UPLOADS_DIR });

// Helper to get the latest JSON file from data directory
const getLatestDbPath = () => {
  if (!fs.existsSync(DATA_DIR)) return DB_PATH;
  const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
  if (files.length === 0) return DB_PATH;
  
  // Sort by mtime
  const latest = files.map(f => ({
    name: f,
    time: fs.statSync(path.join(DATA_DIR, f)).mtime.getTime()
  })).sort((a, b) => b.time - a.time)[0];
  
  return path.join(DATA_DIR, latest.name);
};

interface PriceItem {
  name: string;
  param_raw: string;
  all_params: string[];
  param_name: string;
  price: number;
  category: string;
  subcategory: string;
  supplier: string;
  unit?: string;
}

interface PriceData {
  lastUpdated: string;
  items: PriceItem[];
}

// Helper to get cell value
const getCellValue = (cell: any) => {
  if (!cell) return '';
  return cell.v !== undefined ? cell.v : (cell.w || '');
};

// Helper to check if row is a category
const isCategory = (val: string, row: any[], offset: number = 0) => {
  if (!val) return false;
  const s = val.toString().trim();
  if (s.length < 2) return false;
  
  // If it's all caps and long, likely a header
  if (s.toUpperCase() === s && s.length > 5 && !s.includes('ЦЕНА') && !s.includes('ПРАЙС')) return true;
  
  // Check if price and params columns are empty in this section
  const otherIndices = offset === 0 ? [1, 2, 3] : [5, 6, 7, 8];
  const hasOtherData = otherIndices.some(idx => {
    const cellVal = row[idx];
    return cellVal !== undefined && cellVal !== null && cellVal.toString().trim() !== '';
  });
  
  if (!hasOtherData) {
    // A category header usually doesn't consist only of digits
    if (/^\d+$/.test(s.replace(/[\s.,/-]/g, ''))) return false;
    
    const lower = s.toLowerCase();
    if (lower.includes('цена') || lower.includes('прайс') || lower.includes('наименование')) return false;
    
    return true;
  }
  
  return false;
};

// Normalization logic
async function processZip(buffer: Buffer) {
  const zip = new AdmZip(buffer);
  const zipEntries = zip.getEntries();
  const excelFiles = zipEntries.filter(e => 
    (e.entryName.toLowerCase().endsWith('.xls') || e.entryName.toLowerCase().endsWith('.xlsx')) &&
    !e.entryName.includes('__MACOSX')
  );
  
  const allItems: PriceItem[] = [];

  for (const entry of excelFiles) {
    try {
      const fileName = entry.entryName.split('/').pop() || entry.entryName;
      const workbook = XLSX.read(entry.getData(), { type: 'buffer' });
      // Process ALL sheets
      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        let rows: any[][] = [];
        for (let R = range.s.r; R <= range.e.r; ++R) {
          const row: any[] = [];
          for (let C = 0; C <= 9; C++) {
            row.push(getCellValue(worksheet[XLSX.utils.encode_cell({ c: C, r: R })]));
          }
          rows.push(row);
        }

        // Skip header (usually first 4 rows)
        if (rows.length > 4) rows = rows.slice(4);

        let currentSubcategory = 'Общее';
        const categoryName = fileName.replace(/\.(xls|xlsx)$/i, '');

        for (const row of rows) {
          const valA = row[0] ? row[0].toString().trim() : '';
          const valE = row[4] ? row[4].toString().trim() : '';

          // Subcategory detection - prioritize E as per user request
          // Column A trigger is ignored for the Calculator's database
          if (valE && (valE.toLowerCase().includes('подкатегория') || isCategory(valE, row, 4))) {
            const cleaned = valE.toLowerCase().includes('подкатегория') ? valA : valE;
            if (cleaned) {
              currentSubcategory = cleaned.replace(/\s*\(продолжение\)\s*$/i, '').trim();
            }
            continue;
          }

          // Left column (A-D)
          const valB = row[1] ? row[1].toString().trim() : '';
          const valC = row[2] ? row[2].toString().trim() : '';
          const valD = row[3] ? row[3].toString().trim() : '';

          if (valA && (valB || valC || valD)) {
            const price = parseFloat(valD.toString().replace(/[^\d.]/g, ''));
            if (!isNaN(price) && price > 0) {
              const unitMatch = valA.match(/\s(м|кг|т|шт|м2|м²)\.?$/i) || valB.match(/\s(м|кг|т|шт|м2|м²)\.?$/i);
              const unit = unitMatch ? unitMatch[1].toLowerCase() : 'кг';
              const allParams = valB.split(';').map(p => p.trim()).filter(p => p);
              const paramName = categoryName.toLowerCase().includes('truby') ? 'стенка' : 
                               categoryName.toLowerCase().includes('list') ? 'толщина' : 
                               categoryName.toLowerCase().includes('krug') ? 'диаметр' : 'параметр';

              allItems.push({
                name: valA,
                param_raw: valB,
                all_params: allParams,
                param_name: paramName,
                price: price,
                category: categoryName,
                subcategory: currentSubcategory,
                supplier: 'mc.ru',
                unit: unit
              });
            }
          }

          // Right column (F-I)
          const valF = row[5] ? row[5].toString().trim() : '';
          const valG = row[6] ? row[6].toString().trim() : '';
          const valH = row[7] ? row[7].toString().trim() : '';
          const valI = row[8] ? row[8].toString().trim() : '';

          if (valF && (valG || valH || valI)) {
            const price = parseFloat(valI.toString().replace(/[^\d.]/g, ''));
            if (!isNaN(price) && price > 0) {
              const unitMatch = valF.match(/\s(м|кг|т|шт|м2|м²)\.?$/i) || valG.match(/\s(м|кг|т|шт|м2|м²)\.?$/i);
              const unit = unitMatch ? unitMatch[1].toLowerCase() : 'кг';
              const allParams = valG.split(';').map(p => p.trim()).filter(p => p);
              const paramName = categoryName.toLowerCase().includes('truby') ? 'стенка' : 
                               categoryName.toLowerCase().includes('list') ? 'толщина' : 
                               categoryName.toLowerCase().includes('krug') ? 'диаметр' : 'параметр';

              allItems.push({
                name: valF,
                param_raw: valG,
                all_params: allParams,
                param_name: paramName,
                price: price,
                category: categoryName,
                subcategory: currentSubcategory,
                supplier: 'mc.ru',
                unit: unit
              });
            }
          }
        }
      }
    } catch (err) {
      console.error(`Error processing ${entry.entryName}:`, err);
    }
  }

  const data: PriceData = {
    lastUpdated: new Date().toISOString(),
    items: allItems
  };
  fs.writeFileSync(DB_PATH, JSON.stringify(data));
  return data;
}

app.use(express.json());

// Proxy endpoint to bypass CORS for mc.ru
app.get("/api/proxy", async (req, res) => {
  const targetUrl = req.query.url as string;
  if (!targetUrl) {
    return res.status(400).json({ error: "URL is required" });
  }

  try {
    const response = await fetch(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://mc.ru/prices',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      },
      timeout: 15000, // 15 seconds timeout
      redirect: 'follow'
    });
    
    if (!response.ok) {
      console.error(`Target URL returned ${response.status}: ${response.statusText}`);
      return res.status(response.status).json({ 
        error: `Target returned ${response.status} (${response.statusText}). The site mc.ru might be blocking the server's IP address.` 
      });
    }

    const contentType = response.headers.get("content-type");
    if (contentType) {
      res.setHeader("Content-Type", contentType);
    }

    // Use arrayBuffer and send as buffer for better compatibility with different Node environments
    const arrayBuffer = await response.arrayBuffer();
    res.send(Buffer.from(arrayBuffer));
  } catch (error: any) {
    console.error("Proxy error:", error);
    let message = error.message;
    if (message.includes('ETIMEDOUT')) {
      message = "Connection timed out. The site mc.ru is likely blocking Vercel's IP range. Try downloading the file manually and uploading it.";
    }
    res.status(500).json({ error: `Failed to fetch target URL: ${message}` });
  }
});

app.get("/api/prices", (req, res) => {
  const dbPath = getLatestDbPath();
  if (fs.existsSync(dbPath)) {
    const data = JSON.parse(fs.readFileSync(dbPath, "utf-8"));
    res.json(data);
  } else {
    res.status(404).json({ error: "No data found" });
  }
});

app.post("/api/update-prices", async (req, res) => {
  try {
    const zipPath = path.join(baseDir, "price.zip");
    if (!fs.existsSync(zipPath)) {
      return res.status(400).json({ error: "price.zip not found in root directory" });
    }
    const buffer = fs.readFileSync(zipPath);
    const data = await processZip(buffer);
    res.json(data);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/save-prices", (req, res) => {
  try {
    const { items, supplier } = req.body;
    if (!items || !Array.isArray(items)) {
      return res.status(400).json({ error: "Items array is required" });
    }

    let data: PriceData = { lastUpdated: new Date().toISOString(), items: [] };
    if (fs.existsSync(DB_PATH)) {
      data = JSON.parse(fs.readFileSync(DB_PATH, "utf-8"));
    }

    // Remove old items from this supplier to avoid duplicates
    data.items = data.items.filter(item => item.supplier !== supplier);
    
    // Add new items with supplier tag
    const itemsWithSupplier = items.map(item => ({ ...item, supplier: supplier || 'unknown' }));
    data.items.push(...itemsWithSupplier);
    data.lastUpdated = new Date().toISOString();

    fs.writeFileSync(DB_PATH, JSON.stringify(data));
    
    // Return items and subcategories to help frontend maintain state on Vercel
    const subcategories = Array.from(new Set(data.items.map((i: any) => i.subcategory))).filter(s => s).sort();
    
    res.json({ success: true, count: itemsWithSupplier.length, subcategories, items: data.items });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Diagnostics endpoint
app.get("/api/diagnostics", (req, res) => {
  if (!fs.existsSync(DATA_DIR)) {
    return res.json({ error: "Папка data/ не найдена" });
  }

  const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.json'));
  if (files.length === 0) {
    return res.json({ error: "JSON файлы не найдены. Сначала выполните нормализацию." });
  }

  // Find latest file
  const latestFile = files.map(f => ({
    name: f,
    path: path.join(DATA_DIR, f),
    stat: fs.statSync(path.join(DATA_DIR, f))
  })).sort((a, b) => b.stat.mtime.getTime() - a.stat.mtime.getTime()).pop();

  if (!latestFile) return res.json({ error: "Файл не найден" });

  try {
    const rawData = fs.readFileSync(latestFile.path, "utf-8");
    const data = JSON.parse(rawData);
    
    // Ensure all_params exists
    if (data.items) {
      data.items.forEach((item: any) => {
        if (!item.all_params && item.param_raw) {
          item.all_params = item.param_raw.split(';').map((p: string) => p.trim()).filter((p: string) => p);
        }
      });
    }

    const subcategories = Array.from(new Set(data.items.map((i: any) => i.subcategory))).filter(s => s);
    const subcategoryStats = subcategories.map(sub => ({
      name: sub,
      count: data.items.filter((i: any) => i.subcategory === sub).length
    })).sort((a, b) => b.count - a.count);

    res.json({
      filename: latestFile.name,
      size: (latestFile.stat.size / 1024).toFixed(2) + " KB",
      count: data.items.length,
      samples: data.items.slice(0, 3),
      subcategories: subcategoryStats,
      last_update: data.last_update || data.lastUpdated || "unknown"
    });
  } catch (e: any) {
    res.json({ error: "Ошибка чтения JSON: " + e.message });
  }
});

app.get("/api/get_subcategories", (req, res) => {
  const dbPath = getLatestDbPath();
  if (!fs.existsSync(dbPath)) return res.json([]);
  try {
    const data = JSON.parse(fs.readFileSync(dbPath, "utf-8"));
    const subs = Array.from(new Set(data.items.map((i: any) => i.subcategory.trim()))).filter(s => s).sort();
    res.json(subs);
  } catch (e) {
    res.json([]);
  }
});

app.post("/api/get_names", (req, res) => {
  const { subcategory, param } = req.body;
  if (!subcategory || !param) return res.status(400).json({ error: "Subcategory and param are required" });
  
  const dbPath = getLatestDbPath();
  if (!fs.existsSync(dbPath)) return res.json([]);
  const data = JSON.parse(fs.readFileSync(dbPath, "utf-8"));
  
  // Ensure all_params exists for filtering
  const normalizeParam = (p: string) => String(p).replace(',', '.').trim();
  const targetParam = normalizeParam(param);

  const names = Array.from(new Set(
    data.items
      .filter((i: any) => {
        if (i.subcategory !== subcategory) return false;
        const itemParams = (i.all_params || (i.param_raw ? i.param_raw.split(';').map((p: string) => p.trim()) : []))
          .map(normalizeParam);
        return itemParams.includes(targetParam);
      })
      .map((i: any) => i.name)
  )).sort();
  
  res.json(names);
});

app.post("/api/calculate", (req, res) => {
  const { positions } = req.body;
  if (!positions || !Array.isArray(positions)) return res.status(400).json({ error: "Positions array is required" });

  const dbPath = getLatestDbPath();
  if (!fs.existsSync(dbPath)) return res.status(404).json({ error: "No price data found" });
  const data = JSON.parse(fs.readFileSync(dbPath, "utf-8"));
  const all_items = data.items;

  const results_max: any[] = [];
  const results_avg: any[] = [];

  const normalizeParam = (p: any) => String(p || '').replace(/[,]/g, '.').replace(/[xх*×]/g, 'x').trim();
  const normalizeSubcat = (s: any) => {
    let val = String(s || '').trim().toLowerCase();
    // Remove common endings for better matching (Russian specific: а, ы, и)
    if (val.endsWith('ы') || val.endsWith('и') || val.endsWith('а')) val = val.slice(0, -1);
    return val;
  };

  for (const pos of positions) {
    const targetParam = normalizeParam(pos.param);
    const targetSubcat = normalizeSubcat(pos.subcategory);

    const matches = all_items.filter((item: any) => {
      const itemParams = (item.all_params || (item.param_raw ? item.param_raw.split(';').map((p: string) => p.trim()) : []))
        .map(normalizeParam);
      return normalizeSubcat(item.subcategory) === targetSubcat && itemParams.includes(targetParam);
    });

    if (matches.length === 0) {
      return res.status(400).json({ error: `Нет позиций для ${pos.subcategory} с параметром ${pos.param}` });
    }

    if (pos.name) {
      const exact = matches.filter((m: any) => m.name === pos.name);
      if (exact.length > 0) {
        const price = exact[0].price;
        const item = {
          name: exact[0].name,
          subcategory: pos.subcategory,
          param: pos.param,
          price: price,
          quantity: pos.quantity,
          unit: exact[0].unit || "т",
          total: price * pos.quantity
        };
        results_max.push(item);
        results_avg.push(item);
      } else {
        return res.status(400).json({ error: `Позиция ${pos.name} не найдена` });
      }
    } else if (matches.length === 1) {
      const price = matches[0].price;
      const item = {
        name: matches[0].name,
        subcategory: pos.subcategory,
        param: pos.param,
        price: price,
        quantity: pos.quantity,
        unit: matches[0].unit || "т",
        total: price * pos.quantity
      };
      results_max.push(item);
      results_avg.push(item);
    } else {
      const prices = matches.map((m: any) => m.price);
      const max_price = Math.max(...prices);
      const avg_price = prices.reduce((a: number, b: number) => a + b, 0) / prices.length;

      const closest = matches.reduce((prev: any, curr: any) => 
        Math.abs(curr.price - avg_price) < Math.abs(prev.price - avg_price) ? curr : prev
      );

      const max_item = matches.reduce((prev: any, curr: any) => 
        curr.price > prev.price ? curr : prev
      );

      results_max.push({
        name: max_item.name,
        subcategory: pos.subcategory,
        param: pos.param,
        price: max_price,
        quantity: pos.quantity,
        unit: max_item.unit || "т",
        total: max_price * pos.quantity
      });

      results_avg.push({
        name: closest.name,
        subcategory: pos.subcategory,
        param: pos.param,
        price: closest.price,
        quantity: pos.quantity,
        unit: closest.unit || "т",
        total: closest.price * pos.quantity
      });
    }
  }

  res.json({ results_max, results_avg });
});

app.post("/api/post-process-excel", upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const originalExcel = path.join(EXPORTS_DIR, "normalized_prices.xlsx");
    const jsonPath = path.join(DATA_DIR, "prices.json");

    // Шаг 1. Сохраняем загруженный файл как оригинал
    fs.copyFileSync(req.file.path, originalExcel);
    console.log(`📋 Файл сохранен: ${originalExcel}`);

    // Шаг 2. Читаем Excel с помощью ExcelJS (аналог openpyxl)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(originalExcel);
    
    const allItems = [];
    let rowCount = 0;

    // Шаг 3. Проходим по ВСЕМ вкладкам (листам)
    workbook.eachSheet((worksheet) => {
      console.log(`📖 Обработка вкладки: ${worksheet.name}`);
      
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Пропускаем заголовки

        const name = row.getCell(1).value;        // колонка A
        const paramRaw = row.getCell(2).value;    // колонка B
        const unit = row.getCell(3).value;         // колонка C
        const price = row.getCell(4).value;        // колонка D
        const subcategory = row.getCell(5).value;  // колонка E

        // Проверяем, что это строка с данными (как в Python-коде)
        if (!name || price === null || price === undefined || !subcategory || subcategory === 'ПОДКАТЕГОРИЯ') {
          return;
        }

        // Разбираем параметры
        let allParams: string[] = [];
        if (paramRaw) {
          const sParam = String(paramRaw).trim();
          if (sParam.includes(';')) {
            allParams = sParam.split(';').map(p => p.trim()).filter(p => p);
          } else {
            allParams = [sParam];
          }
        }

        // Цена в число
        let priceValue = 0;
        if (typeof price === 'number') {
          priceValue = price;
        } else if (typeof price === 'object' && price !== null && 'result' in price) {
          // Handle formula result if any
          priceValue = Number(price.result) || 0;
        } else {
          priceValue = parseFloat(String(price)) || 0;
        }

        // Создаём запись
        const item = {
          name: String(name).trim(),
          param_raw: paramRaw ? String(paramRaw).trim() : "",
          all_params: allParams,
          unit: unit ? String(unit).trim() : "т",
          price: priceValue,
          subcategory: String(subcategory).trim()
        };

        allItems.push(item);
        rowCount++;
      });
    });

    // Шаг 4. Сохраняем JSON
    const jsonData = {
      last_update: new Date().toISOString().split('T')[0],
      items: allItems
    };

    fs.writeFileSync(jsonPath, JSON.stringify(jsonData, null, 2));

    // Шаг 5. Удаляем загруженный временный файл
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

    console.log(`\n✅ JSON СОЗДАН ИЗ EXCEL (ExcelJS)`);
    console.log(`📁 Оригинал: ${originalExcel}`);
    console.log(`📦 JSON: ${jsonPath}`);
    console.log(`📈 Позиций: ${rowCount}`);
    
    res.json({
      success: true,
      rowCount,
      jsonPath,
      originalExcel,
      sizeKb: (fs.statSync(jsonPath).size / 1024).toFixed(2),
      sample: allItems.length > 0 ? `${allItems[0].name} (${allItems[0].subcategory})` : null,
      subcategories: Array.from(new Set(allItems.map((i: any) => i.subcategory))).filter(s => s).sort(),
      items: allItems
    });

  } catch (err: any) {
    console.error("Post-processing error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/search", (req, res) => {
  const q = (req.query.q as string || "").toLowerCase();
  if (!fs.existsSync(DB_PATH)) return res.json({ items: [], subcategories: [] });
  
  const data: PriceData = JSON.parse(fs.readFileSync(DB_PATH, "utf-8"));
  const matchedItems = data.items.filter(item => 
    item.name.toLowerCase().includes(q) || 
    item.subcategory.toLowerCase().includes(q) ||
    item.param_raw.toLowerCase().includes(q)
  );

  const subSet = new Set<string>();
  matchedItems.forEach(item => subSet.add(item.subcategory));
  
  // Also add subcategories that match directly
  data.items.forEach(item => {
    if (item.subcategory.toLowerCase().includes(q)) {
      subSet.add(item.subcategory);
    }
  });

  res.json({
    items: matchedItems.slice(0, 50),
    subcategories: Array.from(subSet).slice(0, 30)
  });
});

export { app };

async function startServer() {
  // Only start the server if we're not on Vercel (where it's handled as a serverless function)
  if (isVercel) return;

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
