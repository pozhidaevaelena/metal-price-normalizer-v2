import express from "express";
import { createServer as createViteServer } from "vite";
import * as XLSX from "xlsx";
import AdmZip from "adm-zip";
import fs from "fs";
import path from "path";

const app = express();
const PORT = 3000;
const DATA_DIR = path.join(process.cwd(), "data");
const DB_PATH = path.join(DATA_DIR, "prices.json");

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR);
}

interface PriceItem {
  name: string;
  param1: string;
  param2: string;
  price: number;
  category: string;
  subcategory: string;
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
const isCategory = (val: string, row: any[]) => {
  if (!val) return false;
  const s = val.toString().trim();
  if (s.length < 3) return false;
  
  // If it's all caps and long, likely a header
  if (s.toUpperCase() === s && s.length > 5 && !s.includes('ЦЕНА') && !s.includes('ПРАЙС')) return true;
  
  // If most columns are empty, it's likely a category separator
  const emptyCount = row.filter(v => v === null || v === undefined || v === '').length;
  if (emptyCount >= 3 && s.length > 3) return true;
  
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
      const sheetName = workbook.SheetNames[0];
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
        // Left column (A-D)
        const valA = row[0] ? row[0].toString().trim() : '';
        const valB = row[1] ? row[1].toString().trim() : '';
        const valC = row[2] ? row[2].toString().trim() : '';
        const valD = row[3] ? row[3].toString().trim() : '';

        if (valA && isCategory(valA, row.slice(0, 4))) {
          currentSubcategory = valA.replace(/\s*\(продолжение\)\s*$/i, '').trim();
          continue;
        }

        if (valA && (valB || valC || valD)) {
          const price = parseFloat(valD.toString().replace(/[^\d.]/g, ''));
          if (!isNaN(price) && price > 0) {
            allItems.push({
              name: valA,
              param1: valB,
              param2: valC,
              price: price,
              category: categoryName,
              subcategory: currentSubcategory
            });
          }
        }

        // Right column (F-I)
        const valF = row[5] ? row[5].toString().trim() : '';
        const valG = row[6] ? row[6].toString().trim() : '';
        const valH = row[7] ? row[7].toString().trim() : '';
        const valI = row[8] ? row[8].toString().trim() : '';

        if (valF && isCategory(valF, row.slice(5, 9))) {
          currentSubcategory = valF.replace(/\s*\(продолжение\)\s*$/i, '').trim();
          continue;
        }

        if (valF && (valG || valH || valI)) {
          const price = parseFloat(valI.toString().replace(/[^\d.]/g, ''));
          if (!isNaN(price) && price > 0) {
            allItems.push({
              name: valF,
              param1: valG,
              param2: valH,
              price: price,
              category: categoryName,
              subcategory: currentSubcategory
            });
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
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Referer': 'https://mc.ru/'
      }
    });
    
    if (!response.ok) {
      console.error(`Target URL returned ${response.status}: ${response.statusText}`);
      return res.status(response.status).json({ error: `Target returned ${response.status}` });
    }

    const contentType = response.headers.get("content-type");
    if (contentType) {
      res.setHeader("Content-Type", contentType);
    }

    // Stream the response
    (response.body as any).pipe(res);
  } catch (error) {
    console.error("Proxy error:", error);
    res.status(500).json({ error: "Failed to fetch target URL" });
  }
});

app.get("/api/prices", (req, res) => {
  if (fs.existsSync(DB_PATH)) {
    const data = JSON.parse(fs.readFileSync(DB_PATH, "utf-8"));
    res.json(data);
  } else {
    res.status(404).json({ error: "No data found" });
  }
});

app.post("/api/update-prices", async (req, res) => {
  try {
    const zipPath = path.join(process.cwd(), "price.zip");
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

app.get("/api/search", (req, res) => {
  const q = (req.query.q as string || "").toLowerCase();
  if (!fs.existsSync(DB_PATH)) return res.json({ items: [], subcategories: [] });
  
  const data: PriceData = JSON.parse(fs.readFileSync(DB_PATH, "utf-8"));
  const matchedItems = data.items.filter(item => 
    item.name.toLowerCase().includes(q) || 
    item.subcategory.toLowerCase().includes(q)
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

async function startServer() {
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
