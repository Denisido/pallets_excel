import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

// ВАЖНО: импортируем ESM-версию библиотеки
import * as XLSX from "xlsx/xlsx.mjs";

// Подключаем fs для Node.js
XLSX.set_fs(fs);

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const INPUT_NAME = "111.xlsx";               // исходный файл в папке files
const OUTPUT_NAME = "111_converted.xlsx";    // новый нормальный .xlsx

const inputPath = path.join(__dirname, "files", INPUT_NAME);
const outputPath = path.join(__dirname, "files", OUTPUT_NAME);

function convert() {
  console.log("📂 Читаю исходный файл:", inputPath);

  // SheetJS сам поймёт формат (xls/xlsx и т.п.)
  const workbook = XLSX.readFile(inputPath);

  console.log("💾 Сохраняю в новый .xlsx:", outputPath);

  // Перезаписываем в чистый формат .xlsx
  XLSX.writeFile(workbook, outputPath, { bookType: "xlsx" });

  console.log("✅ Готово! Новый файл создан:", outputPath);
}

convert();