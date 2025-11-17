import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import ExcelJS from "exceljs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Имя файла в папке files
const EXCEL_FILE_NAME = "111.xlsx";

// Путь к файлу Excel
const excelFilePath = path.join(__dirname, "files", EXCEL_FILE_NAME);

// Куда сохраняем JSON
const jsonOutputPath = path.join(__dirname, "result.json");

// Функция для проверки “целое ли деление”
function divideIfInteger(sum, divisor) {
    if (typeof sum !== "number" || isNaN(sum)) return null;

    const result = sum / divisor;

    // Проверяем, что результат целый (остаток == 0)
    if (Number.isInteger(result)) {
        return result;
    }

    return null;
}

// Нормализация значения ячейки (ExcelJS иногда возвращает объекты)
function normalizeCell(value) {
    if (value === null || value === undefined) return null;

    // Если формула, берём result
    if (typeof value === "object") {
        if ("result" in value) return value.result;
        if (value.text) return value.text;
        if (value.richText) {
            return value.richText.map((p) => p.text).join("");
        }
    }

    return value;
}

async function main() {
    try {
        console.log("📂 Загружаю файл:", excelFilePath);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);

        // Берём первый лист
        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
            console.error("❌ Лист в книге не найден!");
            return;
        }

        console.log("📑 Лист:", worksheet.name);

        const data = [];

        // Предполагаем, что первая строка — заголовки, данные с 2-й строки
        const startRow = 2;

        for (let rowNumber = startRow; rowNumber <= worksheet.rowCount; rowNumber++) {
            const row = worksheet.getRow(rowNumber);

            // Колонки: I = 9, Q = 17, S = 19, Z = 26
            const valI = normalizeCell(row.getCell(9).value);   // Дата
            const valQ = normalizeCell(row.getCell(17).value);  // ИНН
            const valS = normalizeCell(row.getCell(19).value);  // Наименование
            const valZ = normalizeCell(row.getCell(26).value);  // Сумма

            // Если все четыре колонки пустые — пропускаем строку
            if (
                (valI === null || valI === "") &&
                (valQ === null || valQ === "") &&
                (valS === null || valS === "") &&
                (valZ === null || valZ === "")
            ) {
                continue;
            }

            const sum = Number(
                typeof valZ === "string"
                    ? valZ.replace(/\s+/g, "").replace(",", ".")
                    : valZ
            );

            const val300 = divideIfInteger(sum, 300);
            const val325 = divideIfInteger(sum, 325);
            const val700 = divideIfInteger(sum, 700);

            data.push({
                Дата: valI ?? null,
                ИНН: valQ ?? null,
                Наименование: valS ?? null,
                Сумма: sum,
                Кол_300: val300,
                Кол_325: val325,
                Кол_700: val700,
                _row: i + 1
            });
        }

        // Записываем в JSON
        fs.writeFileSync(jsonOutputPath, JSON.stringify(data, null, 2), "utf8");

        console.log("✅ Готово! JSON записан в:", jsonOutputPath);
        console.log("🔢 Количество записей:", data.length);
    } catch (err) {
        console.error("❌ Ошибка при обработке Excel:", err);
    }
}

main();