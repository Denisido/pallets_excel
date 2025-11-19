import path from "path";
import fs from "fs";
import * as XLSX from "xlsx/xlsx.mjs";
import { fileURLToPath } from "url";
import { spawnSync } from "child_process";

XLSX.set_fs(fs);

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// —————————————— Вспомогательные функции ——————————————

function divideIfInteger(sum, divisor) {
    if (typeof sum !== "number" || isNaN(sum)) return null;
    const res = sum / divisor;
    return Number.isInteger(res) ? res : null;
}

function normalizeNumber(v) {
    if (v == null) return null;
    if (typeof v === "number") return v;

    return Number(
        v.toString()
            .replace(/\s+/g, "")
            .replace(",", ".")
    );
}

// 💬 POWER‒SHELL диалог выбора файла
function openDialog() {
    const ps = `
Add-Type -AssemblyName System.Windows.Forms;
$fd = New-Object System.Windows.Forms.OpenFileDialog;
$fd.Filter = "Excel Files|*.xlsx;*.xls";
$null = $fd.ShowDialog();
$fd.FileName
`;
    const result = spawnSync("powershell", ["-command", ps], { encoding: "utf8" });
    return result.stdout.trim();
}

// 💬 POWER‒SHELL диалог сохранения файла
function saveDialog(defaultName) {
    const ps = `
Add-Type -AssemblyName System.Windows.Forms;
$sd = New-Object System.Windows.Forms.SaveFileDialog;
$sd.Filter = "Excel Files|*.xlsx";
$sd.FileName = "${defaultName}";
$null = $sd.ShowDialog();
$sd.FileName
`;
    const result = spawnSync("powershell", ["-command", ps], { encoding: "utf8" });
    return result.stdout.trim();
}

// —————————————— ОСНОВНОЙ СКРИПТ ——————————————

async function main() {

    console.log("📁 Выберите Excel файл...");

    const filePath = openDialog();
    if (!filePath) {
        console.log("❌ Файл не выбран");
        return;
    }

    console.log("📄 Исходный файл:", filePath);

    const originalName = path.basename(filePath);
    const baseName = originalName.replace(/\.[^.]+$/, "");
    const defaultSaveName = baseName + "_result.xlsx";

    // Читаем Excel
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const result = [];

    for (let i = 16; i < rows.length; i++) {
        const row = rows[i];

        const valI = row[8];
        const valQ = row[16];
        const valS = row[18];
        const valZ = row[25];

        if (!valI && !valQ && !valS && !valZ) continue;

        const sum = normalizeNumber(valZ);

        const v300 = divideIfInteger(sum, 300);
        const v325 = divideIfInteger(sum, 325);
        const v700 = divideIfInteger(sum, 700);

        if (!v300 && !v325 && !v700) continue;

        result.push({
            Дата: valI,
            ИНН: valQ,
            Наименование: valS,
            Сумма: sum,
            Кол_300: v300,
            Кол_325: v325,
            Кол_700: v700
        });
    }

    // Создаём Excel
    const outWB = XLSX.utils.book_new();
    const outWS = XLSX.utils.json_to_sheet(result);
    XLSX.utils.book_append_sheet(outWB, outWS, "Result");

    console.log("💾 Выберите место сохранения...");

    const savePath = saveDialog(defaultSaveName);

    if (!savePath) {
        console.log("❌ Путь для сохранения не выбран");
        return;
    }

    XLSX.writeFile(outWB, savePath);

    console.log("🎉 Файл успешно сохранён!");
    console.log("📂 Путь:", savePath);
}

main();
