/**
 * ======================================================
 * PAPERMOB - Konversi Gambar ke Grid Excel (versi JavaScript)
 * ======================================================
 * Hasil konversi dari script Python (PIL + openpyxl) ke Node.js.
 *
 * Library pengganti:
 *  - PIL / Pillow      -> Jimp        (baca gambar, resize, baca pixel)
 *  - openpyxl          -> ExcelJS     (fill warna, font, alignment, freeze panes)
 *
 * Cara install dependency:
 *   npm init -y
 *   npm install jimp@0.22.12 exceljs
 *   (catatan: jimp versi 1.x mengubah API secara total, jadi dipakai versi 0.22.x
 *    yang API-nya cocok dengan kode di bawah ini)
 *
 * Cara menjalankan:
 *   node convert.js
 *   (lalu masukkan path gambar satu per satu seperti versi Python)
 *
 * Atau langsung lewat argumen CLI (tanpa perlu ketik interaktif):
 *   node convert.js gambar1.png gambar2.png gambar3.png
 * ======================================================
 */

const fs = require("fs");
const path = require("path");
const readline = require("readline/promises");
const { stdin: input, stdout: output } = require("process");
const Jimp = require("jimp");
const ExcelJS = require("exceljs");

// ======================================================
// CONFIG
// ======================================================

const WIDTH = 110;
const HEIGHT = 80;

// Sama seperti COLOR_MAP di Python, urutannya dijaga
// agar hasil nearest_color konsisten.
const COLOR_MAP = [
  { rgb: [255, 255, 255], code: "PT", hex: "FFFFFF" }, // Putih
  { rgb: [34, 34, 34], code: "HT", hex: "222222" }, // Hitam
  { rgb: [255, 105, 180], code: "PK", hex: "FF69B4" }, // Pink
  { rgb: [255, 255, 0], code: "KN", hex: "FFFF00" }, // Kuning
  { rgb: [0, 0, 139], code: "BT", hex: "00008B" }, // Biru Tua
  { rgb: [255, 0, 0], code: "MR", hex: "FF0000" }, // Merah
  { rgb: [0, 176, 80], code: "HJ", hex: "00B050" }, // Hijau
  { rgb: [255, 165, 0], code: "JG", hex: "FFA500" }, // Jingga
];

const DEFAULT_ACTION = "B";

// ======================================================
// Cari warna terdekat (sama seperti nearest_color())
// ======================================================

function nearestColor(rgb) {
  let bestDistance = Infinity;
  let best = null;

  const maxVal = Math.max(rgb.r, rgb.g, rgb.b);
  const minVal = Math.min(rgb.r, rgb.g, rgb.b);
  const isNeutral = (maxVal - minVal) < 35;
  const isBlueish = rgb.b > rgb.r;

  // Preservasi outline: jika abu-abu dan agak gelap (r < 195), langsung petakan ke Hitam (HT)
  if (isNeutral && rgb.r < 195) {
    return COLOR_MAP.find(c => c.code === "HT");
  }

  for (const ref of COLOR_MAP) {
    // Jika pixel netral (gray), hanya cocokkan dengan Putih (PT) dan Hitam (HT)
    if (isNeutral && ref.code !== "PT" && ref.code !== "HT") {
      continue;
    }

    // Jika pixel kebiruan (B > R), jangan cocokkan dengan warna kemerahan/kuning (PK, MR, JG, KN)
    if (isBlueish && (ref.code === "PK" || ref.code === "MR" || ref.code === "JG" || ref.code === "KN")) {
      continue;
    }

    // Jika selisih Blue - Green kurang dari 40, pixel tersebut bukan Pink (karena Pink memiliki selisih 75)
    if ((rgb.b - rgb.g) < 40 && ref.code === "PK") {
      continue;
    }

    const [r, g, b] = ref.rgb;
    const distance =
      (rgb.r - r) ** 2 + (rgb.g - g) ** 2 + (rgb.b - b) ** 2;

    if (distance < bestDistance) {
      bestDistance = distance;
      best = ref;
    }
  }

  return best; // { rgb, code, hex }
}

// ======================================================
// Bikin nama sheet yang valid & unik (sama seperti safe_sheet_name())
// ======================================================

function safeSheetName(name, usedNames) {
  // Excel: max 31 karakter, tidak boleh ada karakter : \ / ? * [ ]
  let cleaned = name.replace(/[:\\/?*\[\]]/g, "_");
  if (cleaned.length > 28) cleaned = cleaned.slice(0, 28);
  if (!cleaned) cleaned = "Gambar";

  let final = cleaned;
  let counter = 2;

  while (usedNames.has(final)) {
    const suffix = `_${counter}`;
    final = cleaned.slice(0, 31 - suffix.length) + suffix;
    counter++;
  }

  usedNames.add(final);
  return final;
}

// ======================================================
// Input banyak gambar (sama seperti input_gambar())
// ======================================================

async function inputGambar() {
  const paths = [];
  const rl = readline.createInterface({ input, output });

  console.log("Masukkan path gambar satu per satu.");
  console.log("Ketik 'selesai' (tanpa tanda kutip) kalau sudah tidak ada lagi.\n");

  while (true) {
    const answer = (
      await rl.question(`Path gambar #${paths.length + 1} (atau 'selesai') : `)
    ).trim();

    if (answer.toLowerCase() === "selesai") {
      if (paths.length === 0) {
        console.log("Belum ada gambar yang dimasukkan, masukkan minimal 1 gambar.");
        continue;
      }
      break;
    }

    if (!answer) continue;

    if (!fs.existsSync(answer) || !fs.statSync(answer).isFile()) {
      console.log(`File tidak ditemukan: ${answer}`);
      continue;
    }

    paths.push(answer);
  }

  rl.close();
  return paths;
}

// ======================================================
// Load & validasi gambar, kembalikan grid kode "KODE-AKSI"
// (sama seperti load_gambar_grid())
// ======================================================

async function loadGambarGrid(imagePath) {
  const img = await Jimp.read(imagePath);

  if (img.bitmap.width !== WIDTH || img.bitmap.height !== HEIGHT) {
    console.log(
      `  Ukuran asli ${img.bitmap.width}x${img.bitmap.height} -> di-resize ke ${WIDTH}x${HEIGHT}`
    );
    // Menggunakan bilinear resize default agar detail garis tipis (seperti mata) tidak terpotong
    img.resize(WIDTH, HEIGHT);
  }

  const grid = new Map(); // key: "y,x" -> { value, hex, code }

  for (let y = 0; y < HEIGHT; y++) {
    for (let x = 0; x < WIDTH; x++) {
      const rgba = Jimp.intToRGBA(img.getPixelColor(x, y));
      // Blend dengan background putih jika ada transparansi (alpha channel)
      const alpha = rgba.a / 255;
      const blendedRgb = {
        r: Math.round(rgba.r * alpha + 255 * (1 - alpha)),
        g: Math.round(rgba.g * alpha + 255 * (1 - alpha)),
        b: Math.round(rgba.b * alpha + 255 * (1 - alpha))
      };
      const nearest = nearestColor(blendedRgb);

      const action = nearest.code === "PT" ? "D" : DEFAULT_ACTION;
      const value = `${nearest.code}-${action}`;

      grid.set(`${y},${x}`, { value, hex: nearest.hex, code: nearest.code });
    }
  }

  return grid;
}

// ======================================================
// Sheet visual per gambar (grid berwarna, 1 sheet/gambar)
// (sama seperti buat_sheet_visual())
// ======================================================

function buatSheetVisual(workbook, sheetName, grid) {
  const sheet = workbook.addWorksheet(sheetName);

  for (let c = 1; c <= WIDTH; c++) {
    sheet.getColumn(c).width = 9;
  }

  for (let y = 0; y < HEIGHT; y++) {
    for (let x = 0; x < WIDTH; x++) {
      const { value, hex, code } = grid.get(`${y},${x}`);

      const cell = sheet.getCell(y + 1, x + 1);
      cell.value = value;
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF" + hex },
      };

      const dark = code === "HT" || code === "BT";

      cell.font = {
        color: { argb: dark ? "FFFFFFFF" : "FF000000" },
        bold: true,
        size: 9,
      };

      cell.alignment = { horizontal: "center", vertical: "middle" };
    }

    sheet.getRow(y + 1).height = 24; // tinggi baris (dalam point, sama seperti openpyxl)
  }
}

// ======================================================
// Sheet daftar pixel gabungan (1 baris per pixel, 1 kolom per gambar)
// (sama seperti buat_sheet_daftar())
// ======================================================

function buatSheetDaftar(sheet, frameLabels, allGrids) {
  const header = ["Posisi", "Row", "Col", ...frameLabels];
  sheet.addRow(header);

  sheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  for (let y = 0; y < HEIGHT; y++) {
    for (let x = 0; x < WIDTH; x++) {
      const rowValues = [`${y + 1}.${x + 1}`, y + 1, x + 1];

      for (const grid of allGrids) {
        rowValues.push(grid.get(`${y},${x}`).value);
      }

      sheet.addRow(rowValues);
    }
  }

  sheet.getColumn(1).width = 12;
  sheet.getColumn(2).width = 8;
  sheet.getColumn(3).width = 8;

  for (let i = 0; i < frameLabels.length; i++) {
    sheet.getColumn(4 + i).width = 12;
  }

  sheet.views = [{ state: "frozen", xSplit: 3, ySplit: 1 }]; // setara freeze_panes = "D2"
}

// ======================================================
// MAIN
// ======================================================

async function main() {
  // Jika path gambar dikasih lewat argumen CLI, pakai itu.
  // Kalau tidak, minta input interaktif seperti versi Python.
  const cliArgs = process.argv.slice(2);
  const imagePaths = cliArgs.length > 0 ? cliArgs : await inputGambar();

  const workbook = new ExcelJS.Workbook();

  // Sheet "Daftar Pixel" dibuat lebih dulu (kosong) supaya otomatis
  // menjadi sheet paling depan -- tidak perlu "move sheet" seperti di Python.
  const daftarSheet = workbook.addWorksheet("Daftar Pixel");

  const usedSheetNames = new Set();
  const frameLabels = [];
  const allGrids = [];
  const sheetNames = [];

  for (let i = 0; i < imagePaths.length; i++) {
    const imagePath = imagePaths[i];
    console.log(`Memproses (${i + 1}/${imagePaths.length}): ${imagePath} ...`);

    const grid = await loadGambarGrid(imagePath);

    const baseName = path.basename(imagePath, path.extname(imagePath));
    const sheetName = safeSheetName(baseName, usedSheetNames);

    const frameLabel = `Frame ${i + 1}`;

    frameLabels.push(frameLabel);
    allGrids.push(grid);
    sheetNames.push(sheetName);

    buatSheetVisual(workbook, sheetName, grid);
  }

  buatSheetDaftar(daftarSheet, frameLabels, allGrids);

  const outputFile = "hasil_papermob.xlsx";
  await workbook.xlsx.writeFile(outputFile);

  console.log("=".repeat(40));
  console.log("SELESAI");
  console.log(`Total gambar diproses : ${imagePaths.length}`);
  frameLabels.forEach((label, idx) => {
    console.log(`  ${label} -> sheet '${sheetNames[idx]}'`);
  });
  console.log(outputFile);
  console.log("=".repeat(40));
}

main().catch((err) => {
  console.error("Terjadi error:", err);
  process.exit(1);
});