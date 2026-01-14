// ---------- util ----------

function readExcelFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      resolve(workbook);
    };
    reader.readAsArrayBuffer(file);
  });
}

function getCellValue(sheet, cell) {
  const c = sheet[cell];
  return c ? c.v : "";
}

// abis diganti
function getCellValueRC(sheet, r, c) {
  const cell = sheet[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";

  // HANYA ambil cell.w jika cell adalah TEXT/FORMATTED IDENTIFIER
  if (cell.t === "s") return String(cell.v).trim();

  // Number → ambil value asli
  return cell.v ?? "";
}

function getCellTextRC(sheet, r, c) {
  const cell = sheet[XLSX.utils.encode_cell({ r, c })];
  if (!cell) return "";

  // Jika TEXT
  if (cell.t === "s") return String(cell.v).trim();

  // Jika NUMBER → ambil tampilan Excel
  if (cell.t === "n" && cell.w) return String(cell.w).trim();

  return String(cell.v ?? "").trim();
}

// Normalisasi kurs (contoh: "16.460,00" -> 16460)
function parseKurs(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") return val;
  let s = String(val).trim();
  s = s.replace(/\u00A0/g, ""); // non-breaking spaces
  // hapus simbol mata uang & spasi
  s = s.replace(/[^\d,\.\-]/g, "");
  if (s.indexOf(",") > -1 && s.indexOf(".") > -1) {
    // format "16.460,00"
    s = s.replace(/\./g, "").replace(",", ".");
  } else {
    s = s.replace(",", ".");
  }
  const n = parseFloat(s);
  return isNaN(n) ? "" : n;
}

// Format angka (QTY & kemasan integer, lainnya float)
function formatValue(val, isQty = false, unit = "") {
  if (val === null || val === undefined || val === "") return "";

  const str = String(val).trim();
  const match = str.match(/^(-?\d+(\.\d+)?)/);
  if (!match) return str;

  const num = parseFloat(match[1]);
  if (isNaN(num)) return str;

  const rounded = isQty ? Math.round(num) : Math.round(num * 100) / 100;
  const rest = str.substring(match[0].length).trim();
  const suffix = unit || rest;

  return suffix ? `${rounded} ${suffix}` : `${rounded}`;
}

function cleanNumber(val) {
  if (!val) return "";
  return String(val)
    .replace(/.*?:\s*/i, "")
    .trim();
}

function detectFileType(workbook) {
  const names = workbook.SheetNames.map((s) => s.toUpperCase());

  // ====== DATA (Draft EXIM) ======
  if (
    names.includes("HEADER") ||
    names.includes("DOKUMEN") ||
    names.includes("ENTITAS") ||
    names.includes("BARANG")
  ) {
    return "DATA";
  }

  // ====== Cek isi sheet pertama ======
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet || !sheet["!ref"]) return "INV";

  const range = XLSX.utils.decode_range(sheet["!ref"]);

  let foundPacking = false;
  let foundGW = false;
  let foundNW = false;

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || typeof cell.v !== "string") continue;

      const v = cell.v.toUpperCase();

      if (v.includes("PACKING LIST")) foundPacking = true;
      if (v.includes("KEMASAN")) foundPacking = true;
      if (v === "GW" || v.includes("GROSS")) foundGW = true;
      if (v === "NW" || v.includes("NET")) foundNW = true;
    }
  }

  if (foundPacking || (foundGW && foundNW)) {
    return "PL";
  }

  return "INV";
}

// Cari kolom berdasarkan header (tidak diubah)
function findHeaderColumns(sheet, headers) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let found = {};
  let headerRow = null;

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || typeof cell.v !== "string") continue;

      const val = cell.v.toString().trim().toUpperCase();

      for (const key in headers) {
        const target = headers[key];

        // 🔥 JIKA HEADER MAPPING KOSONG → SKIP
        if (!target) continue;

        if (val.includes(String(target).toUpperCase())) {
          found[key] = c;
        }
      }
    }

    if (Object.keys(found).length > 0) {
      headerRow = r;
      break;
    }
  }

  return { ...found, headerRow };
}

// Hitung total dari PL + deteksi satuan kemasan
function hitungKemasanNWGW(sheet) {
  if (!sheet || !sheet["!ref"]) {
    return { kemasanSum: 0, bruttoSum: 0, nettoSum: 0, kemasanUnit: "" };
  }
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let colKemasan = null,
    colGW = null,
    colNW = null,
    headerRow = null;

  // cari kolom & headerRow
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string") {
        const val = cell.v.toString().toUpperCase();
        if (val.includes("KEMASAN")) colKemasan = c;
        if (val.includes("GW")) colGW = c;
        if (val.includes("NW")) colNW = c;
      }
    }
    if (colKemasan !== null && colGW !== null && colNW !== null) {
      headerRow = r;
      break;
    }
  }

  // ==================== AMBIL UNIT KEMASAN ====================
  function detectKemasanUnit(sheet, colKemasan, headerRow, range) {
    // 1️⃣ Dari header: KEMASAN CT
    let headerText = getCellValueRC(sheet, headerRow, colKemasan);
    let m = String(headerText || "").match(/KEMASAN\s*(.*)/i);
    if (m && m[1] && m[1].trim()) return m[1].trim().toUpperCase();

    // 2️⃣ Dari baris setelah header
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const v = getCellValueRC(sheet, r, colKemasan);
      if (v && isNaN(v)) return String(v).trim().toUpperCase();
    }

    // 3️⃣ Dari BARIS TERAKHIR
    for (let r = range.e.r; r >= range.s.r; r--) {
      const v = getCellValueRC(sheet, r, colKemasan);
      if (v && isNaN(v)) return String(v).trim().toUpperCase();
    }

    return "";
  }

  let kemasanUnit = "";
  if (colKemasan !== null && headerRow !== null) {
    kemasanUnit = detectKemasanUnit(sheet, colKemasan, headerRow, range);
  }

  // cari dataStartRow
  let dataStartRow = headerRow !== null ? headerRow + 1 : range.s.r;
  let foundDataStart = false;
  for (let rr = dataStartRow; rr <= range.e.r; rr++) {
    const serial = getCellValueRC(sheet, rr, 0); // kolom A -> c=0
    if (serial !== "" && !isNaN(Number(serial))) {
      dataStartRow = rr;
      foundDataStart = true;
      break;
    }
  }
  if (!foundDataStart) {
    dataStartRow = headerRow !== null ? headerRow + 1 : range.s.r + 1;
  }

  // akumulasi totals dari dataStartRow ke bawah
  let totalKemasan = 0,
    totalGW = 0,
    totalNW = 0;
  if (colKemasan !== null && colGW !== null && colNW !== null) {
    for (let r = dataStartRow; r <= range.e.r; r++) {
      const serial = getCellValueRC(sheet, r, 0);
      if (serial === "" || isNaN(Number(serial))) {
        continue;
      }
      const kemVal = parseInt(getCellValueRC(sheet, r, colKemasan)) || 0;
      const gwVal = parseFloat(getCellValueRC(sheet, r, colGW)) || 0;
      const nwVal = parseFloat(getCellValueRC(sheet, r, colNW)) || 0;

      totalKemasan += kemVal;
      totalGW += gwVal;
      totalNW += nwVal;
    }
  }

  return {
    kemasanSum: totalKemasan,
    bruttoSum: totalGW,
    nettoSum: totalNW,
    kemasanUnit: kemasanUnit,
  };
}

function normalizeQtyUnit(u) {
  if (!u) return "";

  const v = String(u).trim().toUpperCase();

  // PAIRS family → NPR
  if (v === "PAIRS" || v === "PAIR" || v === "PRS" || v === "PR") return "NPR";

  // PCS family → PCE
  if (v === "PCS" || v === "PIECE" || v === "PC" || v === "PCE") return "PCE";

  return v;
}

function getPLUnits(sheetPL) {
  const range = XLSX.utils.decode_range(sheetPL["!ref"]);

  let colQty = null;
  let colUnit = null;
  let headerRow = null;
  let globalUnit = "";

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const raw = getCellValueRC(sheetPL, r, c);
      if (!raw) continue;

      const v = String(raw).toUpperCase();

      // Cari QTY
      if (v.includes("QTY")) {
        colQty = c;

        // 🔥 Ambil unit dari header: TOTAL QTY (PAIRS)
        const m = v.match(/\(([^)]+)\)/);
        if (m) globalUnit = m[1].trim().toUpperCase();
      }

      // Cari kolom UNIT terpisah (kalau ada)
      if (v.includes("SATUAN") || v.includes("UNIT")) {
        colUnit = c;
      }
    }

    if (colQty !== null) {
      headerRow = r;
      break;
    }
  }

  if (colQty === null) return { type: "UNKNOWN", data: [] };

  let items = [];
  let unitSet = new Set();

  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const qty = getCellValueRC(sheetPL, r, colQty);
    if (!qty || isNaN(qty)) continue;

    let unit = "";

    // 1️⃣ Kalau ada kolom UNIT → pakai itu
    if (colUnit !== null) {
      unit = getCellValueRC(sheetPL, r, colUnit);
    }

    // 2️⃣ Kalau tidak ada → pakai global unit dari header
    if (!unit && globalUnit) {
      unit = globalUnit;
    }

    const normUnit = normalizeQtyUnit(unit);

    if (normUnit) unitSet.add(normUnit);

    items.push({
      qty: Number(qty),
      unit: normUnit || null,
    });
  }

  if (unitSet.size > 1) return { type: "PER_ITEM", data: items };
  if (unitSet.size === 1)
    return { type: "GLOBAL", unit: [...unitSet][0], data: items };

  return { type: "UNKNOWN", data: items };
}

// === Ekstraksi data kontrak dari file PL ===
function extractKontrakInfoFromPL(sheetPL) {
  const range = XLSX.utils.decode_range(sheetPL["!ref"]);
  let kontrakNo = "";
  let kontrakTgl = "";

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = sheetPL[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell || typeof cell.v !== "string") continue;

      // Normalisasi & pisah per-baris (multiline cell)
      const lines = cell.v
        .replace(/\r/g, "")
        .split("\n")
        .map((l) => l.trim())
        .filter((l) => l.length > 0);

      // ───────────────────────────────────────────────────
      // 🔥 1) SCAN PER BARIS
      // ───────────────────────────────────────────────────
      for (const line of lines) {
        // No. Kontrak
        if (/No\.?\s*Kontrak/i.test(line)) {
          const m = line.match(/No\.?\s*Kontrak\s*[:\-]?\s*(.*)/i);
          if (m) kontrakNo = m[1].trim();
        }

        // Tanggal Kontrak
        if (/Tanggal\s*Kontrak/i.test(line)) {
          const m = line.match(/Tanggal\s*Kontrak\s*[:\-]?\s*(.*)/i);
          if (m) {
            let raw = m[1].trim();

            const dmatch = raw.match(
              /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/
            );
            if (dmatch) {
              const [_, d, mo, y] = dmatch;
              const yyyy = y.length === 2 ? "20" + y : y;
              raw = `${yyyy}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
            }
            kontrakTgl = raw;
          }
        }
      }

      // ───────────────────────────────────────────────────
      // 🔥 2) EXTRA HANDLER:
      // Jika No Kontrak & Tanggal Kontrak ada dalam satu CELL
      // sejajar seperti:
      // "No. Kontrak : XXX   Tanggal Kontrak : DD-MM-YYYY"
      // ───────────────────────────────────────────────────
      const val = cell.v.replace(/\s+/g, " ").trim();

      if (/No\.?\s*Kontrak/i.test(val) && /Tanggal\s*Kontrak/i.test(val)) {
        // Ambil No Kontrak
        const mNo = val.match(
          /No\.?\s*Kontrak\s*[:\-]?\s*([^:]+?)(?=Tanggal\s*Kontrak|$)/i
        );
        if (mNo) kontrakNo = mNo[1].trim();

        // Ambil Tanggal Kontrak
        const mTgl = val.match(
          /Tanggal\s*Kontrak\s*[:\-]?\s*([A-Za-z0-9\/\-\s]+)/i
        );
        if (mTgl) {
          let raw = mTgl[1].trim();

          const dmatch = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
          if (dmatch) {
            const [_, d, mo, y] = dmatch;
            const yyyy = y.length === 2 ? "20" + y : y;
            raw = `${yyyy}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
          }

          kontrakTgl = raw;
        }
      }
    }
  }

  return { kontrakNo, kontrakTgl };
}
function getNPWPDraft(sheetsDATA) {
  const sheet =
    sheetsDATA.ENTITAS || sheetsDATA.HDR_ENTITAS || sheetsDATA.entitas;

  if (!sheet) return "";

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let colKode = null,
    colIdentitas = null;

  // Cari kolom
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;

    const header = String(cell.v).toUpperCase();
    if (header.includes("KODE ENTITAS")) colKode = c;
    if (header.includes("NOMOR IDENTITAS")) colIdentitas = c;
  }

  if (colKode === null || colIdentitas === null) return "";

  // Cari baris dengan KODE ENTITAS = 8
  for (let r = 1; r <= range.e.r; r++) {
    const kode = getCellValueRC(sheet, r, colKode);
    if (String(kode).trim() === "8") {
      let raw = getCellValueRC(sheet, r, colIdentitas);

      // AUTO-FIX NPWP
      return fixNpwp(raw);
    }
  }

  return "";
}

function fixNpwp(raw) {
  if (!raw) return "";

  // Convert to string
  let s = String(raw).trim();

  // Case 1 — scientific notation (misal 7.698498e+21)
  if (/e\+/i.test(s)) {
    try {
      // Gunakan BigInt untuk menjaga seluruh digit
      const big = BigInt(Number(raw).toFixed(0));
      s = big.toString();
    } catch (e) {
      // fallback
      s = String(Number(raw));
    }
  }

  // Case 2 — bersihkan non-digit
  s = s.replace(/[^0-9]/g, "");

  // Case 3 — jika digit kurang dari 22, tambahkan leading zero
  if (s.length < 22) {
    s = s.padStart(22, "0");
  }

  // Case 4 — jika digit lebih panjang (jarang terjadi), ambil 22 digit terakhir
  if (s.length > 22) {
    s = s.slice(-22);
  }

  return s;
}

function getAddressDraft(sheetsDATA) {
  const sheet = sheetsDATA.ENTITAS;

  if (!sheet) return "";

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let colKode = null,
    colAddress = null;

  // Cari kolom
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;

    const header = String(cell.v).toUpperCase();
    if (header.includes("KODE ENTITAS")) colKode = c;
    if (header.includes("ALAMAT ENTITAS")) colAddress = c;
  }

  if (colKode === null || colAddress === null) return "";

  // Cari baris dengan KODE ENTITAS = 8
  for (let r = 1; r <= range.e.r; r++) {
    const kode = getCellValueRC(sheet, r, colKode);
    if (String(kode).trim() === "8") {
      let raw = getCellValueRC(sheet, r, colAddress);
      return raw;
    }
  }

  return "";
}

function getCustomerDraft(sheetsDATA) {
  const sheet = sheetsDATA.ENTITAS;

  if (!sheet) return "";

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let colKode = null,
    colCustomer = null;

  // Cari kolom
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;

    const header = String(cell.v).toUpperCase();
    if (header.includes("KODE ENTITAS")) colKode = c;
    if (header.includes("NAMA ENTITAS")) colCustomer = c;
  }

  if (colKode === null || colCustomer === null) return "";

  // Cari baris dengan KODE ENTITAS = 8
  for (let r = 1; r <= range.e.r; r++) {
    const kode = getCellValueRC(sheet, r, colKode);
    if (String(kode).trim() === "8") {
      let raw = getCellValueRC(sheet, r, colCustomer);
      return raw;
    }
  }

  return "";
}
