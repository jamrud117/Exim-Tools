// =========================================
// NAVBAR HIGHLIGHT
// =========================================
const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

// =========================================
// KONFIGURASI KOLUMN
// =========================================
const barangCols = [
  "NO",
  "HS",
  "KODE BARANG",
  "SERI BARANG",
  "URAIAN",
  "KODE SATUAN",
  "JUMLAH SATUAN",
  "NETTO",
  "BRUTO",
  "CIF",
  "CIF RUPIAH",
  "NDPBM",
  "HARGA PENYERAHAN",
];

const ekstraksiCols = [
  "KODE ASAL BB",
  "HS",
  "KODE BARANG",
  "URAIAN",
  "MEREK",
  "TIPE",
  "UKURAN",
  "SPESIFIKASI LAIN",
  "KODE SATUAN",
  "JUMLAH SATUAN",
  "KODE KEMASAN",
  "JUMLAH KEMASAN",
  "KODE DOKUMEN ASAL",
  "KODE KANTOR ASAL",
  "NOMOR DAFTAR ASAL",
  "TANGGAL DAFTAR ASAL",
  "NOMOR AJU ASAL",
  "SERI BARANG ASAL",
  "NETTO",
  "BRUTO",
  "VOLUME",
  "CIF",
  "CIF RUPIAH",
  "NDPBM",
  "HARGA PENYERAHAN",
];

// =========================================
// GLOBAL STATE
// =========================================
let originalEkstrRows = [];
let currentEkstrRows = [];
let seriIndexMap = {}; // 🔑 GLOBAL MAP

// =========================================
// FUNGSI BANTUAN
// =========================================
const sheetToJSON = (sheet) =>
  XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

const buildTable = (headers, rows) => {
  let html = "<table><thead><tr>";
  headers.forEach((h) => (html += `<th>${h}</th>`));
  html += "<th>Aksi</th></tr></thead><tbody>";

  rows.forEach((r, i) => {
    html += "<tr>";
    headers.forEach((_, j) => (html += `<td>${r[j] ?? ""}</td>`));
    html += `<td><button class="copyRowBtn" data-index="${i}">📋 Copy</button></td>`;
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
};

const attachCopyButtons = (id, data) => {
  document.querySelectorAll(`#${id} .copyRowBtn`).forEach((btn) =>
    btn.addEventListener("click", () => {
      navigator.clipboard.writeText(data[btn.dataset.index].join("\t"));
      btn.textContent = "✅ Copied!";
      setTimeout(() => (btn.textContent = "📋 Copy"), 1000);
    })
  );
};

const fadeUpdate = (el, html, after) => {
  el.classList.add("fade-out");
  setTimeout(() => {
    el.innerHTML = html;
    el.classList.remove("fade-out");
    el.classList.add("fade-in");
    if (after) after();
  }, 300);
};

const formatNumber = (v) => {
  const num = parseFloat(v);
  if (!num || isNaN(num)) return ""; // 🔥 0 / NaN / null → kosong
  return Number.isInteger(num) ? num.toString() : num.toFixed(2);
};

// =========================================
// AMBIL NDPBM DARI HEADER
// =========================================
function getNDPBMFromHeader(headerSheet) {
  const json = XLSX.utils.sheet_to_json(headerSheet, { header: 1 });
  if (!json.length) return 1;

  const headerRow = json[0].map((x) =>
    (x || "").toString().trim().toUpperCase()
  );
  const idx = headerRow.indexOf("NDPBM");
  if (idx === -1) return 1;

  const val = parseFloat(json[1]?.[idx]);
  return !val || isNaN(val) ? 1 : val;
}

// =========================================
// PROSES WORKBOOK
// =========================================
function processWorkbook(wb) {
  const headerSheet = wb.Sheets["HEADER"];
  const barangSheet = wb.Sheets["BARANG"];
  const entitasSheet = wb.Sheets["ENTITAS"];

  if (!headerSheet || !barangSheet)
    return Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: "Sheet HEADER atau BARANG tidak ditemukan!",
    });

  // HEADER
  const header = {
    nomorAju: headerSheet["A2"]?.v || "",
    dokumen: headerSheet["B2"]?.v || "",
    kantor: headerSheet["C2"]?.v || "",
    daftar: headerSheet["CP2"]?.v || "",
    tanggal: headerSheet["CF2"]?.v || "",
  };

  // SUPPLIER
  let namaSupplier = "-";
  if (entitasSheet) {
    const ent = XLSX.utils.sheet_to_json(entitasSheet, { header: 1 });
    const hdr = ent[0].map((h) => (h || "").toString().toUpperCase());

    const kodeIdx = hdr.indexOf("KODE ENTITAS");
    const namaIdx = hdr.indexOf("NAMA ENTITAS");

    let target = 3;
    if (header.dokumen == 40) target = 9;
    else if (header.dokumen == 23) target = 5;

    const row = ent.find((r, i) => i > 0 && parseInt(r[kodeIdx]) === target);
    if (row) namaSupplier = row[namaIdx] || "-";
  }

  document.getElementById("headerContent").innerHTML = `
    <table>
      <tr><th>Informasi</th><th>Data</th></tr>
      <tr><td>Nama Supplier</td><td>${namaSupplier}</td></tr>
      <tr><td>Nomor Aju Asal</td><td>${header.nomorAju}</td></tr>
      <tr><td>Kode Dokumen Asal</td><td>${header.dokumen}</td></tr>
      <tr><td>Kode Kantor Asal</td><td>${header.kantor}</td></tr>
      <tr><td>Nomor Daftar Asal</td><td>${header.daftar}</td></tr>
      <tr><td>Tanggal Daftar Asal</td><td>${header.tanggal}</td></tr>
    </table>
  `;

  const ndpbmGlobal = getNDPBMFromHeader(headerSheet);

  // BARANG
  const raw = sheetToJSON(barangSheet);
  const headers = raw[0];
  const data = raw.slice(1);

  const idx = (n) =>
    headers.findIndex((h) => (h || "").toString().toUpperCase() === n);

  const barangRows = data.map((r, i) =>
    barangCols.map((c) => (c === "NO" ? i + 1 : r[idx(c)] ?? ""))
  );

  document.getElementById("barangCard").style.display = "block";
  document.getElementById("barangTableWrap").innerHTML = buildTable(
    barangCols,
    barangRows
  );
  attachCopyButtons("barangTableWrap", barangRows);

  // EKSTRAKSI
  // EKSTRAKSI (FIX LOGIKA DOKUMEN 40)
  // =========================================
  // EKSTRAKSI (FINAL FIX – DOKUMEN 40 MENTAH)
  // =========================================
  const ekstrRows = data.map((r) => {
    const cifExcel = parseFloat(r[idx("CIF")]) || 0;
    const ndpbmExcel = parseFloat(r[idx("NDPBM")]) || "";
    const hargaExcel = parseFloat(r[idx("HARGA PENYERAHAN")]) || 0;
    const cifRpExcel = parseFloat(r[idx("CIF RUPIAH")]) || 0;

    const isDoc40 = header.dokumen == 40;

    return ekstraksiCols.map((c) => {
      if (c === "KODE ASAL BB") return isDoc40 ? 1 : 0;
      if (c === "KODE DOKUMEN ASAL") return header.dokumen;
      if (c === "KODE KANTOR ASAL") return header.kantor;
      if (c === "NOMOR DAFTAR ASAL") return header.daftar;
      if (c === "TANGGAL DAFTAR ASAL") return header.tanggal;
      if (c === "NOMOR AJU ASAL") return header.nomorAju;
      if (c === "SERI BARANG ASAL") return r[idx("SERI BARANG")] ?? "";

      // ===============================
      // 🔥 KHUSUS DOKUMEN 40 (MENTAH)
      // ===============================
      if (isDoc40) {
        if (c === "CIF") return cifExcel ? formatNumber(cifExcel) : "";
        if (c === "CIF RUPIAH")
          return cifRpExcel ? formatNumber(cifRpExcel) : "";
        if (c === "NDPBM") return formatNumber(ndpbmExcel);
        if (c === "HARGA PENYERAHAN") return formatNumber(hargaExcel);
      }

      // ===============================
      // 🔁 DOKUMEN SELAIN 40
      // ===============================
      if (c === "CIF") return formatNumber(cifExcel);
      if (c === "CIF RUPIAH") return formatNumber(cifRpExcel);
      if (c === "NDPBM") return formatNumber(ndpbmGlobal);
      if (c === "HARGA PENYERAHAN") return formatNumber(cifExcel * ndpbmGlobal);

      const i = idx(c);
      return i >= 0 ? r[i] ?? "" : "";
    });
  });

  originalEkstrRows = JSON.parse(JSON.stringify(ekstrRows));
  currentEkstrRows = JSON.parse(JSON.stringify(ekstrRows));

  // 🔢 MAP SERI → INDEX
  seriIndexMap = {};
  currentEkstrRows.forEach((row, i) => {
    const seri = parseInt(row[ekstraksiCols.indexOf("SERI BARANG ASAL")]);
    if (!isNaN(seri)) seriIndexMap[seri] = i;
  });

  const wrap = document.getElementById("ekstraksiTableWrap");
  document.getElementById("ekstraksiCard").style.display = "block";
  fadeUpdate(wrap, buildTable(ekstraksiCols, currentEkstrRows), () =>
    attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
  );

  // DROPDOWN
  const select = document.getElementById("barangSelect");
  select.innerHTML = "";
  select.appendChild(new Option("TAMPILKAN SEMUA", "all"));

  Object.keys(seriIndexMap)
    .map(Number)
    .sort((a, b) => a - b)
    .forEach((seri) =>
      select.appendChild(new Option(`BARANG SERI KE ${seri}`, seri))
    );

  select.onchange = () => {
    if (select.value === "all") {
      fadeUpdate(wrap, buildTable(ekstraksiCols, currentEkstrRows), () =>
        attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
      );
    } else {
      const row = currentEkstrRows[seriIndexMap[select.value]];
      fadeUpdate(wrap, buildTable(ekstraksiCols, [row]), () =>
        attachCopyButtons("ekstraksiTableWrap", [row])
      );
    }
  };
}

// =========================================
// APPLY QUANTITY
// =========================================
function applyQuantity() {
  const qty = parseFloat(document.getElementById("quantityInput").value);
  const select = document.getElementById("barangSelect");

  if (select.value === "all")
    return Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: "Pilih barang tertentu!",
    });

  if (isNaN(qty) || qty <= 0)
    return Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: "Quantity tidak valid!",
    });

  const index = seriIndexMap[select.value];
  const row = [...currentEkstrRows[index]];

  const ci = (n) => ekstraksiCols.indexOf(n);

  const qtyIdx = ci("JUMLAH SATUAN");
  const nettoIdx = ci("NETTO");
  const brutoIdx = ci("BRUTO");
  const cifIdx = ci("CIF");
  const ndpbmIdx = ci("NDPBM");
  const hargaIdx = ci("HARGA PENYERAHAN");
  const docIdx = ci("KODE DOKUMEN ASAL");

  const qtyAwal = parseFloat(row[qtyIdx]) || 1;
  const nettoAwal = parseFloat(row[nettoIdx]) || 0;
  const brutoAwal = parseFloat(row[brutoIdx]) || 0;
  const cifAwal = parseFloat(row[cifIdx]) || 0;
  const ndpbm = parseFloat(row[ndpbmIdx]) || 1;
  const hargaAwal = parseFloat(row[hargaIdx]) || 0;

  const isDoc40 = row[docIdx] == 40;

  // ================================
  // 🔥 DOKUMEN 40 (PROPORSIONAL)
  // ================================
  if (isDoc40) {
    const unitNetto = nettoAwal / qtyAwal;
    const unitBruto = brutoAwal / qtyAwal;
    const unitHarga = hargaAwal / qtyAwal;

    row[qtyIdx] = formatNumber(qty);
    row[nettoIdx] = formatNumber(unitNetto * qty);
    row[brutoIdx] = formatNumber(unitBruto * qty);
    row[hargaIdx] = formatNumber(unitHarga * qty);

    currentEkstrRows[index] = row;

    fadeUpdate(
      document.getElementById("ekstraksiTableWrap"),
      buildTable(ekstraksiCols, [row]),
      () => attachCopyButtons("ekstraksiTableWrap", [row])
    );
    return;
  }

  // ================================
  // 🔁 DOKUMEN SELAIN 40
  // ================================
  const unitNetto = nettoAwal / qtyAwal;
  const unitBruto = brutoAwal / qtyAwal;
  const unitCIF = cifAwal / qtyAwal;

  const nettoBaru = unitNetto * qty;
  const brutoBaru = unitBruto * qty;
  const cifBaru = unitCIF * qty;
  const hargaBaru = cifBaru * ndpbm;

  row[qtyIdx] = formatNumber(qty);
  row[nettoIdx] = formatNumber(nettoBaru);
  row[brutoIdx] = formatNumber(brutoBaru);
  row[cifIdx] = formatNumber(cifBaru);
  row[hargaIdx] = formatNumber(hargaBaru);

  currentEkstrRows[index] = row;

  fadeUpdate(
    document.getElementById("ekstraksiTableWrap"),
    buildTable(ekstraksiCols, [row]),
    () => attachCopyButtons("ekstraksiTableWrap", [row])
  );
}

// =========================================
// RESET
// =========================================
function resetData() {
  currentEkstrRows = JSON.parse(JSON.stringify(originalEkstrRows));
  document.getElementById("quantityInput").value = "";
  document.getElementById("barangSelect").value = "all";

  fadeUpdate(
    document.getElementById("ekstraksiTableWrap"),
    buildTable(ekstraksiCols, currentEkstrRows),
    () => attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
  );
}

// =========================================
// EVENTS
// =========================================
document.getElementById("fileInput").addEventListener("change", (e) => {
  const reader = new FileReader();
  reader.onload = (ev) =>
    processWorkbook(
      XLSX.read(new Uint8Array(ev.target.result), { type: "array" })
    );
  reader.readAsArrayBuffer(e.target.files[0]);
});

document
  .getElementById("applyQuantityBtn")
  .addEventListener("click", applyQuantity);
document.getElementById("resetBtn").addEventListener("click", resetData);
