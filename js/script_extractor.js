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
// KONFIGURASI KOLUMNYA
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

let originalEkstrRows = [];
let currentEkstrRows = [];

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

function formatNumber(value) {
  return Number.isInteger(value) ? value.toString() : value.toFixed(2);
}

// =========================================
// 🔥 FUNGSI BARU: Ambil NDPBM dari HEADER
// =========================================
function getNDPBMFromHeader(headerSheet) {
  const json = XLSX.utils.sheet_to_json(headerSheet, { header: 1 });
  if (!json.length) return 1;

  const headerRow = json[0].map((x) =>
    (x || "").toString().trim().toUpperCase()
  );
  const colIndex = headerRow.indexOf("NDPBM");

  if (colIndex === -1) return 1; // tidak ada kolom NDPBM → default 1

  const raw = json[1]?.[colIndex];
  const num = parseFloat(raw);

  return !num || isNaN(num) ? 1 : num; // jika kosong/0 → default 1
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

  // --- HEADER ASAL ---
  const header = {
    nomorAju: headerSheet["A2"]?.v || "",
    dokumen: headerSheet["B2"]?.v || "",
    kantor: headerSheet["C2"]?.v || "",
    daftar: headerSheet["CP2"]?.v || "",
    tanggal: headerSheet["CF2"]?.v || "",
  };

  // Supplier
  let namaSupplier = "-";

  if (entitasSheet) {
    const ent = XLSX.utils.sheet_to_json(entitasSheet, { header: 1 });
    const hdr = ent[0].map((h) => (h || "").toString().trim().toUpperCase());

    const kodeIdx = hdr.indexOf("KODE ENTITAS");
    const namaIdx = hdr.indexOf("NAMA ENTITAS");

    if (kodeIdx >= 0 && namaIdx >= 0) {
      // Tentukan kode entitas sesuai dokumen asal
      let targetKodeEntitas = 3; // default

      if (header.dokumen === 40 || header.dokumen === "40") {
        targetKodeEntitas = 9;
      } else if (header.dokumen === 27 || header.dokumen === "27") {
        targetKodeEntitas = 3;
      } else if (header.dokumen === 23 || header.dokumen === "23") {
        targetKodeEntitas = 5;
      }

      // Cari baris supplier sesuai kode entitas
      const row = ent.find(
        (r, i) => i > 0 && parseInt(r[kodeIdx]) === targetKodeEntitas
      );

      if (row) namaSupplier = row[namaIdx] || "-";
    }
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

  // =========================================
  // 📌 AMBIL NDPBM DARI SHEET HEADER
  // =========================================
  const ndpbmGlobal = getNDPBMFromHeader(headerSheet);

  // =========================================
  // BARANG EXCEL
  // =========================================
  const raw = sheetToJSON(barangSheet);
  const headers = raw[0];
  const data = raw.slice(1);

  const idx = (n) =>
    headers.findIndex((h) => (h || "").toString().trim().toUpperCase() === n);

  const barangRows = data.map((r, i) =>
    barangCols.map((c) => {
      if (c === "NO") return i + 1;
      return idx(c) >= 0 ? r[idx(c)] ?? "" : "";
    })
  );

  document.getElementById("barangCard").style.display = "block";
  document.getElementById("barangTableWrap").innerHTML = buildTable(
    barangCols,
    barangRows
  );
  attachCopyButtons("barangTableWrap", barangRows);

  // =========================================
  // EKSTRAKSI (Logika CIF Baru + NDPBM HEADER)
  // =========================================

  const ekstrRows = data.map((r) => {
    const cifExcel = parseFloat(r[idx("CIF")]) || 0;
    const hargaExcel = parseFloat(r[idx("HARGA PENYERAHAN")]) || 0;
    const cifRpExcel = parseFloat(r[idx("CIF RUPIAH")]) || 0;

    const cifFinal = cifExcel === 0 ? hargaExcel : cifExcel;
    const hargaFinal = cifExcel === 0 ? hargaExcel : cifFinal * ndpbmGlobal;

    return ekstraksiCols.map((c) => {
      if (c === "KODE ASAL BB") {
        return header.dokumen == 40 ? 1 : 0;
      }
      if (c === "KODE DOKUMEN ASAL") return header.dokumen;
      if (c === "KODE KANTOR ASAL") return header.kantor;
      if (c === "NOMOR DAFTAR ASAL") return header.daftar;
      if (c === "TANGGAL DAFTAR ASAL") return header.tanggal;
      if (c === "NOMOR AJU ASAL") return header.nomorAju;
      if (c === "SERI BARANG ASAL") return r[idx("SERI BARANG")] ?? "";

      if (c === "CIF") return formatNumber(cifFinal);
      if (c === "CIF RUPIAH") return formatNumber(cifRpExcel);
      if (c === "NDPBM") return formatNumber(ndpbmGlobal);
      if (c === "HARGA PENYERAHAN") return formatNumber(hargaFinal);

      const i = idx(c);
      return i >= 0 ? r[i] ?? "" : "";
    });
  });

  originalEkstrRows = JSON.parse(JSON.stringify(ekstrRows));
  currentEkstrRows = JSON.parse(JSON.stringify(ekstrRows));

  const wrap = document.getElementById("ekstraksiTableWrap");
  document.getElementById("ekstraksiCard").style.display = "block";

  fadeUpdate(wrap, buildTable(ekstraksiCols, ekstrRows), () =>
    attachCopyButtons("ekstraksiTableWrap", ekstrRows)
  );

  // =========================================
  // DROPDOWN FILTER BARANG
  // =========================================
  const select = document.getElementById("barangSelect");
  select.innerHTML = "";
  select.appendChild(new Option("TAMPILKAN SEMUA", "all"));
  ekstrRows.forEach((_, i) =>
    select.appendChild(new Option(`BARANG KE ${i + 1}`, i))
  );

  select.addEventListener("change", () => {
    const v = select.value;
    if (v === "all") {
      fadeUpdate(wrap, buildTable(ekstraksiCols, currentEkstrRows), () =>
        attachCopyButtons("ekstraksiTableWrap", currentEkstrRows)
      );
    } else {
      fadeUpdate(wrap, buildTable(ekstraksiCols, [currentEkstrRows[v]]), () =>
        attachCopyButtons("ekstraksiTableWrap", [currentEkstrRows[v]])
      );
    }
  });
}

// =========================================
// APPLY QUANTITY
// =========================================
function applyQuantity() {
  const qty = parseFloat(document.getElementById("quantityInput").value);
  const select = document.getElementById("barangSelect");
  const index = parseInt(select.value);

  if (select.value === "all")
    return Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: "Pilih barang tertentu!",
    });
  if (isNaN(qty))
    return Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: "Masukkan quantity yang valid!",
    });

  const row = [...currentEkstrRows[index]];

  // 🔹 Kolom sesuai ekstraksiCols
  const colIndex = (name) => ekstraksiCols.indexOf(name);

  const qtyIndex = colIndex("JUMLAH SATUAN");
  const cifIndex = colIndex("CIF");
  const cifRpIndex = colIndex("CIF RUPIAH");
  const ndpbmIndex = colIndex("NDPBM");
  const hargaIndex = colIndex("HARGA PENYERAHAN");

  const qtyAwal = parseFloat(row[qtyIndex]) || 1;
  const cifAwal = parseFloat(row[cifIndex]) || 0;
  const cifRupiahAwal = parseFloat(row[cifRpIndex]) || 0;
  const ndpbm = parseFloat(row[ndpbmIndex]) || 1;

  const unitCIF = cifAwal / qtyAwal;
  const unitCIFRp = cifRupiahAwal / qtyAwal;

  const cifBaru = unitCIF * qty;
  const cifRpBaru = unitCIFRp * qty;
  const hargaBaru = cifBaru * ndpbm;

  // Update row dengan nilai baru
  row[qtyIndex] = formatNumber(qty);
  row[cifIndex] = formatNumber(cifBaru);
  row[cifRpIndex] = formatNumber(cifRpBaru);
  row[hargaIndex] = formatNumber(hargaBaru);

  currentEkstrRows[index] = row;

  fadeUpdate(
    document.getElementById("ekstraksiTableWrap"),
    buildTable(ekstraksiCols, [row]),
    () => attachCopyButtons("ekstraksiTableWrap", [row])
  );
}

// =========================================
// RESET BUTTON
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
// FILE INPUT
// =========================================
document.getElementById("fileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (ev) => {
    const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
    processWorkbook(wb);
  };
  reader.readAsArrayBuffer(file);
});

// =========================================
// BUTTON EVENTS
// =========================================
document
  .getElementById("applyQuantityBtn")
  .addEventListener("click", applyQuantity);
document.getElementById("resetBtn").addEventListener("click", resetData);
