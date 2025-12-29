// ---------- highlight menu aktif di navbar ----------
const currentPage = window.location.pathname.split("/").pop() || "index.html";

document.querySelectorAll(".navbar-nav .nav-link").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

// ---------- utilitas umum ----------
const $ = (id) => document.getElementById(id);
const fmtDate = (d) =>
  `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(
    2,
    "0"
  )}/${d.getFullYear()}`;

const fmtNum = (n) =>
  typeof n === "number"
    ? n.toLocaleString("id-ID")
    : Number(n || 0).toLocaleString("id-ID");

// ---------- pembacaan file ----------
async function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        resolve({ file, wb });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ---------- helper akses cell ----------
function getCell(wb, sheet, addr) {
  const s = wb.Sheets[sheet];
  return s && s[addr] ? s[addr].v : undefined;
}

// ---------- ambil nama entitas ----------
function getEntitas(wb, targetKodeEntitas) {
  const s = wb.Sheets["ENTITAS"];
  if (!s) return "";

  const data = XLSX.utils.sheet_to_json(s, { header: 1 });
  if (!data.length) return "";

  let kodeIdx = -1;
  let namaIdx = -1;

  for (let i = 0; i < data[0].length; i++) {
    const val = String(data[0][i] || "")
      .trim()
      .toUpperCase();
    if (val === "KODE ENTITAS") kodeIdx = i;
    if (val === "NAMA ENTITAS") namaIdx = i;
  }

  for (let r = 1; r < data.length; r++) {
    const kode = String(data[r][kodeIdx] || "").trim();
    if (kode === String(targetKodeEntitas)) {
      return String(data[r][namaIdx] || "").trim();
    }
  }
  return "";
}

// ---------- ekstraksi data utama dari workbook ----------
function extractDataFromWorkbook(file, wb) {
  const jenisBC = document.getElementById("jenisBC").value;
  const kodeEntitas = jenisBC === "Keluar" ? 8 : 3;
  const pengirim = getEntitas(wb, kodeEntitas);
  const bc = getCell(wb, "HEADER", "CP2") || "";
  const segel = getCell(wb, "HEADER", "BC2") || "";

  // ---- KEMASAN ----
  const sheetKemasan = wb.Sheets["KEMASAN"];
  let kemasanMap = {};

  if (sheetKemasan && sheetKemasan["!ref"]) {
    const dataKemasan = XLSX.utils.sheet_to_json(sheetKemasan, { header: 1 });
    const header = dataKemasan[0] || [];
    let kodeIdx = -1,
      jumlahIdx = -1;

    header.forEach((h, i) => {
      const v = String(h || "")
        .trim()
        .toUpperCase();
      if (v === "KODE KEMASAN") kodeIdx = i;
      if (v === "JUMLAH KEMASAN") jumlahIdx = i;
    });

    if (kodeIdx === -1) kodeIdx = 2;
    if (jumlahIdx === -1) jumlahIdx = 3;

    for (let r = 1; r < dataKemasan.length; r++) {
      const kode = String(dataKemasan[r][kodeIdx] || "").trim();
      const qty = Number(dataKemasan[r][jumlahIdx]) || 0;
      if (!kode) continue;
      kemasanMap[kode] = (kemasanMap[kode] || 0) + qty;
    }
  }

  // ---- BARANG ----
  let barangMap = {};
  let barangTotal = 0;
  let barangUnit = "";
  let namaBarang = [];

  const sheetBarang = wb.Sheets["BARANG"];
  if (sheetBarang && sheetBarang["!ref"]) {
    const dataBarang = XLSX.utils.sheet_to_json(sheetBarang, { header: 1 });
    const header = dataBarang[0] || [];

    let uraianIdx = -1;
    let jumlahIdx = -1;
    let satuanIdx = -1;

    header.forEach((h, i) => {
      const v = String(h || "")
        .trim()
        .toUpperCase();
      if (v === "URAIAN") uraianIdx = i;

      if (v === "JUMLAH" || v === "JUMLAH BARANG" || v === "JUMLAH SATUAN")
        jumlahIdx = i;

      if (
        v === "SATUAN" ||
        v === "KODE SATUAN" ||
        v === "SATUAN BARANG" ||
        v === "KODE SATUAN BARANG" ||
        v === "KODE SATUAN JUMLAH"
      )
        satuanIdx = i;
    });

    if (jumlahIdx === -1) jumlahIdx = 9;
    if (satuanIdx === -1) satuanIdx = 10;

    for (let r = 1; r < dataBarang.length; r++) {
      const row = dataBarang[r];
      const qty = Number(row[jumlahIdx]) || 0;
      const unit = String(row[satuanIdx] || "").trim();

      if (qty > 0 && unit) {
        barangMap[unit] = (barangMap[unit] || 0) + qty;
        barangTotal += qty;
        if (!barangUnit) barangUnit = unit;
      }

      if (uraianIdx !== -1 && row[uraianIdx]) {
        namaBarang.push(String(row[uraianIdx]).trim());
      }
    }
  }

  const t = getCell(wb, "HEADER", "CF2");
  const aju = getCell(wb, "HEADER", "A2") || "";
  const n2Val = getCell(wb, "HEADER", "N2");

  const jenisMap = {
    1: "PENYERAHAN BKP",
    2: "PENYERAHAN JKP",
    3: "RETUR",
    4: "NON PENYERAHAN",
    5: "LAINNYA",
  };

  return {
    jenistrx: jenisMap[String(n2Val).trim()] || "TIDAK DIKETAHUI",
    aju,
    pengirim,
    bc,
    segel,
    kemasan: kemasanMap,
    barang: {
      map: barangMap,
      total: barangTotal,
      unit: barangUnit,
    },
    tanggal: t ? new Date(t) : null,
    namaBarang: [...new Set(namaBarang)],
  };
}

// ---------- format tanggal dokumen ----------
function formatTanggalDokumen(arr) {
  if (!arr.length) return "";
  const sorted = [...new Set(arr.map((t) => t.getTime()))]
    .map((t) => new Date(t))
    .sort((a, b) => a - b);

  const groups = [];
  let start = sorted[0];
  let end = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    if ((cur - end) / 86400000 === 1) end = cur;
    else {
      groups.push([start, end]);
      start = end = cur;
    }
  }
  groups.push([start, end]);

  return groups
    .map(([s, e]) =>
      s.getTime() === e.getTime()
        ? fmtDate(s)
        : `${String(s.getDate()).padStart(2, "0")}-${String(
            e.getDate()
          ).padStart(2, "0")}/${String(s.getMonth() + 1).padStart(
            2,
            "0"
          )}/${s.getFullYear()}`
    )
    .join(", ");
}

// ---------- tampilan UI ----------
function renderPreview(dataArr) {
  const tbody = $("previewTableBody");
  tbody.innerHTML = dataArr
    .map(
      (d, idx) => `
      <tr>
        <td>${idx + 1}</td>
        <td>${d.jenistrx}</td>
        <td>${d.aju}</td>
        <td>${d.pengirim}</td>
        <td>${d.bc || ""}</td>
        <td>${d.segel || ""}</td>
        <td>${Object.entries(d.kemasan)
          .map(([u, q]) => `${fmtNum(q)} ${u}`)
          .join("<br>")}</td>
        <td>${Object.entries(d.barang.map)
          .map(([u, q]) => `${fmtNum(q)} ${u}`)
          .join("<br>")}</td>
        <td>${d.tanggal ? fmtDate(d.tanggal) : ""}</td>
        <td>${d.namaBarang.join("<br>") || "-"}</td>
      </tr>`
    )
    .join("");

  $("tableWrap").classList.remove("d-none");
}
// Helper
function getSelectedValues(selectId) {
  return Array.from($(selectId).selectedOptions).map((o) => o.value);
}
function formatKeyValue(map) {
  return Object.entries(map)
    .map(([u, q]) => `${fmtNum(q)} ${u}`)
    .join(" + ");
}
function parseJalurOverride(text) {
  const map = {};
  text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .forEach((line) => {
      const [jalur, list] = line.split("=");
      if (!jalur || !list) return;

      list.split(",").forEach((no) => {
        map[no.trim()] = jalur.trim().toUpperCase();
      });
    });
  return map;
}

// ---------- generate result text ----------
function generateResultText(dataArr) {
  const jenisBC = $("jenisBC").value;

  // Ambil nilai dropdown sebagai default jalur
  const defaultJalur = $("statusJalur")?.value || "HIJAU";

  // Ambil override per BC / No Aju dari textarea
  const jalurOverrideMap = parseJalurOverride($("jalurOverride")?.value || "");

  // Ambil jenis barang yang dipilih
  const jenisBarangArr = getSelectedValues("jenisBarang");
  const jenisBarang = jenisBarangArr.join(" + ");
  const masukTxt = fmtDate(new Date($("masukTgl").value));

  // Ambil list pengirim unik
  const pengirim = [
    ...new Set(dataArr.map((d) => d.pengirim).filter(Boolean)),
  ].join(" | ");

  const bcList = {};
  const segelList = [];
  const kemasanMap = {};
  const barangMap = {};
  const tanggalArr = [];
  const bcGrouped = {};

  // ===============================
  // Proses tiap data
  // ===============================
  dataArr.forEach((d) => {
    // Jalur = override per BC/No Aju, kalau tidak ada pakai dropdown
    const jalur =
      jalurOverrideMap[d.bc] || jalurOverrideMap[d.aju] || defaultJalur;

    // Key gabungan Jalur + Jenis Transaksi
    const key = `${jalur} | ${d.jenistrx}`;
    if (!bcGrouped[key]) bcGrouped[key] = [];
    if (d.bc) bcGrouped[key].push(d.bc);

    if (!bcList[d.jenistrx]) bcList[d.jenistrx] = [];
    if (d.bc) bcList[d.jenistrx].push(d.bc);
    if (d.segel) segelList.push(d.segel);

    for (const [u, q] of Object.entries(d.kemasan))
      kemasanMap[u] = (kemasanMap[u] || 0) + q;

    for (const [u, q] of Object.entries(d.barang.map))
      barangMap[u] = (barangMap[u] || 0) + q;

    if (d.tanggal) tanggalArr.push(d.tanggal);
  });

  // ================================
  // FORMAT KHUSUS 2.7 KELUAR
  // ================================
  if (jenisBC === "Keluar") {
    const totalDokumen = dataArr.length;
    const labelTanggal = "Tanggal Keluar";

    // Urutan jalur: Hijau → Merah → Kuning
    const jalurOrder = { HIJAU: 1, MERAH: 2, KUNING: 3 };
    const sortedKeys = Object.keys(bcGrouped).sort((a, b) => {
      const jalurA = a.split("|")[0].trim().toUpperCase();
      const jalurB = b.split("|")[0].trim().toUpperCase();
      return (jalurOrder[jalurA] || 99) - (jalurOrder[jalurB] || 99);
    });

    const bcLines = sortedKeys
      .map((key) => `BC 2.7 (${key}) : ${bcGrouped[key].join(", ")}`)
      .join("\n");

    return [
      `*BC 2.7 Keluar*`,
      `Customer : ${pengirim}`,
      bcLines,
      `No. Segel : ${segelList.join(", ") || ""}`,
      `Jumlah Dokumen : ${totalDokumen} Dokumen`,
      `Jenis Barang : ${jenisBarang}`,
      `Jumlah Barang : ${formatKeyValue(barangMap)}`,
      `Kemasan : ${formatKeyValue(kemasanMap)}`,
      `${labelTanggal} : ${masukTxt}`,
    ].join("\n");
  }

  // ================================
  // FORMAT KHUSUS BC 2.7 MASUK
  // ================================
  const labelTanggal = "Tanggal Masuk";
  const totalDocs = dataArr.length;

  return [
    `*BC 2.7 ${jenisBC}*`,
    `Supplier : ${pengirim}`,
    ...Object.entries(bcList).map(
      ([j, l]) => `No BC 2.7 ( ${j} ) : ${l.join(", ")}`
    ),
    `No Segel : ${segelList.join(", ")}`,
    `Jumlah Dokumen : ${totalDocs} Dokumen`,
    `Jenis Barang : ${jenisBarang}`,
    `Jumlah barang : ${formatKeyValue(barangMap)}`,
    `Jumlah kemasan : ${formatKeyValue(kemasanMap)}`,
    `Tanggal Dokumen : ${formatTanggalDokumen(tanggalArr)}`,
    `${labelTanggal} : ${masukTxt}`,
  ].join("\n");
}

// ---------- event handler ----------
document.getElementById("masukTgl").addEventListener("click", function () {
  this.showPicker();
});

$("masukTgl").value = new Date().toISOString().slice(0, 10);
let selectedFiles = [];

$("files").addEventListener("change", (e) => {
  selectedFiles = Array.from(e.target.files);
  $("fileList").textContent = selectedFiles.length
    ? selectedFiles.map((f) => f.name).join(", ")
    : "Belum ada file dipilih.";
});

$("processBtn").addEventListener("click", async () => {
  if (!selectedFiles.length)
    return Swal.fire({ icon: "error", text: "Pilih minimal 1 file Excel!" });

  if (!getSelectedValues("jenisBarang").length)
    return Swal.fire({
      icon: "error",
      text: "Pilih minimal 1 jenis barang!",
    });

  $("processBtn").disabled = true;
  $("processBtn").textContent = "Memproses...";

  try {
    const workbooks = await Promise.all(selectedFiles.map(readWorkbook));
    const extracted = workbooks.map(({ file, wb }) =>
      extractDataFromWorkbook(file, wb)
    );
    renderPreview(extracted);
    $("result").value = generateResultText(extracted);
  } finally {
    $("processBtn").disabled = false;
    $("processBtn").textContent = "Proses";
  }
});

$("copyBtn").addEventListener("click", () => {
  if (!$("result").value) return;
  navigator.clipboard.writeText($("result").value);
  Swal.fire({ icon: "success", title: "Disalin!" });
});

$("clearBtn").addEventListener("click", () => {
  $("files").value = "";
  selectedFiles = [];
  $("fileList").textContent = "Belum ada file dipilih.";
  $("previewTableBody").innerHTML = "";
  $("tableWrap").classList.add("d-none");
  $("result").value = "";
  $("jenisBarang").value = "";
  $("masukTgl").value = new Date().toISOString().slice(0, 10);
});
