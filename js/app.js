let mappings = {};

async function loadMappings() {
  const local = localStorage.getItem("companyMappings");
  if (local) {
    mappings = JSON.parse(local);
    return;
  }

  try {
    const res = await fetch("mapping.json");
    mappings = await res.json();
  } catch (e) {
    console.error("Gagal load mapping.json");
  }
}
loadMappings();

const currentPage = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === currentPage) {
    link.classList.add("active");
  }
});

// === Fungsi utama untuk memproses 3 file ===
async function processFiles(files) {
  if (!files || files.length !== 3) {
    Swal.fire({
      icon: "error",
      title: "Oops...",
      scrollbarPadding: false,
      text: "Upload 3 file (Draft, INV, PL) terlebih dahulu!",
    });
    return;
  }

  console.log("Mulai proses", files.length, "file...");

  try {
    // Konversi FileList ke array agar bisa di-loop
    const fileArray = Array.from(files);

    // Baca workbook dari semua file (hasil Promise.all)
    const workbooks = await Promise.all(fileArray.map(readExcelFile));

    // Variabel untuk menampung hasil deteksi
    let sheetPL = null;
    let sheetINV = null;
    let sheetsDATA = null;

    // === Deteksi otomatis tipe file ===
    workbooks.forEach((wb, i) => {
      const type = detectFileType(wb);
      console.log(`File ${i + 1}: tipe terdeteksi =`, type);

      if (type === "PL") {
        sheetPL = wb.Sheets[wb.SheetNames[0]];
      } else if (type === "INV") {
        sheetINV = wb.Sheets[wb.SheetNames[0]];
      } else if (type === "DATA") {
        sheetsDATA = {
          HEADER: wb.Sheets["HEADER"],
          DOKUMEN: wb.Sheets["DOKUMEN"],
          KEMASAN: wb.Sheets["KEMASAN"],
          BARANG: wb.Sheets["BARANG"],
          ENTITAS: wb.Sheets["ENTITAS"],
        };
      }
    });

    // === Validasi hasil deteksi ===
    if (!sheetPL || !sheetINV || !sheetsDATA) {
      Swal.fire({
        icon: "error",
        title: "Oops...",
        scrollbarPadding: false,
        text: "Tidak bisa mendeteksi file DRAFT / INV / PL.\nPastikan struktur dan isi file sudah benar.",
      });

      console.error({ sheetPL, sheetINV, sheetsDATA });
      return;
    }

    // === Parsing kurs agar input type=number tidak error ===
    const kursCell = getCellValue(sheetsDATA.HEADER, "BW2");
    const kursParsed = parseKurs(kursCell) || 1;
    document.getElementById("kurs").value = kursParsed;

    // === Ambil kontrak dari PL ===
    const { kontrakNo, kontrakTgl } = extractKontrakInfoFromPL(sheetPL);
    console.log("Kontrak ditemukan:", kontrakNo, kontrakTgl);

    // === Jalankan pengecekan utama ===
    checkAll(sheetPL, sheetINV, sheetsDATA, kursParsed, kontrakNo, kontrakTgl);

    console.log("CheckAll selesai dieksekusi ✅");
  } catch (err) {
    console.error("Terjadi error saat memproses file:", err);
    Swal.fire({
      icon: "error",
      title: "Oops...",
      scrollbarPadding: false,
      text: "Terjadi kesalahan saat memproses file. Lihat konsol (F12) untuk detailnya.",
    });
  }
}

// === Deteksi otomatis tipe file berdasarkan isi sheet ===
function detectFileType(wb) {
  const sheetNames = wb.SheetNames.map((n) => n.toUpperCase());

  // File Draft (DATA) memiliki 4 sheet utama
  if (
    sheetNames.includes("HEADER") &&
    sheetNames.includes("BARANG") &&
    sheetNames.includes("KEMASAN") &&
    sheetNames.includes("DOKUMEN")
  ) {
    return "DATA";
  }

  // Cek isi beberapa baris pertama untuk kata kunci
  const firstSheet = wb.Sheets[wb.SheetNames[0]];
  if (!firstSheet || !firstSheet["!ref"]) return "UNKNOWN";

  const ref = XLSX.utils.decode_range(firstSheet["!ref"]);
  const maxRow = Math.min(ref.e.r, 10);
  const maxCol = Math.min(ref.e.c, 10);

  for (let r = ref.s.r; r <= maxRow; r++) {
    for (let c = ref.s.c; c <= maxCol; c++) {
      const cell = firstSheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || !cell.v) continue;
      const v = String(cell.v).toUpperCase();

      if (v.includes("PACKING LIST")) return "PL";
      if (v.includes("INVOICE")) return "INV";
    }
  }

  return "UNKNOWN";
}

// === Event listener tombol ===
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("btnCheck");
  const fileInput = document.getElementById("files");

  btn.addEventListener("click", async () => {
    const files = fileInput.files;
    if (!files || files.length === 0) {
      Swal.fire({
        icon: "error",
        title: "Oops...",
        text: "Pilih 3 file Excel terlebih dahulu!",
        scrollbarPadding: false,
      });
      return;
    }
    await processFiles(files);
    document.getElementById("filter").value = "beda";
    applyFilter();
  });

  // Filter hasil
  const filterSelect = document.getElementById("filter");
  if (filterSelect) {
    filterSelect.addEventListener("change", applyFilter);
  }
});
