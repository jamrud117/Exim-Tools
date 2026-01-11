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
  let sheetPL = null;
  let sheetINV = null;
  let sheetsDATA = null;
  let kontrakNo = "";
  let kontrakTgl = "";

  for (const file of files) {
    const wb = await readExcelFile(file);
    const type = detectFileType(wb);

    if (type === "DATA") sheetsDATA = wb.Sheets;
    if (type === "INV") sheetINV = wb.Sheets[wb.SheetNames[0]];
    if (type === "PL") {
      sheetPL = wb.Sheets[wb.SheetNames[0]];
      const kontrak = extractKontrakInfoFromPL(sheetPL);
      kontrakNo = kontrak.kontrakNo;
      kontrakTgl = kontrak.kontrakTgl;
    }
  }

  if (!sheetPL || !sheetINV || !sheetsDATA) {
    Swal.fire("File belum lengkap");
    return;
  }

  checkAll(sheetPL, sheetINV, sheetsDATA, kurs, kontrakNo, kontrakTgl);
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
