// ===============================
// JENIS BARANG - LOCAL STORAGE (PER BC)
// ===============================
const JENIS_BARANG_KEY = "yxf_jenis_barang";

// Default data awal
const DEFAULT_JENIS_BARANG = {
  "BC 2.7 Masuk": [
    "INSOLE",
    "EVA FOOTBED",
    "PU FOAM",
    "TEXTILE",
    "LOGO",
    "BOX KEMASAN",
  ],
  "BC 2.7 Keluar": ["INSOLE", "EVA FOOTBED", "TEXTILE", "BOX KEMASAN"],
  "BC 4.0 Masuk": [
    "SMART FOAM",
    "PU FOAM",
    "CHEMICAL",
    "CARTON BOX",
    "STICKER FIFO",
    "PRINT FILM",
  ],
  "BC 4.1 Keluar": ["SMART FOAM", "PU FOAM", "CHEMICAL"],
};

// ===============================
// STORAGE HELPERS
// ===============================
function loadJenisBarang() {
  return JSON.parse(localStorage.getItem(JENIS_BARANG_KEY) || "{}");
}

function saveJenisBarang(data) {
  localStorage.setItem(JENIS_BARANG_KEY, JSON.stringify(data));
}

// ===============================
// INIT STORAGE
// ===============================
function initJenisBarang() {
  if (!localStorage.getItem(JENIS_BARANG_KEY)) {
    saveJenisBarang(DEFAULT_JENIS_BARANG);
  }
}

let jenisBarangSelect;

// ===============================
// FILTER BY BC
// ===============================
function filterJenisBarangByBC(jenisBC) {
  const data = loadJenisBarang();
  const allowed = data[jenisBC] || [];

  jenisBarangSelect.clearStore();
  jenisBarangSelect.clearChoices();

  jenisBarangSelect.setChoices(
    allowed.map((v) => ({ value: v, label: v })),
    "value",
    "label",
    true
  );
}

// ===============================
// INIT UI
// ===============================
document.addEventListener("DOMContentLoaded", () => {
  initJenisBarang();

  jenisBarangSelect = new Choices("#jenisBarang", {
    removeItemButton: true,
    placeholder: true,
    placeholderValue: "Pilih jenis barang",
    searchPlaceholderValue: "Ketik untuk mencari...",
    shouldSort: false,
  });

  // Render awal sesuai BC terpilih
  const jenisBCEl = document.getElementById("jenisBC");
  filterJenisBarangByBC(jenisBCEl.value);

  // Ganti BC → refresh list
  jenisBCEl.addEventListener("change", () => {
    filterJenisBarangByBC(jenisBCEl.value);
  });

  // ===============================
  // TAMBAH JENIS BARANG BARU
  // ===============================
  document.getElementById("addJenisBtn").addEventListener("click", () => {
    const input = document.getElementById("newJenisBarang");
    const value = input.value.trim().toUpperCase();
    if (!value) return;

    const jenisBC = jenisBCEl.value;
    const data = loadJenisBarang();

    if (!data[jenisBC]) data[jenisBC] = [];

    if (data[jenisBC].includes(value)) {
      Swal.fire({ icon: "warning", text: "Jenis barang sudah ada!" });
      return;
    }

    // Simpan ke storage
    data[jenisBC].push(value);
    saveJenisBarang(data);

    // Refresh pilihan sesuai BC aktif
    filterJenisBarangByBC(jenisBC);

    // Auto select item baru
    jenisBarangSelect.setChoiceByValue(value);

    input.value = "";
  });

  // Enter = tambah
  document.getElementById("newJenisBarang").addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      document.getElementById("addJenisBtn").click();
    }
  });
});
