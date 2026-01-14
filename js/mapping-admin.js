const page = window.location.pathname.split("/").pop();
document.querySelectorAll(".nav-links a").forEach((link) => {
  if (link.getAttribute("href") === page) {
    link.classList.add("active");
  }
});

// ==============================
// STORAGE & INITIAL LOAD
// ==============================

let mappings = {};
let currentPage = 1;
const perPage = 5;
let editingKey = null;

function load() {
  const saved = localStorage.getItem("companyMappings");
  if (saved) mappings = JSON.parse(saved);

  render();
}

function saveToStorage() {
  localStorage.setItem("companyMappings", JSON.stringify(mappings));
}

// ==============================
// INPUT LIMITER (MAX 22 CHAR NPWP)
// ==============================

function limitNPWPInput(id) {
  const el = document.getElementById(id);
  el.addEventListener("input", () => {
    if (el.value.length > 22) {
      el.value = el.value.slice(0, 22);
    }
  });
}

// Terapkan pada input create + edit
limitNPWPInput("npwp");
limitNPWPInput("editNPWP");

// ==============================
// ADD MAPPING
// ==============================

function saveMapping() {
  const pt = document.getElementById("mapPT").value.trim();
  const check = document.getElementById("mapCheck").value.trim();
  const address = document.getElementById("mapAddress").value.trim();
  const code = document.getElementById("mapKode").value.trim();
  const uraian = document.getElementById("mapUraian").value.trim();
  const qty = document.getElementById("mapQty").value.trim();
  const cif = document.getElementById("mapCIF").value.trim();
  const suratjalan = document.getElementById("mapSJ").value.trim();
  const no = document.getElementById("mapNo").value.trim();
  const npwp = document.getElementById("npwp").value.trim();

  // Validasi wajib isi
  const fields = [
    { value: pt, label: "Nama customer untuk dropdown" },
    { value: check, label: "Nama customer untuk pengecekkan" },
    { value: address, label: "Customer Address" },
    { value: code, label: "Code customer" },
    { value: uraian, label: "Nama item customer" },
    { value: qty, label: "Header quantity customer" },
    { value: cif, label: "Header CIF customer" },
    { value: suratjalan, label: "Header surat jalan customer" },
    { value: no, label: "Header No / Item No" },
    { value: npwp, label: "NPWP customer" },
  ];

  const emptyField = fields.find((f) => !f.value);
  if (emptyField) {
    Swal.fire({
      icon: "error",
      scrollbarPadding: false,
      text: `${emptyField.label} wajib diisi!`,
    });
    return;
  }

  mappings[pt] = {
    check: check,
    address: address,
    kode: code,
    uraian: uraian,
    qty: qty,
    cif: cif,
    suratjalan: suratjalan,
    no,
    npwp: npwp,
  };

  saveToStorage();
  render();

  const modal = bootstrap.Modal.getInstance(
    document.getElementById("addModal")
  );
  if (modal) modal.hide();

  Swal.fire({
    position: "top-mid",
    icon: "success",
    scrollbarPadding: false,
    title: "Mapping berhasil disimpan!",
    showConfirmButton: false,
    timer: 1500,
  });
}

// ==============================
// RENDER TABLE + PAGINATION
// ==============================

function render() {
  const tbody = document.getElementById("mappingTable");
  const tableWrapper = document.querySelector(".table-wrapper");
  const pagination = document.getElementById("pagination");

  const keys = Object.keys(mappings);
  const totalPages = Math.ceil(keys.length / perPage);

  const start = (currentPage - 1) * perPage;
  const end = start + perPage;

  const pageKeys = keys.slice(start, end);

  tbody.innerHTML = "";

  if (pageKeys.length === 0) {
    tableWrapper.classList.add("empty");
    tbody.innerHTML = `
      <tr>
        <td colspan="11" class="text-center">Tidak ada data</td>
      </tr>
    `;
  } else {
    tableWrapper.classList.remove("empty");
    pageKeys.forEach((pt, i) => {
      const m = mappings[pt];
      tbody.innerHTML += `
        <tr>
          <td>${start + i + 1}</td>
          <td>${pt}</td>
          <td>${m.check}</td>
          <td>${m.address}</td>
          <td>${m.kode}</td>
          <td>${m.uraian}</td>
          <td>${m.qty}</td>
          <td>${m.cif}</td>
          <td>${m.suratjalan || "-"}</td>
          <td>${m.no || "-"}</td>
          <td>${m.npwp}</td>
          <td>
            <button class="btn btn-warning" onclick="openEdit('${pt}')">Edit</button>
            <button class="btn btn-danger" onclick="deleteMapping('${pt}')">Hapus</button>
          </td>
        </tr>
      `;
    });
  }

  pagination.innerHTML = "";
  for (let i = 1; i <= totalPages; i++) {
    pagination.innerHTML += `
      <li class="page-item ${i === currentPage ? "active" : ""}">
        <button class="page-link" onclick="changePage(${i})">${i}</button>
      </li>
    `;
  }
}

function changePage(page) {
  currentPage = page;
  render();
}

// ==============================
// EDIT DATA
// ==============================

function openEdit(pt) {
  editingKey = pt;
  const d = mappings[pt];

  document.getElementById("editPT").value = pt;
  document.getElementById("editCheck").value = d.check;
  document.getElementById("editAddress").value = d.address;
  document.getElementById("editKode").value = d.kode;
  document.getElementById("editUraian").value = d.uraian;
  document.getElementById("editQty").value = d.qty;
  document.getElementById("editCIF").value = d.cif;
  document.getElementById("editSJ").value = d.suratjalan;
  document.getElementById("editNo").value = d.no || "";
  document.getElementById("editNPWP").value = d.npwp;

  new bootstrap.Modal(document.getElementById("editModal")).show();
}

function saveEdit() {
  const pt = editingKey;

  mappings[pt] = {
    check: document.getElementById("editCheck").value.trim(),
    kode: document.getElementById("editKode").value.trim(),
    address: document.getElementById("editAddress").value.trim(),
    uraian: document.getElementById("editUraian").value.trim(),
    qty: document.getElementById("editQty").value.trim(),
    cif: document.getElementById("editCIF").value.trim(),
    suratjalan: document.getElementById("editSJ").value.trim(),
    no: document.getElementById("editNo").value.trim(),
    npwp: document.getElementById("editNPWP").value.trim(),
  };

  saveToStorage();

  bootstrap.Modal.getInstance(document.getElementById("editModal")).hide();
  render();
  Swal.fire({
    position: "top-mid",
    icon: "success",
    title: "Berhasil edit data!",
    showConfirmButton: false,
    scrollbarPadding: false,
    timer: 1500,
  });
}
function deleteMapping(pt) {
  Swal.fire({
    title: `Hapus data "${pt}"?`,
    text: "Data ini akan dihapus secara permanen!",
    icon: "warning",
    showCancelButton: true,
    scrollbarPadding: false,
    confirmButtonColor: "#d33",
    cancelButtonColor: "#3085d6",
    confirmButtonText: "Ya, hapus!",
    cancelButtonText: "Batalkan",
  }).then((result) => {
    if (result.isConfirmed) {
      delete mappings[pt];
      saveToStorage();

      if ((currentPage - 1) * perPage >= Object.keys(mappings).length)
        currentPage--;

      render();

      Swal.fire({
        icon: "success",
        title: "Terhapus!",
        scrollbarPadding: false,
        text: `Data "${pt}" berhasil dihapus.`,
        timer: 1500,
        showConfirmButton: false,
      });
    }
  });
}

// ==============================
// EXPORT / IMPORT
// ==============================

function exportMapping() {
  const blob = new Blob([JSON.stringify(mappings, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "mapping.json";
  a.click();
}

function importMapping() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json";

  input.onchange = (e) => {
    const reader = new FileReader();
    reader.onload = () => {
      mappings = JSON.parse(reader.result);
      saveToStorage();

      currentPage = 1;
      render();
      Swal.fire({
        position: "top-mid",
        icon: "success",
        title: "Berhasil import mapping!",
        scrollbarPadding: false,
        showConfirmButton: false,
        timer: 1500,
      });
    };

    reader.readAsText(e.target.files[0]);
  };

  input.click();
}

// Initial load
load();
