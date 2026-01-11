function toggleStatusJalur() {
  const jenisBC = document.getElementById("jenisBC").value;
  const isBC4 = jenisBC.startsWith("BC 4.");
  const isMasuk = jenisBC.includes("Masuk");

  const headerPengirim = document.getElementById("headerPengirim");
  const labelTanggal = document.getElementById("labelTanggal");

  const wrapJalur = document.getElementById("statusJalurWrap");
  const wrapOverride = document.getElementById("jalurOverrideWrap");

  const colTanggal = document.getElementById("colTanggal");
  const colJenisBC = document.getElementById("colJenisBC");

  // ===============================
  // HELPER GRID
  // ===============================
  function setCol(el, col) {
    el.classList.remove("col-md-4", "col-md-6", "col-md-12");
    el.classList.add(`col-md-${col}`);
  }

  // ===============================
  // HEADER & LABEL
  // ===============================
  headerPengirim.textContent = isMasuk ? "PENGIRIM" : "PENERIMA";
  labelTanggal.textContent = isMasuk ? "Tanggal Masuk" : "Tanggal Keluar";

  // ===============================
  // LOGIC JALUR & GRID
  // ===============================
  if (isBC4 || !isMasuk) {
    // 🔥 BC 4.x & BC 2.7 Keluar → jalur aktif
    wrapJalur.style.display = "";
    wrapOverride.style.display = "";

    setCol(colTanggal, 4);
    setCol(colJenisBC, 4);
  } else {
    // 🔥 BC 2.7 Masuk → tanpa jalur
    wrapJalur.style.display = "none";
    wrapOverride.style.display = "none";

    setCol(colTanggal, 6);
    setCol(colJenisBC, 6);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const jenisBCEl = document.getElementById("jenisBC");

  toggleStatusJalur();

  jenisBCEl.addEventListener("change", () => {
    toggleStatusJalur();
  });
});
