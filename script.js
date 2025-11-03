document.getElementById("process").addEventListener("click", function () {
  const fileInput = document.getElementById("upload");
  if (!fileInput.files.length) {
    alert("請先選擇一個 Excel 檔案！");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const newWorkbook = XLSX.utils.book_new();

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      XLSX.utils.book_append_sheet(newWorkbook, sheet, sheetName);
    });

    const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);

    const downloadLink = document.getElementById("download");
    downloadLink.href = url;
    downloadLink.download = "合併結果.xlsx";
    downloadLink.style.display = "inline";
    downloadLink.textContent = "📥 下載合併結果.xlsx";
  };

  reader.readAsArrayBuffer(fileInput.files[0]);
});
