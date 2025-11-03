document.getElementById('processBtn').addEventListener('click', () => {
  const fileInput = document.getElementById('excelFile');
  const file = fileInput.files[0];
  if (!file) {
    alert('請先選擇 Excel 檔案');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    const processedData = jsonData.map(row => ({
      ...row,
      處理結果: row.數值 ? row.數值 * 2 : ''
    }));

    const newSheet = XLSX.utils.json_to_sheet(processedData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, '處理結果');

    XLSX.writeFile(newWorkbook, '轉換後結果.xlsx');
  };

  reader.readAsArrayBuffer(file);
});
