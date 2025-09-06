async function loadExcel(filePath, sheetName, containerId) {
  try {
    const response = await fetch(filePath);
    if (!response.ok) throw new Error(`Failed to fetch ${filePath}`);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    if (!workbook.Sheets[sheetName]) {
      throw new Error(`Sheet "${sheetName}" not found in ${filePath}`);
    }

    const sheet = workbook.Sheets[sheetName];
    const html = XLSX.utils.sheet_to_html(sheet);
    document.getElementById(containerId).innerHTML = html;
  } catch (err) {
    document.getElementById(containerId).innerHTML =
      `<p style="color:red;">Error: ${err.message}</p>`;
  }
}
