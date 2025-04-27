class ExcelTableManager {
  constructor(fileInputId, downloadBtnId, tableContainerId) {
    this.fileInput = document.getElementById(fileInputId);
    this.downloadBtn = document.getElementById(downloadBtnId);
    this.tableContainerId = tableContainerId;
    this.table = null;
    this.originalData = [];

    this.fileInput.addEventListener('change', (event) => this.handleFile(event));
    this.downloadBtn.addEventListener('click', () => this.downloadFile());
  }

  handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      for (const cellAddress in worksheet) {
        if (cellAddress[0] === '!') continue;
      }

      this.originalData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      console.table(this.originalData);
      this.renderTable(this.originalData);
    };
    reader.readAsArrayBuffer(file);
  }

  renderTable(data) {
    if (this.table) {
      this.table.destroy();
    }

    this.table = new Tabulator(`#${this.tableContainerId}`, {
      data: data,
      layout: "fitData",
      reactiveData: true,
      columns: this.generateColumns(data),
    });

    setTimeout(() => {
      this.table.redraw(true);
    }, 0);

    setTimeout(() => {
      this.table.getRows().forEach(row => {
        row.getCells().forEach(cell => {
          // Здесь можно добавить обработку ячеек
        });
      });
    }, 0);
  }

  generateColumns(data) {
    const keys = Object.keys(data[0] || {});
    return keys.map(key => ({
      title: key,
      field: key,
      editor: "input",
      cellFormatter: function(cell) {
        let value = cell.getValue();
        return value === undefined || value === null ? '' : String(value);
      }
    }));
  }

  downloadFile() {
    if (!this.table) return alert('Таблица пустая');

    const data = this.table.getData();
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, 'corrected.xlsx');
  }
}

// Инициализация менеджера таблицы после загрузки страницы
document.addEventListener('DOMContentLoaded', () => {
  new ExcelTableManager('fileInput', 'downloadBtn', 'table');
});