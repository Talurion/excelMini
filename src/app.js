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

  validateTableData(data) {
    const types = ['date', 'string', 'string', 'string', 'number', 'number'];
    const validators = {
      date: value => {

        if (value instanceof Date) return true;
        if (typeof value === 'string') {
          return /^\d{4}-\d{2}-\d{2}$/.test(value);
        }
        return false;
      },
      string: value => typeof value === 'string',
      number: value => typeof value === 'number' && !isNaN(value),
    };
    const keys = Object.keys(data[0] || {});
    data.forEach(row => {

      const firstKey = keys[0];
      const val = row[firstKey];
      if (typeof val === 'number') {

        const jsDate = new Date((val - 25569) * 86400 * 1000);
        row[firstKey] = jsDate.toISOString().slice(0, 10);
      }
      row._invalidFields = {};
      keys.forEach((key, idx) => {
        const type = types[idx];
        const validator = validators[type];
        if (validator && !validator(row[key])) {
          row._invalidFields[key] = true;
        }
      });
    });
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
      this.renderTable(this.originalData);
    };
    reader.readAsArrayBuffer(file);
  }

  renderTable(data) {
    this.validateTableData(data);

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
        });
      });
    }, 0);
  }

  generateColumns(data) {
    const keys = Object.keys(data[0] || {}).filter(key => key !== '_invalidFields');
    return keys.map(key => ({
      title: key,
      field: key,
      editor: "input",
      formatter: function(cell) {
        const value = cell.getValue();
        const row = cell.getRow().getData();
        const invalid = row._invalidFields && row._invalidFields[key];
        const display = value == null ? '' : String(value);
        return invalid
          ? `<span class="invalid-cell">${display}</span>`
          : display;
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

document.addEventListener('DOMContentLoaded', () => {
  new ExcelTableManager('fileInput', 'downloadBtn', 'table');
});