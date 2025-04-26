import dayjs from 'dayjs';
import validator from 'validator';
import { parse, isValid } from 'date-fns';
let table;
let originalData = [];

document.getElementById('fileInput').addEventListener('change', handleFile);

document.getElementById('downloadBtn').addEventListener('click', () => {
  if (!table) return alert('Таблица пустая');

  const data = table.getData();
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  XLSX.writeFile(workbook, 'corrected.xlsx');
});

function handleFile(event) {
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

      const cell = worksheet[cellAddress];
      if (cell && cell.t === 'n') {
        if (cell.v > 40000) {
          const jsDate = new Date((cell.v - 25569) * 86400 * 1000);
          const yyyy = jsDate.getFullYear();
          const mm = String(jsDate.getMonth() + 1).padStart(2, '0');
          const dd = String(jsDate.getDate()).padStart(2, '0');
          cell.v = `${yyyy}-${mm}-${dd}`;
          cell.t = 's';
        } else {
        }
      }
    }

    originalData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

    renderTable(originalData);
  };
  reader.readAsArrayBuffer(file);
}

function renderTable(data) {
  if (table) {
    table.destroy();
  }

  table = new Tabulator("#table", {
    data: data,
    layout: "fitData",
    reactiveData: true,
    columns: generateColumns(data),
    cellEdited: validateCell,
  });
  setTimeout(() => {
    table.redraw(true);
  }, 0);
  setTimeout(() => {
    table.getRows().forEach(row => {
      row.getCells().forEach(cell => {
        validateCell(cell);
      });
    });
  }, 0);
}

function generateColumns(data) {
  const keys = Object.keys(data[0] || {});
  return keys.map(key => ({
    title: key,
    field: key,
    editor: "input",
    cellFormatter: function(cell) {
      let value = cell.getValue();
      const field = cell.getColumn().getField();

      if (value === undefined || value === null) {
        value = '';
      } else {
        value = String(value);
      }

      const { valid, message } = validateField(field, value);

      const element = cell.getElement();
      if (!valid) {
        element.classList.add('highlight');
        element.setAttribute('title', message);
      } else {
        element.classList.remove('highlight');
        element.removeAttribute('title');
      }
      return value;
    }
  }));
}

function validateCell(cell) {
  const key = cell.getColumn().getField();
  const value = cell.getValue();
  const { valid, message } = validateField(key, value);

  const element = cell.getElement();
  if (!valid) {
    element.classList.add('highlight');
    element.setAttribute('title', message);
  } else {
    element.classList.remove('highlight');
    element.removeAttribute('title');
  }
}

function validateField(columnName, value) {
  if (value === null || value === undefined || String(value).trim() === '') {
    return { valid: false, message: 'Поле пустое' };
  }

  const normalizedColumn = columnName.toLowerCase().replace(/\s+/g, '');

  switch (normalizedColumn) {
    case 'возраст':
      return validator.isInt(value + '') ? { valid: true } : { valid: false, message: 'Возраст должен быть целым числом' };
    case 'email':
      return validator.isEmail(value + '') ? { valid: true } : { valid: false, message: 'Некорректный email' };
    case 'датарегистрации':
      if (typeof value !== 'string') return { valid: false, message: 'Некорректная дата' };

      const trimmedValue = value.trim();
      const formats = ['yyyy-MM-dd', 'dd.MM.yyyy', 'yyyy/MM/dd'];

      const parsedDate = formats
        .map(fmt => parse(trimmedValue, fmt, new Date()))
        .find(date => isValid(date));

      return parsedDate
        ? { valid: true }
        : { valid: false, message: 'Некорректная дата' };
    default:
      return { valid: true };
  }
}