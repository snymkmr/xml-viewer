const { ipcRenderer, clipboard } = require('electron');
let possibleTags = ['Books', 'Customer']; // Add other possible tags as needed
const Papa = require('papaparse');
let foundTag = null;
let elements = null;
let isExporting = false;
let fileName = '';

document.getElementById('fileInput').addEventListener('change', (event) => {
  const filePath = event.target.files[0].path;
  fileName = extractFileName(filePath);
  ipcRenderer.send('file-selected', filePath);
  showNotification('File browsed successfully!', 'fileBrowseNotification');
});

ipcRenderer.on('xml-data', (event, data) => {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(data, 'text/xml');

  for (const tagName of possibleTags) {
    elements = xmlDoc.getElementsByTagName(tagName);

    if (elements.length > 0) {
      console.log(`Uploaded file contains ${tagName} tag.`);
      foundTag = tagName;
      // Process the tag accordingly
      break; // Exit the loop after finding the first matching tag
    }
  }

  // If no matching tag is found
  if (!foundTag) {
    console.log('Uploaded file does not contain any expected tags.');
  }

  // After appending the table content
  const tableContainer = document.getElementById('xmlTableContainer');
  const nothingToSeeHere = document.getElementById('nothingToSeeHere');
  tableContainer.classList.toggle('empty', tableContainer.children.length === 0);

  xmlTableContainer.innerHTML = ''; // Clear previous content

  // Create a table element
  const table = document.createElement('table');

  // Create the header row
  const headerRow = table.insertRow();
  const headers = Array.from(elements[0].children).map((child) => child.tagName);

  headers.forEach((headerText) => {
    const headerCell = document.createElement('th');
    headerCell.textContent = headerText;
    headerRow.appendChild(headerCell);
  });

  // Iterate over elements and create rows in the table
  for (let i = 0; i < elements.length; i++) {
    const book = elements[i];

    const newRow = table.insertRow();

    // Iterate over child elements of the current book and create cells
    for (let j = 0; j < book.children.length; j++) {
      const dataCell = newRow.insertCell(j);
      dataCell.textContent = book.children[j].textContent;
    }
  }

  xmlTableContainer.appendChild(table);

  // Add event listeners for Excel and Copy buttons
  document.getElementById('copyTableBtn').addEventListener('click', () => {
    copyToClipboard(table);
    showNotification('Table data copied to clipboard!', 'copyTableNotification');
  });
});

ipcRenderer.on('csv-data', (event, data) => {
  Papa.parse(data, {
    complete: (result) => {
      const headers = result.data[0];

      const table = document.createElement('table');
      const headerRow = table.insertRow();

      headers.forEach((headerText) => {
        const headerCell = document.createElement('th');
        headerCell.textContent = headerText;
        headerRow.appendChild(headerCell);
      });

      for (let i = 1; i < result.data.length; i++) {
        const rowData = result.data[i];
        const newRow = table.insertRow();

        rowData.forEach((cellText, j) => {
          const dataCell = newRow.insertCell(j);
          dataCell.textContent = cellText;
        });
      }

      const tableContainer = document.getElementById('xmlTableContainer');
      tableContainer.innerHTML = '';
      tableContainer.appendChild(table);

      document.getElementById('copyTableBtn').addEventListener('click', () => {
        copyToClipboard(table);
        showNotification('Table data copied to clipboard!', 'copyTableNotification');
      });
    }
  });
});

function copyToClipboard(table) {
  const tableData = Array.from(table.rows).map((row) => Array.from(row.cells).map((cell) => cell.textContent));
  const tableText = tableData.map((row) => row.join('\t')).join('\n');
  clipboard.writeText(tableText);
  console.log('Table content copied to clipboard:', tableText);
}

function toggleExportButton(disabled) {
  const exportButton = document.getElementById('exportButton');
  exportButton.disabled = disabled;
}

function exportToExcel() {
  if (isExporting) {
    return; // If already exporting, do nothing
  }
  isExporting = true;
  toggleExportButton(true); // Disable the export button

  const tableData = Array.from(document.querySelectorAll('table tr')).map(row =>
    Array.from(row.querySelectorAll('td, th')).map(cell => cell.textContent)
  );

  ipcRenderer.send('export-to-excel', tableData, fileName);
  showNotification('Opening table in Excel...', 'exportToExcelNotification');
  isExporting = false;
  toggleExportButton(false);
}

function extractFileName(filePath) {
  return filePath.split('\\').pop().split('/').pop().replace(/\.[^/.]+$/, '');;
}

ipcRenderer.on('open-excel', (event, filePath) => {
  ipcRenderer.send('open-excel', filePath);
});

document.getElementById('fileInputBtn').addEventListener('click', () => {
  document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', (event) => {
  const fileName = event.target.files[0].name;
  document.querySelector('.file-input-label').textContent = fileName;
});

function showNotification(message, notificationId) {
  const notification = document.getElementById(notificationId);
  notification.textContent = message;
  notification.classList.add('show');

  setTimeout(() => {
    notification.classList.remove('show');
  }, 5000); // Hide after 5 seconds
}
