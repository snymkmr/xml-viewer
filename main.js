// main.js

const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const url = require('url');
const fs = require('fs');
const ExcelJS = require('exceljs');

let mainWindow;
let fileName;
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule: true
        }
    });

    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file:',
        slashes: true
    }));

    mainWindow.on('closed', function () {
        mainWindow = null;
    });
}

app.on('ready', createWindow);

ipcMain.on('file-selected', (event, filePath) => {
    console.log(`File selected: ${filePath}`);
    fs.readFile(filePath, 'utf-8', (err, data) => {
        if (err) {
            console.error(`Error reading file: ${err.message}`);
            event.sender.send('xml-data', `Error reading file: ${err.message}`);
        } else {
            event.sender.send('xml-data', data);
        }
    });
});

ipcMain.on('copy-to-clipboard', (event, tableData) => {
    const clipboardData = tableData.map(row => row.join('\t')).join('\n');
    event.sender.send('copy-success', clipboardData);
});

ipcMain.on('open-excel', (event, filePath) => {
    shell.openPath(filePath);
});

ipcMain.on('export-to-excel', (event, tableData, fileName) => {
    console.log("🚀 ~ ipcMain.on ~ fileName:", fileName)
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');

    // Add the table data to the worksheet
    tableData.forEach(row => {
        worksheet.addRow(row);
    });

    const excelFilePath = dialog.showSaveDialogSync(mainWindow, {
        title: 'Save Excel File',
        defaultPath: fileName,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
    });

    if (excelFilePath) {
        workbook.xlsx.writeFile(excelFilePath)
            .then(() => {
                event.sender.send('excel-success', excelFilePath);
                event.sender.send('open-excel', excelFilePath); // Open the file in Excel
            })
            .catch(err => {
                console.error(`Error writing Excel file: ${err.message}`);
            });
    }
});