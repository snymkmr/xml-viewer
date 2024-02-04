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

    mainWindow.maximizable = true;  // Enable maximizable
    mainWindow.maximize();  // Maximize the window

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

    // Extract the file extension from the filePath
    const fileExtension = path.extname(filePath).toLowerCase();

    // Read the file content
    fs.readFile(filePath, 'utf-8', (err, data) => {
        if (err) {
            console.error(`Error reading file: ${err.message}`);
            event.sender.send('xml-data', `Error reading file: ${err.message}`);
        } else {
            // Determine the file type based on the extension
            if (fileExtension === '.xml') {
                event.sender.send('xml-data', data);
            } else if (fileExtension === '.csv') {
                event.sender.send('csv-data', data);
            } else {
                console.error('Unsupported file type');
                event.sender.send('unsupported-file-type', 'Unsupported file type');
            }
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
    const workbook = new ExcelJS.Workbook();
    if (!fileName) fileName = 'Untitled';
    const worksheet = workbook.addWorksheet(fileName);

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
