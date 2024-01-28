// main.js
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const url = require('url');
const fs = require('fs');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule: true,
            enableOverlayScrollbars: true // Enable overlay scrollbars
        }
    });

    mainWindow.maximize();

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
