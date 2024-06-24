const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx-populate');

function createWindow() {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'renderer.js'),
            nodeIntegration: true,
            contextIsolation: false,
        },
    });

    win.loadFile('index.html');
}

app.on('ready', createWindow);

ipcMain.handle('select-input-file', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    });
    if (canceled) {
        return null;
    } else {
        return filePaths[0];
    }
});

ipcMain.handle('select-template-file', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    });
    if (canceled) {
        return null;
    } else {
        return filePaths[0];
    }
});

ipcMain.handle('select-output-folder', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openDirectory']
    });
    if (canceled) {
        return null;
    } else {
        return filePaths[0];
    }
});

ipcMain.handle('process-excel', async (event, inputFilePath, templateFilePath, outputFolder) => {
    try {
        // Read the input file
        const inputWorkbook = await XLSX.fromFileAsync(inputFilePath);
        const inputSheet = inputWorkbook.sheet(0);
        const inputData = inputSheet.usedRange().value();

        // Read the template file
        const templateWorkbook = await XLSX.fromFileAsync(templateFilePath);

        await Promise.all(inputData.map(async (row, index) => {
            // Create a copy of the template for each row
            const newWorkbook = await XLSX.fromFileAsync(templateFilePath);
            const newSheet = newWorkbook.sheet(0);

            // Replace placeholders in the template
            newSheet.usedRange().forEach(cell => {
                const cellValue = cell.value();
                if (typeof cellValue === 'string' && cellValue.startsWith('$$')) {
                    const placeholderIndex = parseInt(cellValue.slice(2), 10);
                    if (placeholderIndex < row.length) {
                        cell.value(row[placeholderIndex]);
                    }
                }
            });

            // Save the new workbook
            const outputFilePath = path.join(outputFolder, `${row[0]}_${row[1]}.xlsx`);
            await newWorkbook.toFileAsync(outputFilePath);
        }));

        return 'Processing complete';
    } catch (error) {
        console.error(error);
        return 'Error during processing';
    }
});