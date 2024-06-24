const { ipcRenderer } = require('electron');

const selectInputButton = document.getElementById('selectInput');
const selectTemplateButton = document.getElementById('selectTemplate');
const selectOutputButton = document.getElementById('selectOutput');
const processExcelButton = document.getElementById('processExcel');
const statusDiv = document.getElementById('status');
const inputFilePathSelected = document.getElementById('inputFilePathSelected');
const templateFilePathSelected = document.getElementById('templateFilePathSelected');
const outputPathSelected = document.getElementById('outputPathSelected');

let inputFilePath = null;
let templateFilePath = null;
let outputFolder = null;

selectInputButton.addEventListener('click', async () => {
    inputFilePath = await ipcRenderer.invoke('select-input-file');
    inputFilePathSelected.innerText = inputFilePath ? `Selected input file: ${inputFilePath}` : 'No file selected';
});

selectTemplateButton.addEventListener('click', async () => {
    templateFilePath = await ipcRenderer.invoke('select-template-file');
    templateFilePathSelected.innerText = templateFilePath ? `Selected template file: ${templateFilePath}` : 'No file selected';
});

selectOutputButton.addEventListener('click', async () => {
    outputFolder = await ipcRenderer.invoke('select-output-folder');
    outputPathSelected.innerText = outputFolder ? `Selected output folder: ${outputFolder}` : 'No folder selected';
});

processExcelButton.addEventListener('click', async () => {
    if (!inputFilePath || !templateFilePath || !outputFolder) {
        statusDiv.innerText = 'Please select both input & template file and output folder';
        return;
    }
    const result = await ipcRenderer.invoke('process-excel', inputFilePath, templateFilePath, outputFolder);
    statusDiv.innerText = result;
});
