// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const {app, BrowserWindow} = require('electron');
const { dialog } = require('electron').remote;
var xlsx  = require('xlsx');

function getFile() {
    dialog.showOpenDialog({
        properties: ['openFile', 'multiSelections']
    }, function (file) {
        if (file !== undefined) {
            localStorage.setItem("excelFile", file);
        }
        else {
            dialog.showMessageBox({message: "Please select the file."});
        }
    }
    )
}

function readWorkbook(file) {
    var workbook = xlsx.readFile(file);
}

function generateBills() {
   console.log((localStorage.getItem("excelFile")));
    var workbook = xlsx.readFile(localStorage.getItem("excelFile"));
    var sheet_name_list = workbook.SheetNames;
    var xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    console.log(xlData);
}

document.querySelector('#selectBtn').addEventListener('click', getFile);
document.querySelector('#generateBtn').addEventListener('click', generateBills);