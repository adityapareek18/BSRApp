// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const {app, BrowserWindow} = require('electron');
const { dialog } = require('electron').remote;
var xlsx  = require('xlsx');
require('./reportService');
var exec = require('child_process').exec, child;

function Bill(billNo, clientName) {
    this.billNo = billNo;
    this.clientName = clientName;
    this.particulars = [];
}

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

function generateBills() {
   console.log((localStorage.getItem("excelFile")));
    var workbook = xlsx.readFile(localStorage.getItem("excelFile"));
    var sheet_name_list = workbook.SheetNames;
    var xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    var bills = createBills(xlData);
}

function createBills(records) {
    var uniqueBillNos = [];
    var bills = [];
    let currentBillNo = 0;
    records.forEach(record => {
        if(currentBillNo == record["Bill No."]) {
            bill = bills.find(function(bill) {
                return (bill.billNo == currentBillNo);
            });
            bill.particulars.push(record["L.R. No."]);
        }
        else {
        var bill = new Bill();
        currentBillNo = record["Bill No."]
        bill.billNo = currentBillNo;
        bill.clientName = record["Consignee"];
        bill.particulars.push(record["L.R. No."]);
        bills.push(bill);
        }
    });
    return bills;
}

document.querySelector('#selectBtn').addEventListener('click', getFile);
document.querySelector('#generateBtn').addEventListener('click', generateBills);