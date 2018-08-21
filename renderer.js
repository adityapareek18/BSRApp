// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const { app, BrowserWindow } = require('electron');
const { dialog } = require('electron').remote;
var Excel = require('exceljs');
var fs = require('fs');

function Particular(part, lr, date, qty, amount) {
    this.part = part;
    this.lr = lr;
    this.date = date;
    this.qty = qty;
    this.amount = amount;
}
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
            dialog.showMessageBox({ message: "Please select the file." });
        }
    }
    )
}

function extractData(dataSheet) {
    var dataRecords = [];
    var headerRow = dataSheet.findRow(1).values;
    dataSheet.eachRow(function (row, rowNum) {
        if (rowNum > 1) {
            var record = {};
            var i = 1;
			row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
				record[headerRow[i]] = cell.value;
				i++;
			});
            dataRecords.push(record);
        }
    });
    return dataRecords;
}

function generateBills() {
    dialog.showOpenDialog({
        properties: ['openDirectory']
    }, function (file) {
        if (file !== undefined) {
			localStorage.setItem("dest", file);
            generateBillsAtDest();
        }
        else {
            dialog.showMessageBox({ message: "Please select the directory." });
        }
    }
    )
}

function generateBillsAtDest() {
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(localStorage.getItem("excelFile"))
        .then(wb => {
            var dataSheet = wb.getWorksheet(1);
            var xlData = extractData(dataSheet);
            var bills = createBillObjects(xlData);
            bills.forEach(function (bill) {
                generateExcelBill(bill);
            });
			var dir = localStorage.getItem("dest");
            alert("Bills have been successfully generated at "+dir+" folder");
        });
}

function createBillObjects(records) {
    var bills = [];
    let currentBillNo = 0;
    records.forEach(record => {
        if (currentBillNo == record["Bill No."]) {
            bill = bills.find(function (bill) {
                return (bill.billNo == currentBillNo);
            });
		bill.particulars.push(new Particular(record["Particular"], record["L.R. No."],
                record["Date"], record["Qty."], record["Amount"]));
        }
        else {
            var bill = new Bill();
            currentBillNo = record["Bill No."]
            bill.billNo = currentBillNo;
            bill.clientName = record["Consignee"];
            bill.particulars.push(new Particular(record["Particular"], record["L.R. No."],
                record["Date"], record["Qty."], record["Amount"]));
            bills.push(bill);
        }
    });
    return bills;
}

function generateExcelBill(bill) {
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile("FormatOfBill.xlsx").
        then(function (wb) {
            var dataSheet = wb.getWorksheet(1);
            setStyles(dataSheet, bill);
            setContent(dataSheet, bill);
            var dir = localStorage.getItem("dest");
            if (dir !== 'undefined') {
                if (!fs.existsSync(dir)) {
                    fs.mkdirSync(dir);
                }
                wb.xlsx.writeFile(dir + "\\" + bill.billNo + ".xlsx");
            }
        });
}

function setStyles(dataSheet, bill) {
    dataSheet.getCell('A2').value = {
        'richText': [
            { 'font': { 'bold': true, 'size': 36, 'color': { 'rgb': '000000' }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': ' BALAJI \n' },
            { 'font': { 'bold': true, 'size': 18, 'color': { 'rgb': '000000' }, 'name': 'Calibri', 'scheme': 'minor' }, 'text': '\n SPEED ROADWAYS' }
        ]
    };
    dataSheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    dataSheet.getCell('A16', 'B16', 'C16', 'D16', 'E16').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '880000FF' } };
    dataSheet.getCell('B16').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '880000FF' } };
    for (var i = 16; i <= 40; i++) {
        ['A' + i, 'B' + i, 'C' + i, 'D' + i, 'E' + i].map(key => {
            dataSheet.getCell(key).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    }
}

function setContent(dataSheet, bill) {
    dataSheet.getCell('A11').value = "M/s:  " + bill.clientName;
    dataSheet.getCell('D11').value = "Bill No.:  " + bill.billNo;
    dataSheet.getCell('A14').value = "GSTN No.:  ";
    dataSheet.getCell('D14').value = "Date:  " + getLastDateOfMonth();
    var j = 0;
    var sumOfPerticulars = 0;
    for (var i = 17; i <= 39; i++) {
        if (typeof (bill.particulars[j]) != 'undefined' && j <= bill.particulars.length) {
            dataSheet.getCell('A' + i).value = bill.particulars[j]['part'];
            dataSheet.getCell('B' + i).value = bill.particulars[j]['lr'];
            dataSheet.getCell('C' + i).value = bill.particulars[j]['date'];
            dataSheet.getCell('D' + i).value = bill.particulars[j]['qty'];
            dataSheet.getCell('E' + i).value = bill.particulars[j]['amount'];
            sumOfPerticulars = sumOfPerticulars + bill.particulars[j]['amount'];
        }
        j++;
    }
    dataSheet.getCell('E40').value = sumOfPerticulars;
}

function openFile() {
    dialog.showOpenDialog({
        properties: ['openDirectory']
    }, function (file) {
        if (file !== undefined) {
            var excelDataFile = fs.readFileSync("./orders.xlsx");
            fs.writeFileSync(file + "\\"+"orders.xlsx", excelDataFile);
            alert("Data File with sample data has Been generated at following location. \n" + file);
        }
        else {
            dialog.showMessageBox({ message: "Please select the directory." });
        }
    }
    )
}

function getLastDateOfMonth(){
	var currentTime = new Date();
	var todayMonth = currentTime.getMonth() + 1;
	var toDay = currentTime.getDate();
	var todayYear = currentTime.getFullYear();
	var lDM = new Date((new Date(todayYear, todayMonth,1))-1);
	return lDM.getDate() + "/" + lDM.getMonth() + "/" + lDM.getFullYear();
 }

document.querySelector('#openBtn').addEventListener('click', openFile);
document.querySelector('#selectBtn').addEventListener('click', getFile);
document.querySelector('#generateBtn').addEventListener('click', generateBills);