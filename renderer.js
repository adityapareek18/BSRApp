// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const {dialog} = require('electron').remote;
function getFile() {
    dialog.showOpenDialog({
        properties: ['openFile', 'multiSelections']
    }), function (files) {
        if (files !== undefined) {
            dialog.showMessageBox("Please select a file");
        }
    }
}

document.querySelector('#selectBtn').addEventListener('click', getFile);