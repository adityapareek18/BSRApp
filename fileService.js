const { file } = require('fs');
const { dialog } = require('electron').remote;

    function showFileSelectorDialog() {
        files = dialog.showOpenDialog();
    }