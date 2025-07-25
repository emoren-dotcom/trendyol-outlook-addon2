// Office Add-in komutları
function showTaskPane() {
    Office.addin.showAsTaskpane();
}

// Office.js başlatma
Office.onReady(() => {
    console.log('Office Add-in hazır');
});
            