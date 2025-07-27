function showPopup() {
    Office.context.ui.displayDialogAsync('https://emoren-dotcom.github.io/trendyol-outlook-addon2/popup.html',
      { height: 30, width: 20 },
      function (result) {
        console.log('Dialog opened');
      }
    );
  }
  