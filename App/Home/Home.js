/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
      
      $('#generate-template').on('click', generateTemplate);
            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };
  
  function generateTemplate() {
    Excel.run(function (ctx) {
      
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      
      sheet.getRange('A1').values = 'num x';
      sheet.getRange('B1').values = 'num y'

    })
  }

    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
      function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
              calculateSum(result.value);
          } else {
              app.showNotification('Error:', result.error.message);
          }
      }
    );
    }
  
  function calculateSum(num) {
    num = parseInt(num);
    if (isNaN(num)) {
      num = 0;
    }   

    Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      
      var x = sheet.getRange('A2');
      var y = sheet.getRange('B2');
      
      x.load('values');
      y.load('values');
      
      return ctx.sync().then(function() {
        var sum = x.values[0][0] + y.values[0][0] + num;
        sheet.getRange('A4').values = sum;
      });
    })
  }
})();