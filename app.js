/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#run').click(run);
    });
  };

  function run() {
    console.log("run forest")

    return Excel.run(function (context) {
      var eifurl = "https://eif-research.feit.uts.edu.au/api/json/?" +
        "rFromDate=2017-03-14T12%3A57%3A18&rToDate=2017-03-16T12%3A57%3A18&rFamily=wasp&rSensor=ES_B_08_423_7BE2&rSubSensor=BAT"
      var x = "12"
      jQuery.ajax({ url: eifurl, async: false }).done(function (data) {
        x = jQuery.parseJSON(data)
      });


      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var bl = (x.length+4);
      var values = [
        ["Query", eifurl],
        ["",x.length],
        ["Time", "Estimate"],
      ];

      var range = sheet.getRange("A1:B3");
      range.values = values;

    	var t = context.workbook.tables.add('A4:B4', false);
      t.name = "MyTable"
      for (var i = 0; i < x.length; i++) {
        context.workbook.tables.getItem('MyTable').rows.add(null, [[x[i][0],x[i][1]]]);      
      }
      

      // var range = sheet.getRange("A4:B" + bl);
      // // Assign array value to the proxy object's values property.
      // range.values = x;


      // Create a proxy object for the range
      return context.sync();
    });

  }

})();