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
        "rFromDate=" + encodeURI( $("#rFromDate").val() ) +
        "&rToDate=" + encodeURI( $("#rToDate").val() ) +
        "&rFamily=wasp" + 
        "&rSensor=" + $("#rSensor").val() + 
        "&rSubSensor=" + $("#rSubSensor").val()
      var x = "12"
      jQuery.ajax({ url: eifurl, async: false }).done(function (data) {
        x = jQuery.parseJSON(data)
      });


      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var bl = (x.length+4);
      var name = "EIFTable"
      var values = [
        ["Query", eifurl],
        [name,x.length],
        ["Time", "Value"],
      ];

      var range = sheet.getRange("A1:B3");
      range.values = values;

    	var t = context.workbook.tables.add('A4:B4', false);
      var name = "EIFTable" + $("#rSubSensor").val()
      t.name = name
//      context.workbook.tables.getItem('MyTable').rows.add(null, x );      
      for (var i = 0; i < x.length; i++) {
        context.workbook.tables.getItem(name).rows.add(null, [[x[i][0],x[i][1]]]);      
      }
      
      return context.sync();
    });

  }

})();