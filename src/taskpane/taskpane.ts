/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
};

export async function run() {
  try {
    // This sample creates an image as a Shape object in the worksheet.
    var myFile = document.getElementById("fileUpload") as HTMLInputElement;
    var reader = new FileReader();
    reader.onload = () => {
      Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
      }).catch();
    };

    // Read in the image file as a data URL.
    reader.readAsDataURL(myFile.files[0]);
  } catch (error) {
    console.log(error);
  }
}
