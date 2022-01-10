/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window, document, Excel */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});
Office.initialize = () => {};
var _count = 0;
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  // Your code goes here.
  _count++;
  Office.addin.showAsTaskpane();
  document.getElementById("run").textContent = "Go" + _count;

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
/**
 * Insert Img to fill cell with Preview
 * @param event
 */
function InsertImgWithPreview(event: Office.AddinCommands.Event) {
  // dynamic create file input
  let fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.style.display = "none";
  fileInput.onchange = async () => {
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
    reader.readAsDataURL(fileInput.files[0]);
  };
  fileInput.click();
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.InsertImgWithPreview = InsertImgWithPreview;
