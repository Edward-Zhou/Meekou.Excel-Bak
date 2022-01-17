/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { AppConsts } from "../shared/AppConsts";
import { errorHandler } from "../utils/errorHandling";

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
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("C3");
    // range.values = [[5]];
    //get selected range
    //var range = context.workbook.getSelectedRange();
    sheet.getRange("A1").values = [["left"]];
    sheet.getRange("A2").values = [[range.top.toString()]];
    // sheet.getRange("A2").values = [["top" + range.top.toString()]];
    // sheet.getRange("A3").values = [["width" + range.width.toString()]];
    // sheet.getRange("A4").values = [["height" + range.height.toString()]];

    return context.sync();
  }).catch(errorHandler);
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
/**
 * Insert Img to fill cell with Preview
 * @param event
 */
async function InsertImgWithPreview(event: Office.AddinCommands.Event) {
  // dynamic create file input
  let fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.style.display = "none";
  fileInput.accept = "image/*";
  fileInput.onchange = async () => {
    var reader = new FileReader();
    reader.onload = () => {
      Excel.run(async function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        //get selected range
        var range = context.workbook.getSelectedRange();
        range.load({ $all: true });
        await context.sync();
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        //set image location and size
        image.left = range.left;
        image.top = range.top;
        image.width = range.width * 0.99;
        image.height = range.height * 0.99;
        return context.sync();
      }).catch();
    };
    // Read in the image file as a data URL.
    reader.readAsDataURL(fileInput.files[0]);
  };
  fileInput.click();
  event.completed();
}
var loginDialog: Office.Dialog;
async function login() {
  await Office.context.ui.displayDialogAsync(
    `${AppConsts.appBaseUrl}/login.html`,
    { height: 40, width: 20 },
    function (asyncResult) {
      loginDialog = asyncResult.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  loginDialog.close();
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
g.login = login;
