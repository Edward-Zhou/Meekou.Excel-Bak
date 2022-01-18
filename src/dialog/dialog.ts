/* global console, document, Excel, Office , window, OfficeRuntime */

import { DialogEventArg, DialogInput } from "../shared/dialoginput";

let dialogInput: DialogInput;

// the initialize function must be run each time a new page is loaded
Office.initialize = async () => {
  document.getElementById("send").addEventListener("click", (e:Event) => {
    Office.context.ui.messageParent("Hello World");
  });
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, dialogMessageFromParent);
};

function dialogMessageFromParent(arg: any) {
  dialogInput = JSON.parse(arg.message) as DialogInput;
  document.getElementById(dialogInput.name).style.display = "inline";
}
