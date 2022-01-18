/* global console, document, Excel, Office , window, OfficeRuntime */

import { DialogEventArg, DialogInput } from "../shared/dialogInput";

let dialogInput: DialogInput;

// the initialize function must be run each time a new page is loaded
Office.initialize = async () => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, dialogMessageFromParent);
};
var newlabel = document.createElement("Label");
newlabel.innerHTML = "Here goes the text";
document.getElementById("messages").appendChild(newlabel);
function dialogMessageFromParent(arg: any) {
  var newlabel = document.createElement("Label");
  newlabel.innerHTML = "Here goes the text";
  document.getElementById("messages").appendChild(newlabel);
  dialogInput = JSON.parse(arg.message) as DialogInput;
  document.getElementById(dialogInput.name).style.display = "inline";
}
