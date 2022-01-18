/* global console, document, Excel, Office , window, OfficeRuntime */

import { MeekouConsts } from "../shared/meekouconsts";
import { LoginByCodeInput, LoginOutput, MeekouApi } from "../services/meekouapi";
//var intervalId = null;
// the initialize function must be run each time a new page is loaded
Office.initialize = async () => {
  // window.location.href = "https://web.meekou.cn";
  console.log(123);
  let inputCode: HTMLDivElement = document.getElementById("inputCode") as HTMLDivElement;
  let code = Math.floor(1000 + Math.random() * 9000).toString();
  //generate four random digital
  inputCode.innerText = code;
  login(code);
  // intervalId = window.setInterval(() => {
  //   login(code);
  // }, 1 * 1000);
};

// let inputCode: HTMLDivElement = document.getElementById("inputCode") as HTMLDivElement;
// let code = Math.floor(1000 + Math.random() * 9000).toString();
// //generate four random digital
// inputCode.innerText = code;
// login(code);
// // intervalId = window.setInterval(() => {
// //   login(code);
// // }, 1 * 1000);

async function login(code: string) {
  // check login status
  let meekouApi = new MeekouApi();
  var loginInput = new LoginByCodeInput();
  loginInput.inputCode = code;

  var result: LoginOutput = await meekouApi.loginByCode(loginInput);
  console.log(result);
  //loop unitl access token valid
  while (result.accessToken == null || result.accessToken.length == 0) {
    result = await meekouApi.loginByCode(loginInput);
    await delay(2000);
  }
  console.log(result.accessToken);
  OfficeRuntime.storage.setItem(MeekouConsts.AccessToken, result.accessToken);
  var messageObject = { messageType: "loginSuccess", login: result };
  var jsonMessage = JSON.stringify(messageObject);
  Office.context.ui.messageParent(jsonMessage);
  //   if (result.accessToken.length > 0) {
  //     console.log(result.accessToken);
  //     window.clearInterval(intervalId);
  //   } else {
  //     console.log("wait for scan qr code");
  //   }
}

function delay(ms: number) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}
