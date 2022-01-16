/* global console, document, Excel, Office */

// the initialize function must be run each time a new page is loaded
Office.initialize = () => {
    // window.location.href = "https://web.meekou.cn";
    console.log(123);
    let inputCode: HTMLLabelElement = document.getElementById("inputCode") as HTMLLabelElement;
    inputCode.innerText = "8520";
};
