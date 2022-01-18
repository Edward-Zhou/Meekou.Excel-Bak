/* global console, document, Excel, Office , window, OfficeRuntime */

// the initialize function must be run each time a new page is loaded
Office.initialize = async () => {};
let url: HTMLInputElement = document.getElementById("sourceUrl") as HTMLInputElement;
url.addEventListener("input", loadHtml);
function loadHtml(e: Event) {
  var link = (e.target as HTMLInputElement).value;
  var xhr = new XMLHttpRequest();
  xhr.open("GET", link, true);
  xhr.onload = function () {
    console.log(xhr.responseText);
  };
  xhr.send();
}
// This will be disable for website
// function loadHtml(e: Event) {
//   let webView: HTMLEmbedElement = document.getElementById("webView") as HTMLEmbedElement;
//   webView.setAttribute("src", (e.target as HTMLInputElement).value);
// }
