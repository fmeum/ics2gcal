(function() {
  "use strict";

  // Inject empty snackbar into the page if it doesn't already exist
  let snackbar = document.getElementById("snackbar");
  if (!snackbar) {
    snackbar = document.createElement("div");
    snackbar.id = "snackbar";
    document.body.appendChild(snackbar);
  }
  snackbar.innerHTML =
    `<span>Hello World!</span>
     <a href='http://www.google.com' target='_blank'>Undo</a>`;
  setTimeout(function()
    {document.getElementById("snackbar").classList.add('show');}, 2000);
  setTimeout(function()
    {document.getElementById("snackbar").classList.remove('show');}, 7000);

  // Receive messages from the background page
  chrome.runtime.onConnect.addListener(function(port) {
    port.onMessage.addListener(function(message) {
      alert(message);
    });
  });
})();
