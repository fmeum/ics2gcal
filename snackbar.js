(function() {
  "use strict";

  // Inject empty snackbar into the page
  var iframe  = document.createElement("iframe");
  iframe.src  = chrome.extension.getURL("snackbar.html");
  document.body.insertBefore(iframe, document.body.firstChild);

  // Receive messages from the background page
  chrome.runtime.onConnect.addListener(function(port) {
    port.onMessage.addListener(function(message) {
      alert(message);
    });
  });
})();
