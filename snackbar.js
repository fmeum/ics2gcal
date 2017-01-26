(function() {
  "use strict";

  // Only run the injection script once per page
  if (window.hasBeenInjected) return;
  window.hasBeenInjected = true;

  const SNACKBAR_ID = "ics2gcal-snackbar";
  const SNACKBAR_TEXT_ID = "ics2gcal-snackbar-text";
  const SNACKBAR_ACTION_ID = "ics2gcal-snackbar-action";
  const SNACKBAR_TIMEOUT_MS = 10000;

  // Inject empty snackbar into the page if it doesn't already exist
  let snackbar = document.getElementById(SNACKBAR_ID);
  if (!snackbar) {
    snackbar = document.createElement("div");
    snackbar.id = SNACKBAR_ID;
    snackbar.innerHTML =
      `<span id="${SNACKBAR_TEXT_ID}"></span>
       <a href="javascript:void(0);" target="_blank" id="${SNACKBAR_ACTION_ID}"></a>`;
    document.body.appendChild(snackbar);
  }
  let snackbar_text = document.getElementById(SNACKBAR_TEXT_ID);
  let snackbar_action = document.getElementById(SNACKBAR_ACTION_ID);

  // Receive requests from the background page to show snackbar
  chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
    let {
      text,
      action_label
    } = request;
    let hideSnackbar = function() {
      snackbar.classList.remove("show");
    };
    let showSnackbar = function() {
      snackbar_text.innerText = text;
      if (action_label) {
        snackbar_action.innerText = action_label;
        snackbar_action.style.display = "initial";
        snackbar_action.onclick = function() {
          sendResponse({clicked: "true"});
        };
      } else {
        snackbar_action.style.display = "none";
      }
      snackbar.classList.add("show");
      window.setTimeout(hideSnackbar, SNACKBAR_TIMEOUT_MS);
    };
    // Ensure that the old snackbar fades out before the new one fades in
    if (snackbar.classList.contains("show")) {
      console.log("Old snackbar present, fade it out first");
      snackbar.classList.remove("show");
      // TODO: Use ontransitionend here
      window.setTimeout(showSnackbar, 300);
      console.log("Old snackbar fades out");
    } else {
      console.log("No old snackbar, just fade this one in");
      showSnackbar();
    }
    // Listener has to return true if sendResponse will be called asynchronously
    if (action_label) return true;
  });
})();
