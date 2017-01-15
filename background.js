"use strict";
(function() {
  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar/";
  const GAPI_CALENDAR_LIST_COMMAND =
    "https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=";


  function handleStatus(response) {
    if (response.status >= 200 && response.status < 300) {
      return Promise.resolve(response)
    } else {
      return Promise.reject(new Error(response.statusText))
    }
  }

  function toJson(response) {
    return response.json()
  }

  function linkMenuCalendar_onClick(info) {
    alert("calendar " + info.menuItemId.split("/")[1] + " was clicked");
    alert("info: " + JSON.stringify(info));
  }

  function removeContextMenu() {
    // TODO
  }

  function installContextMenu(calendars) {
    chrome.contextMenus.create({
      "id": LINK_MENU_ID,
      "title": "Add to calendar",
      "contexts": ["link"],
      "targetUrlPatterns": [
        "*://*/*.ics",
        "*://*/*.ICS",
        "*://*/*ics_view*"
      ]
    });
    for (let calendar_id in calendars) {
      chrome.contextMenus.create({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendar_id,
        "title": calendars[calendar_id],
        "parentId": LINK_MENU_ID,
        "contexts": ["link"],
        "targetUrlPatterns": [
          "*://*/*.ics",
          "*://*/*.ICS",
          "*://*/*ics_view*"
        ]
      });
    }
    chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
  }

  function fetchCalendars() {
    chrome.identity.getAuthToken({"interactive": false}, function(token) {
      if (chrome.runtime.lastError) {
        alert(chrome.runtime.lastError.message);
        return;
      }

      fetch(GAPI_CALENDAR_LIST_COMMAND + token, {
        method: "GET"
      })
        .then(handleStatus)
        .then(toJson)
        .then(function (response) {
          let calendars = {}
          for (let item of response.items) {
            if (item.accessRole != "owner" || !item.selected)
              continue;
            calendars[item.id] = item.summary;
          }
          installContextMenu(calendars);
        })
        .catch(function (error) {
          alert("Request 'calendarList' failed: " + error);
        });
    });
  }

  chrome.runtime.onInstalled.addListener(fetchCalendars);
  chrome.runtime.onStartup.addListener(fetchCalendars);
})();
