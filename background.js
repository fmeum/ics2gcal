"use strict";
(function() {
  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar";

  function linkMenuCalendar_onClick(info) {
    alert("item " + info.menuItemId + " was clicked");
    alert("info: " + JSON.stringify(info));
  }

  function showContextMenu() {
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
    chrome.contextMenus.create({
      "id": LINK_MENU_CALENDAR_ID_PREFIX,
      "title": "Conferences",
      "parentId": LINK_MENU_ID,
      "contexts": ["link"],
      "targetUrlPatterns": [
        "*://*/*.ics",
        "*://*/*.ICS",
        "*://*/*ics_view*"
      ]
    });
  }

  chrome.runtime.onInstalled.addListener(showContextMenu);
  chrome.runtime.onStartup.addListener(showContextMenu);
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
