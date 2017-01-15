(function() {
  "use strict";
  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_SEPARATOR_ID = "ics2gcal.contextmenu.separator";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar/";
  const GAPI_CALENDAR_LIST_COMMAND =
    "https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=";

  function handleStatus(response) {
    if (response.status >= 200 && response.status < 300) {
      return Promise.resolve(response);
    } else {
      return Promise.reject(new Error(response.statusText));
    }
  }

  function toJson(response) {
    return response.json();
  }

  function linkMenuCalendar_onClick(info) {
    const ics_link = info.linkUrl;
    const calendar_id = info.menuItemId.split("/")[1];
    fetch(ics_link)
      .then(handleStatus)
      .then(response => response.text())
      .then(alert)
      .catch(function (error) {
        alert("Request to fetch .ics failed: " + error);
      });
    createEvent(null, calendar_id);
  }

  function removeContextMenu() {
    // TODO
  }

  function installContextMenu(calendars, hidden_calendars) {
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
    chrome.contextMenus.create({
      "id": LINK_MENU_SEPARATOR_ID,
      "type": "separator",
      "parentId": LINK_MENU_ID,
      "contexts": ["link"],
      "targetUrlPatterns": [
        "*://*/*.ics",
        "*://*/*.ICS",
        "*://*/*ics_view*"
      ]
    });
    for (let calendar_id in hidden_calendars) {
      chrome.contextMenus.create({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendar_id,
        "title": hidden_calendars[calendar_id],
        "parentId": LINK_MENU_ID,
        "contexts": ["link"],
        "targetUrlPatterns": [
          "*://*/*.ics",
          "*://*/*.ICS",
          "*://*/*ics_view*"
        ]
      });
    }
  }

  function createEvent(event, calendar_id) {
    chrome.identity.getAuthToken({"interactive": false}, function(token) {
      if (chrome.runtime.lastError) {
        alert(chrome.runtime.lastError.message);
        return;
      }

      fetch(`https://www.googleapis.com/calendar/v3/calendars/${calendar_id}/events`, {
        method: "POST",
        headers: new Headers({
          "Content-Type": "application/json",
          "Authorization": `Bearer ${token}`
        }),
        body: JSON.stringify({
          'summary': 'Google I/O 2015',
          'location': '800 Howard St., San Francisco, CA 94103',
          'description': 'A chance to hear more about Google\'s developer products.',
          'start': {
            'dateTime': '2017-01-28T09:00:00-07:00',
            'timeZone': 'America/Los_Angeles'
          },
          'end': {
            'dateTime': '2017-01-28T17:00:00-07:00',
            'timeZone': 'America/Los_Angeles'
          },
          'reminders': {
            'useDefault': false
          }})
      });
    });
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
          let calendars = {};
          let hidden_calendars = {};
          for (let item of response.items) {
            if (item.accessRole != "owner" && item.accessRole != "writer")
              continue;
            if (item.selected)
              calendars[item.id] = item.summary;
            else
              hidden_calendars[item.id] = item.summary;
          }
          installContextMenu(calendars, hidden_calendars);
        })
        .catch(function (error) {
          alert("Request 'calendarList' failed: " + error);
        });
    });
  }

  chrome.runtime.onInstalled.addListener(fetchCalendars);
  chrome.runtime.onStartup.addListener(fetchCalendars);
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
