(function() {
  "use strict";
  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_SEPARATOR_ID = "ics2gcal.contextmenu.separator";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar/";
  const LINK_MENU_CONTEXT = {
    "contexts": ["link"],
    "targetUrlPatterns": [
      "*://*/*.ics",
      "*://*/*.ICS",
      "*://*/*ics_view*"
    ]
  };

  function handleStatus(response) {
    if (response.status >= 200 && response.status < 300) {
      return Promise.resolve(response);
    } else {
      return Promise.reject(new Error(response.statusText));
    }
  }

  function linkMenuCalendar_onClick(info) {
    const icsLink = info.linkUrl;
    const calendarId = info.menuItemId.split("/")[1];
    fetch(icsLink)
      .then(handleStatus)
      .then(response => response.text())
      .catch(function (error) {
        alert("Request to fetch .ics failed: " + error);
      })
      .then(ICAL.parse)
      .catch(function (error) {
        alert("The .ics file is invalid: " + error);
      })
      .then(function (jcalData) {
        const vevents = new ICAL.Component(jcalData).getAllSubcomponents();
        for (let vevent of vevents) {
          createEvent(new ICAL.Event(vevent), calendarId);
        }
      });
  }

  function removeContextMenu() {
    // TODO
  }

  function installContextMenu(calendars, hiddenCalendars) {
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_ID,
      "title": "Add to calendar"
    }, LINK_MENU_CONTEXT));
    for (let calendarId in calendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": calendars[calendarId],
        "parentId": LINK_MENU_ID
      }, LINK_MENU_CONTEXT));
    }
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_SEPARATOR_ID,
      "type": "separator",
      "parentId": LINK_MENU_ID
    }, LINK_MENU_CONTEXT));
    for (let calendarId in hiddenCalendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": hiddenCalendars[calendarId],
        "parentId": LINK_MENU_ID,
      }, LINK_MENU_CONTEXT));
    }
  }

  function createEvent(event, calendarId) {
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      let tabUrl = tabs[0].url;
      let gcalEvent = {
        'summary': event.summary,
        'location': event.location,
        'description': `Source: ${tabUrl}`,
        'start': {
          'dateTime': event.startDate.toString(),
          'timeZone': event.startDate.zone.toString()
        },
        'end': {
          'dateTime': event.endDate.toString(),
          'timeZone': event.endDate.zone.toString()
        },
        'reminders': {
          'useDefault': true
        }
      };
      console.log(JSON.stringify(gcalEvent));
      chrome.identity.getAuthToken({"interactive": false}, function(token) {
        if (chrome.runtime.lastError) {
          alert(chrome.runtime.lastError.message);
          return;
        }

        fetch(`https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events`, {
          method: "POST",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
          body: JSON.stringify(gcalEvent)
        });
      });
    });
  }

  function fetchCalendars() {
    chrome.identity.getAuthToken({"interactive": false}, function(token) {
      if (chrome.runtime.lastError) {
        alert(chrome.runtime.lastError.message);
        return;
      }

      fetch(`https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=${token}`, {
        method: "GET"
      })
        .then(handleStatus)
        .then(response => response.json())
        .then(function (response) {
          let calendars = {};
          let hiddenCalendars = {};
          for (let item of response.items) {
            // Only consider calendars in which we can create events
            if (item.accessRole != "owner" && item.accessRole != "writer")
              continue;
            if (item.selected)
              calendars[item.id] = item.summary;
            else
              hiddenCalendars[item.id] = item.summary;
          }
          installContextMenu(calendars, hiddenCalendars);
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
