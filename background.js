(function() {
  "use strict";

  const chromep = new ChromePromise();

  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_SEPARATOR_ID = "ics2gcal.contextmenu.separator";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar/";
  const LINK_MENU_CONTEXT = {
    "contexts": ["link"],
    "targetUrlPatterns": [
      "<all_urls>"
    ]
  };

  function handleStatus(response) {
    if (response.status >= 200 && response.status < 300) {
      return Promise.resolve(response);
    } else {
      return Promise.reject(new Error(response.statusText));
    }
  }

  // Promisify setTimeout
  function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async function messageActiveTab(request, callback) {
    let tabs = await chromep.tabs.query({
      active: true,
      currentWindow: true
    });
    let activeTabId = tabs[0].id;
    chrome.tabs.sendMessage(activeTabId, request, callback);
  }

  async function linkMenuCalendar_onClick(info) {
    const icsLink = info.linkUrl;
    const calendarId = info.menuItemId.split("/")[1];
    let jcalData = '';
    try {
      let response = await fetch(icsLink).then(handleStatus);
      let responseText = await response.text();
      jcalData = ICAL.parse(responseText);
    } catch (error) {
      alert(`Request to fetch .ics failed:\n${error.stack}`);
      return;
    }
    let eventIds = [];
    try {
      let vevents = new ICAL.Component(jcalData).getAllSubcomponents();
      eventIds = await Promise.all(vevents.map(
        vevent => createEvent(new ICAL.Event(vevent), calendarId)));
    } catch (error) {
      alert(`The .ics file is invalid:\n${error.stack}`);
      return;
    }
    chrome.tabs.insertCSS(null, {
      file: "snackbar.css"
    });
    chrome.tabs.executeScript(null, {
      file: "snackbar.js"
    });
    await timeout(2000);
    console.log("Message tab: Event added");
    await messageActiveTab({
        "text": "Event addded",
        "action_label": "View"
      },
      function(response) {
        if (response.clicked) console.log(`View ${eventIds[0]}`);
      });
    await timeout(2000);
    console.log("Message tab: Multiple events added");
    await messageActiveTab({
        "text": "Multiple events addded",
        "action_label": "Undo"
      },
      function(response) {
        if (response.clicked) console.log(`Undo ${eventIds[0]}`);
      });
    await timeout(11000);
    console.log("Message tab: ICS file invalid");
    await messageActiveTab({
      "text": "ICS file invalid"
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

  async function createEvent(event, calendarId) {
    let tabs = await chromep.tabs.query({
      active: true,
      currentWindow: true
    });
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
    try {
      let token = await chromep.identity.getAuthToken({
        "interactive": false
      });
      let response = await fetch(
          `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events`, {
            method: "POST",
            headers: new Headers({
              "Content-Type": "application/json",
              "Authorization": `Bearer ${token}`
            }),
            body: JSON.stringify(gcalEvent)
          })
        .then(handleStatus);
      let responseEvent = await response.json();
      return responseEvent.id;
    } catch (error) {
      alert(`Request 'events' failed:\n${error.stack}`);
      return;
    }
  }

  async function fetchCalendars() {
    try {
      let token = await chromep.identity.getAuthToken({
        "interactive": false
      });
      let response = await fetch(
          `https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=${token}`, {
            method: "GET"
          })
        .then(handleStatus);
      let responseCalendarList = await response.json();
      let calendars = {};
      let hiddenCalendars = {};
      for (let item of responseCalendarList.items) {
        // Only consider calendars in which we can create events
        if (item.accessRole != "owner" && item.accessRole != "writer")
          continue;
        if (item.selected)
          calendars[item.id] = item.summary;
        else
          hiddenCalendars[item.id] = item.summary;
      }
      installContextMenu(calendars, hiddenCalendars);
    } catch (error) {
      alert(`Request 'calendarList' failed:\n${error.stack}`);
      return;
    }

  }

  chrome.runtime.onInstalled.addListener(fetchCalendars);
  chrome.runtime.onStartup.addListener(fetchCalendars);
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
