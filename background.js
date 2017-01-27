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

  async function showSnackbar(text, action_label, callback) {
    chrome.tabs.insertCSS(null, {
      file: "snackbar.css"
    });
    chrome.tabs.executeScript(null, {
      file: "snackbar.js"
    });
    await messageActiveTab({
      text,
      action_label
    }, response => {
      if (response && response.clicked) callback();
    });
  }

  async function linkMenuCalendar_onClick(info) {
    const icsLink = info.linkUrl;
    const calendarId = info.menuItemId.split("/")[1];
    let responseText = '';
    try {
      let response = await fetch(icsLink).then(handleStatus);
      responseText = await response.text();
    } catch (error) {
      showSnackbar("Can't fetch iCal file.");
      console.log(error);
      return;
    }
    let events = [];
    try {
      let jcalData = ICAL.parse(responseText);
      events = new ICAL.Component(jcalData).getAllSubcomponents().map(
        component => new ICAL.Event(component));
    } catch (error) {
      showSnackbar("The iCal file has an invalid format.");
      console.log(error);
      return;
    }
    let gcalEvents = [];
    try {
      gcalEvents = await Promise.all(events.map(
        event => createEvent(event, calendarId)));
    } catch (error) {
      if (events.length === 1) {
        showSnackbar("Can't create the event.");
      } else {
        showSnackbar("Can't create the events.");
      }
      console.log(error);
      return;
    }
    if (gcalEvents.length === 0) {
      await showSnackbar("Empty iCal file, no events added.");
    } else if (gcalEvents.length === 1) {
      await showSnackbar("Event added.", "View",
        () => window.open(gcalEvents[0].htmlLink, "_blank"));
    } else {
      await showSnackbar(`${gcalEvents.length} events added.`, "Undo",
        async function() {
          try {
            Promise.all(gcalEvents.map(
              gcalEvent => deleteEvent(calendarId, gcalEvent.id)
            ));
          } catch (error) {
            showSnackbar("Can't delete the events.");
            console.log(error);
          }
        });
    }
  }

  function removeContextMenu() {
    // TODO
  }

  function installContextMenu(calendars, hiddenCalendars) {
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_ID,
      "title": "Add to calendar"
    }, LINK_MENU_CONTEXT));
    for (let [calendarId, calendarTitle] of calendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": calendarTitle,
        "parentId": LINK_MENU_ID
      }, LINK_MENU_CONTEXT));
    }
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_SEPARATOR_ID,
      "type": "separator",
      "parentId": LINK_MENU_ID
    }, LINK_MENU_CONTEXT));
    for (let [calendarId, calendarTitle] of hiddenCalendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": calendarTitle,
        "parentId": LINK_MENU_ID,
      }, LINK_MENU_CONTEXT));
    }
  }

  async function deleteEvent(calendarId, eventId) {
    let token = await chromep.identity.getAuthToken({
      "interactive": false
    });
    return fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${eventId}`, {
          method: "DELETE",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
        })
      .then(handleStatus);
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
    return response.json();
  }

  async function fetchCalendars() {
    let responseCalendarList = null;
    try {
      let token = await chromep.identity.getAuthToken({
        "interactive": false
      });
      let response = await fetch(
          `https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=${token}`, {
            method: "GET"
          })
        .then(handleStatus);
      responseCalendarList = await response.json();
    } catch (error) {
      console.log("Failed to fetch calendars.");
      console.log(error);
      return;
    }
    let calendars = [];
    let hiddenCalendars = [];
    for (let item of responseCalendarList.items) {
      // Only consider calendars in which we can create events
      if (item.accessRole != "owner" && item.accessRole != "writer")
        continue;
      if (item.selected)
        calendars.push([item.id, item.summary]);
      else
        hiddenCalendars.push([item.id, item.summary]);
    }
    calendars.sort((a, b) => a[1].localeCompare(b[1]));
    hiddenCalendars.sort((a, b) => a[1].localeCompare(b[1]));
    installContextMenu(calendars, hiddenCalendars);
  }

  chrome.runtime.onInstalled.addListener(fetchCalendars);
  chrome.runtime.onStartup
    .addListener(fetchCalendars);
  chrome.contextMenus.onClicked.addListener(
    linkMenuCalendar_onClick);
})();
