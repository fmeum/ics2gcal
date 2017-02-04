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
    let robotoResponse = await fetch(
      'https://fonts.googleapis.com/css?family=Roboto:400,500');
    let robotoCode = await robotoResponse.text();
    chrome.tabs.insertCSS(null, {
      code: robotoCode
    });
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
    let gcalEvents = [];
    try {
      let jcalData = ICAL.parse(responseText);
      let jcalComponents = new ICAL.Component(jcalData).getAllSubcomponents();
      gcalEvents = await Promise.all(jcalComponents.map(
        component => toGcalEvent(new ICAL.Event(component))));
    } catch (error) {
      showSnackbar("The iCal file has an invalid format.");
      console.log(error);
      return;
    }
    let eventResponses = [];
    try {
      eventResponses = await Promise.all(gcalEvents.map(
        gcalEvent => createEvent(gcalEvent, calendarId)));
    } catch (error) {
      if (gcalEvents.length === 1) {
        showSnackbar("Can't create the event.");
      } else {
        showSnackbar("Can't create the events.");
      }
      console.log(error);
      return;
    }
    if (eventResponses.length === 0) {
      await showSnackbar("Empty iCal file, no events added.");
    } else if (eventResponses.length === 1) {
      await showSnackbar("Event added.", "View",
        () => window.open(eventResponses[0].htmlLink, "_blank"));
    } else {
      await showSnackbar(`${eventResponses.length} events added.`, "Undo",
        async function() {
          try {
            Promise.all(eventResponses.map(
              gcalEvent => deleteEvent(calendarId, gcalEvent.id)
            ));
          } catch (error) {
            showSnackbar("Can't delete the events.");
            console.log(error);
          }
        });
    }
  }

  async function installContextMenu(calendars, hiddenCalendars) {
    await chromep.contextMenus.removeAll();
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

  async function toGcalEvent(event) {
    let gcalEvent = {
      'start': {
        'dateTime': event.startDate.toString(),
        'timeZone': event.startDate.zone.toString()
      },
      'description': '',
      'reminders': {
        'useDefault': true
      }
    };
    if (event.summary)
      gcalEvent.summary = event.summary;
    if (event.location)
      gcalEvent.location = event.location;
    if (event.description)
      gcalEvent.description = event.description;
    let url = event.component.getFirstPropertyValue('url');
    if (url)
      if (gcalEvent.description)
        gcalEvent.description += `\n\n${url}`;
      else
        gcalEvent.description += url;
    if (event.hasOwnProperty('endDate')) {
      gcalEvent.end = {
        'dateTime': event.endDate.toString(),
        'timeZone': event.endDate.zone.toString()
      };
    } else {
      // If there is no end date, we assume a duration of 1h
      gcalEvent.end = {
        'dateTime': event.startDate.adjust(0, 1, 0, 0).toString(),
        'timeZone': event.startDate.zone.toString()
      };
    }
    let tabs = await chromep.tabs.query({
      active: true,
      currentWindow: true
    });
    let tabUrl = tabs[0].url;
    if (tabUrl !== url)
      gcalEvent.description += `\n\nAdded from: ${tabUrl}`;
    return gcalEvent;
  }

  async function createEvent(gcalEvent, calendarId) {
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
    await installContextMenu(calendars, hiddenCalendars);
  }

  chrome.runtime.onInstalled.addListener(fetchCalendars);
  chrome.runtime.onStartup.addListener(fetchCalendars);
  chrome.browserAction.onClicked.addListener(fetchCalendars);
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
