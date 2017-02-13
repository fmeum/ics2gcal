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

  async function injectSnackbar(tab) {
    let robotoResponse = await fetch(
      'https://fonts.googleapis.com/css?family=Roboto:400,500');
    let robotoCode = await robotoResponse.text();
    // Load styles first to prevent flashes
    await chromep.tabs.insertCSS(tab, {
      code: robotoCode
    });
    await chromep.tabs.insertCSS(tab, {
      file: "snackbar.css"
    });
    await chromep.tabs.executeScript(tab, {
      file: "snackbar.js"
    });
  }

  function showSnackbar(tab, text, action_label, callback) {
    chrome.tabs.sendMessage(tab, {
        text,
        action_label
      },
      response => {
        if (response && response.clicked) callback();
      });
  }

  async function linkMenuCalendar_onClick(info) {
    let tabs = await chromep.tabs.query({
      active: true,
      currentWindow: true
    });
    const activeTab = tabs[0];
    const activeTabId = activeTab.id;
    // Execute in "parallel", messages to the snackbar will be queued
    injectSnackbar(activeTabId);
    const calendarId = info.menuItemId.split("/")[1];
    let responseText = '';
    try {
      let response = await fetch(info.linkUrl).then(handleStatus);
      responseText = await response.text();
    } catch (error) {
      showSnackbar(activeTabId, "Can't fetch iCal file.");
      console.log(error);
      return;
    }
    showSnackbar(activeTabId, "Parsing...");
    let gcalEvents = [];
    try {
      let icalData = ICAL.parse(responseText);
      let icalRoot = new ICAL.Component(icalData);
      // ical.js does not automatically populate its TimezoneService with
      // custom time zones defined in VTIMEZONE components
      let vtimezones = icalRoot.getAllSubcomponents("vtimezone");
      vtimezones.forEach(vtimezone => ICAL.TimezoneService.register(vtimezone));
      let vevents = icalRoot.getAllSubcomponents("vevent");
      gcalEvents = vevents.map(vevent => toGcalEvent(new ICAL.Event(vevent),
        activeTab));
    } catch (error) {
      showSnackbar(activeTabId, "The iCal file has an invalid format.");
      console.log(error);
      return;
    }
    let eventResponses = [];
    try {
      eventResponses = await Promise.all(gcalEvents.map(
        gcalEvent => importEvent(gcalEvent, calendarId)));
    } catch (error) {
      if (gcalEvents.length === 1) {
        showSnackbar(activeTabId, "Can't create the event.");
      } else {
        showSnackbar(activeTabId, "Can't create the events.");
      }
      console.log(error);
      return;
    }
    if (eventResponses.length === 0) {
      showSnackbar(activeTabId, "Empty iCal file, no events imported.");
    } else if (eventResponses.length === 1) {
      showSnackbar(activeTabId, "Event imported.", "View",
        () => window.open(eventResponses[0].htmlLink, "_blank"));
    } else {
      showSnackbar(activeTabId, `${eventResponses.length} events added.`,
        "View all", () => eventResponses.forEach(response => window.open(
          response.htmlLink, "_blank")));
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
    // Show separator only if there are hidden calendars (but even if there are
    // no visible calendars)
    if (hiddenCalendars.length > 0) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_SEPARATOR_ID,
        "type": "separator",
        "parentId": LINK_MENU_ID
      }, LINK_MENU_CONTEXT));
    }
    for (let [calendarId, calendarTitle] of hiddenCalendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": calendarTitle,
        "parentId": LINK_MENU_ID,
      }, LINK_MENU_CONTEXT));
    }
  }

  async function deleteEvent(calendarId, eventId) {
    let token = '';
    try {
      token = await chromep.identity.getAuthToken({
        interactive: false
      });
    } catch (error) {
      updateBrowserAction(false);
      throw error;
    }
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

  function getRecurrenceRules(event) {
    // Note: EXRULE has been deprecated and can lead to ambiguous results
    const RECURRENCE_PROPERTIES = ['rrule', 'rdate', 'exdate', 'exrule'];
    let recurrenceRuleStrings = [];
    for (let recurrenceProperty of RECURRENCE_PROPERTIES) {
      for (let rule of event.component.getAllProperties(recurrenceProperty)) {
        let ruleString = rule.toICALString();
        // Some servers put trailing comma in date lists, which show up as
        // ",--T::" in the ical.js output. As a courtesy, we remove them.
        const TRAILING_COMMA_HACK = /,--T::/g;
        ruleString = ruleString.replace(TRAILING_COMMA_HACK, '');
        recurrenceRuleStrings.push(ruleString);
      }
    }
    return recurrenceRuleStrings;
  }

  function toGcalEvent(event, tabInfo) {
    let gcalEvent = {
      'iCalUID': event.uid,
      'start': {
        'dateTime': event.startDate.toString(),
        'timeZone': event.startDate.zone.toString()
      },
      'end': {
        'dateTime': event.endDate.toString(),
        'timeZone': event.endDate.zone.toString()
      },
      'description': '',
      'reminders': {
        'useDefault': true
      }
    };
    if (event.isRecurring())
      gcalEvent.recurrence = getRecurrenceRules(event);
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
    if (tabInfo) {
      gcalEvent.source = {};
      gcalEvent.source.title = tabInfo.title;
      gcalEvent.source.url = tabInfo.url;
    }
    return gcalEvent;
  }

  async function importEvent(gcalEvent, calendarId) {
    let token = '';
    try {
      token = await chromep.identity.getAuthToken({
        interactive: true
      });
    } catch (error) {
      updateBrowserAction(false);
      throw error;
    }
    updateBrowserAction(true);
    let response = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/import`, {
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

  function updateBrowserAction(active) {
    if (active) {
      chrome.browserAction.setTitle({
        title: "Update calendar list"
      });
      chrome.browserAction.setIcon({
        path: "images/logo_active.png"
      });
    } else {
      chrome.browserAction.setTitle({
        title: "Authorize"
      });
      chrome.browserAction.setIcon({
        path: "images/logo_inactive.png"
      });
    }
  }

  async function fetchCalendars(interactive) {
    let token = "";
    try {
      token = await chromep.identity.getAuthToken({
        interactive
      });
    } catch (error) {
      if (interactive) {
        console.log("Failed to obtain OAuth token interactively.");
        console.log(error);
      }
      updateBrowserAction(false);
      return;
    }
    let responseCalendarList = null;
    try {
      let response = await fetch(
          `https://www.googleapis.com/calendar/v3/users/me/calendarList?access_token=${token}`, {
            method: "GET"
          })
        .then(handleStatus);
      responseCalendarList = await response.json();
    } catch (error) {
      console.log("Failed to fetch calendars.");
      console.log(error);
      updateBrowserAction(false);
      return;
    }
    updateBrowserAction(true);
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

  chrome.runtime.onInstalled.addListener(fetchCalendars.bind(null, false));
  chrome
    .runtime.onStartup.addListener(fetchCalendars.bind(null, false));
  chrome
    .browserAction.onClicked.addListener(fetchCalendars.bind(null, true));
  chrome
    .contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
