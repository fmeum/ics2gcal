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

  function parseIcs(icsContent) {
    // Some servers put trailing comma in date lists, which trip up ical.js
    // output. As a courtesy we remove trailing commas as long as they do not
    // precede a folded line.
    // https://regex101.com/r/Q2VEZB/1
    const TRAILING_COMMA_PATTERN = /,\n(\S)/g;
    responseText = responseText.replace(TRAILING_COMMA_PATTERN, `\n$1`);
    let icalData = ICAL.parse(icsContent);
    let icalRoot = new ICAL.Component(icalData);
    // ical.js does not automatically populate its TimezoneService with
    // custom time zones defined in VTIMEZONE components
    icalRoot.getAllSubcomponents("vtimezone").forEach(
      vtimezone => ICAL.TimezoneService.register(vtimezone));
    return icalRoot.getAllSubcomponents("vevent").map(
      vevent => toGcalEvent(new ICAL.Event(vevent), activeTab));
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
    let gcalEventsAndExDates = [];
    try {
      gcalEventsAndExDates = parseIcs(responseText);
    } catch (error) {
      showSnackbar(activeTabId, "The iCal file has an invalid format.");
      console.log(error);
      return;
    }
    let eventResponses = [];
    try {
      eventResponses = await Promise.all(gcalEventsAndExDates.map(
        async function(gcalEventAndExDates) {
          let [gcalEvent, exDates] = gcalEventAndExDates;
          let event = await importEvent(gcalEvent, calendarId);
          await cancelExDates(calendarId, event, exDates);
          return event;
        }));
    } catch (error) {
      if (gcalEventsAndExDates.length === 1) {
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
        "View all",
        () => eventResponses.forEach(
          response => window.open(response.htmlLink, "_blank")));
    }
  }

  async function installContextMenu(calendars, hiddenCalendars) {
    await chromep.contextMenus.removeAll();
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_ID,
      "title": "Add to calendar"
    }, LINK_MENU_CONTEXT));
    calendars.sort((a, b) => a[1].localeCompare(b[1]));
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
    hiddenCalendars.sort((a, b) => a[1].localeCompare(b[1]));
    for (let [calendarId, calendarTitle] of hiddenCalendars) {
      chrome.contextMenus.create(Object.assign({
        "id": LINK_MENU_CALENDAR_ID_PREFIX + calendarId,
        "title": calendarTitle,
        "parentId": LINK_MENU_ID,
      }, LINK_MENU_CONTEXT));
    }
  }

  function authenticate(interactive) {
    let token = '';
    try {
      token = chromep.identity.getAuthToken({
        interactive
      });
    } catch (error) {
      updateBrowserAction(false);
      throw error;
    }
    updateBrowserAction(true);
    return token;
  }

  async function cancelExDates(calendarId, event, exDates) {
    let token = await authenticate(false);
    await Promise.all(exDates.map(async function(exDate) {
      let timeString = exDate.toString();
      let instances = [];
      // We may retry fetching instances for exDate
      let retry = false;
      while (true) {
        instances = await fetch(
            `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${event.id}/instances?originalStart=${timeString}`, {
              method: "GET",
              headers: new Headers({
                "Content-Type": "application/json",
                "Authorization": `Bearer ${token}`
              }),
            })
          .then(handleStatus)
          .then(response => response.json());
        if (instances.items.length === 0) {
          if (retry) {
            // The corrected exDate also didn't match and we give up. This will
            // also be triggered if we import an event a second time.
            return;
          } else if (exDate.timezone === "Z" && event.start.hasOwnProperty(
              'dateTime')) {
            // If the exDate is specified in UTC and is not an all-day event,
            // we try again using the time zone in event.start.
            retry = true;
            let startUtcOffset = ICAL.VCardTime.fromDateAndOrTimeString(
              event.start.dateTime).zone;
            exDate.adjust(
              /* d */ 0, /* h */ 0, /* m */ 0, -startUtcOffset.toSeconds());
            timeString = exDate.toString();
          }
        } else {
          // The given exDate matches at least one instance of the recurrent
          // event.
          break;
        }
      };
      await Promise.all(instances.items.map(instance => {
        instance.status = "cancelled";
        return fetch(
            `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${instance.id}`, {
              method: "PUT",
              headers: new Headers({
                "Content-Type": "application/json",
                "Authorization": `Bearer ${token}`
              }),
              body: JSON.stringify(instance),
            })
          .then(handleStatus);
      }));
    }));
  }

  function getRecurrenceRules(event) {
    // Rebind to the top-level component of the event
    let component = event.component;
    // Note: EXRULE has been deprecated and can lead to ambiguous results, so we
    // don't support it. EXDATE is treated specially. If RDATE is present, the
    // resulting recurrent event in the Google Calendar will be broken in the
    // sense that the instances can't be edited all at once (they can still be
    // deleted together).
    let recurrenceRuleStrings = [];
    for (let recurrenceProperty of ['rrule', 'rdate']) {
      for (let i = 0; i < component.getAllProperties(recurrenceProperty).length; i++) {
        let ruleString = component.getAllProperties(recurrenceProperty)[i].toICALString();
        if (recurrenceProperty === 'rrule' || recurrenceProperty ===
          'rdate')
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
    let exDates = [];
    if (event.isRecurring()) {
      gcalEvent.recurrence = getRecurrenceRules(event);
      var expanded = new ICAL.RecurExpansion({
        component: event.component,
        dtstart: event.component.getFirstPropertyValue('dtstart')
      });
      exDates = expanded.exDates;
    }
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
    return [gcalEvent, exDates];
  }

  async function importEvent(gcalEvent, calendarId) {
    let token = await authenticate(true);
    let response = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/import`, {
          method: "POST",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
          body: JSON.stringify(gcalEvent)
        })
      .then(handleStatus)
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
    let token = await authenticate(interactive);
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
    await installContextMenu(calendars, hiddenCalendars);
  }

  chrome.runtime.onInstalled.addListener(fetchCalendars.bind(null, false));
  chrome.runtime.onStartup.addListener(fetchCalendars.bind(null, false));
  chrome.browserAction.onClicked.addListener(fetchCalendars.bind(null, true));
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
