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

  let calendarIdToTitle = {};
  let mutexLinkMenuCalendar_onClick = false;

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
      response => callback ? callback(response.clicked) : {}
    );
  }

  function parseIcs(icsContent) {
    // Some servers put trailing comma in date lists, which trip up ical.js
    // output. As a courtesy we remove trailing commas as long as they do not
    // precede a folded line.
    // https://regex101.com/r/Q2VEZB/1
    const TRAILING_COMMA_PATTERN = /,\n(\S)/g;
    icsContent = icsContent.replace(TRAILING_COMMA_PATTERN, `\n$1`);
    // We also remove empty RDATE and EXDATE properties
    // https://regex101.com/r/5tWWwt/1
    const EMPTY_PROPERTY_PATTERN = /^(RDATE|EXDATE):$/gm;
    icsContent = icsContent.replace(EMPTY_PROPERTY_PATTERN, '');
    let icalData = ICAL.parse(icsContent);
    let icalRoot = new ICAL.Component(icalData);
    // ical.js does not automatically populate its TimezoneService with
    // custom time zones defined in VTIMEZONE components
    icalRoot.getAllSubcomponents("vtimezone").forEach(
      vtimezone => ICAL.TimezoneService.register(vtimezone));
    return icalRoot.getAllSubcomponents("vevent").map(
      vevent => new ICAL.Event(vevent));
  }

  async function linkMenuCalendar_onClick(info) {
    // Temporarily ignore clicks while we process one
    if (mutexLinkMenuCalendar_onClick)
      return;
    mutexLinkMenuCalendar_onClick = true;
    let tabs = await chromep.tabs.query({
      active: true,
      currentWindow: true
    });
    const activeTab = tabs[0];
    const activeTabId = activeTab.id;
    await injectSnackbar(activeTabId);
    // If the user disables and then re-enables the app, the context menus will
    // show up magically without installContextMenus ever having been called.
    // In this case, we have to populate the list of calendars here.
    if (Object.keys(calendarIdToTitle).length === 0)
      await fetchCalendars(true);
    const calendarTitle = info.menuItem;
    const calendarId = info.menuItemId.split("/")[1];
    let responseText = '';
    try {
      let response = await fetch(info.linkUrl).then(handleStatus);
      responseText = await response.text();
    } catch (error) {
      showSnackbar(activeTabId, "Can't fetch iCal file.");
      console.log(error);
      mutexLinkMenuCalendar_onClick = false;
      return;
    }
    let gcalEventsAndExDates = [];
    try {
      gcalEventsAndExDates = parseIcs(responseText).map(
        event => toGcalEvent(event, activeTab));
    } catch (error) {
      showSnackbar(activeTabId, "The iCal file has an invalid format.");
      console.log(error);
      mutexLinkMenuCalendar_onClick = false;
      return;
    }
    let importMessage = '';
    if (gcalEventsAndExDates.length === 0) {
      showSnackbar(activeTabId, "Empty iCal file, no events imported.");
      mutexLinkMenuCalendar_onClick = false;
      return;
    } else if (gcalEventsAndExDates.length === 1) {
      importMessage =
        `Importing event into '${calendarIdToTitle[calendarId]}'...`;
    } else {
      importMessage =
        `Importing ${gcalEventsAndExDates.length} events into '${calendarIdToTitle[calendarId]}'...`;
    }
    // As we implement cancelling via lazy execution, we will notify the user of
    // possible event loss if they try to leave the page while we haven't
    // committed.
    window.addEventListener('onbeforeunload', e => true);
    await chromep.tabs.executeScript(activeTabId, {
      code: "window.onbeforeunload = e => true;"
    });
    showSnackbar(activeTabId, importMessage, "Cancel", async function(clicked) {
      // Use an asynchronous closure as replacement for RAII
      await async function() {
        if (!clicked) {
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
            // TODO
          } else if (eventResponses.length === 1) {
            showSnackbar(activeTabId,
              `Event imported into '${calendarIdToTitle[calendarId]}'.`,
              "View",
              clicked => clicked ? window.open(eventResponses[0].htmlLink,
                "_blank") : {});
          } else {
            showSnackbar(activeTabId,
              `${eventResponses.length} events imported into '${calendarIdToTitle[calendarId]}'.`,
              "View all",
              clicked => clicked ? eventResponses.forEach(response =>
                window.open(response.htmlLink, "_blank")) : {});
          }
        }
      }();
      mutexLinkMenuCalendar_onClick = false;
      await chromep.tabs.executeScript(activeTabId, {
        code: "window.onbeforeunload = e => null;"
      });
    });
  }

  async function installContextMenu(calendars, hiddenCalendars) {
    calendarIdToTitle = {};
    chrome.contextMenus.create(Object.assign({
      "id": LINK_MENU_ID,
      "title": "Add to calendar"
    }, LINK_MENU_CONTEXT));
    calendars.sort((a, b) => a[1].localeCompare(b[1]));
    for (let [calendarId, calendarTitle] of calendars) {
      calendarIdToTitle[calendarId] = calendarTitle;
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
      calendarIdToTitle[calendarId] = calendarTitle;
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
    if (interactive)
      updateBrowserAction(true);
    return token;
  }

  async function fetchInstancesWithOriginalStart(token, calendarId, eventId,
    icalDate) {
    const timeString = icalDate.toString();
    let response = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${eventId}/instances?originalStart=${encodeURIComponent(timeString)}`, {
          method: "GET",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
        })
      .then(handleStatus);
    let responseJson = await response.json();
    return responseJson.items;
  }

  async function fetchInstancesWithinBounds(token, calendarId, event,
    startTimeMin, startTimeMax) {
    let startTimeMaxUtc = ICAL.Timezone.convert_time(startTimeMax,
      startTimeMax.zone,
      ICAL.Timezone.utcTimezone);
    // The Google Calendar API demands that we provide an explicit UTC
    // offset.
    let startTimeMaxUtcString = startTimeMaxUtc.toString().replace('Z', '+00:00');
    // The 'timeMin' parameter used by the Google Calendar API gives a lower
    // bound on the end time, not the start time. We thus have to shift by the
    // duration of the event first.
    // TODO: This only works if start and end are specified in the same time
    // zone since fromDateTimeString ignores the UTC offset.
    let eventDuration = ICAL.Time.fromDateTimeString(event.end.dateTime).subtractDateTz(
      ICAL.Time.fromDateTimeString(event.start.dateTime));
    let endTimeMin = startTimeMin.clone();
    endTimeMin.addDuration(eventDuration);
    let endTimeMinUtc = ICAL.Timezone.convert_time(endTimeMin, endTimeMin.zone,
      ICAL.Timezone.utcTimezone);
    let endTimeMinUtcString = endTimeMinUtc.toString().replace('Z', '+00:00');
    let response = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${event.id}/instances?timeMin=${encodeURIComponent(endTimeMinUtcString)}&timeMax=${encodeURIComponent(startTimeMaxUtcString)}`, {
          method: "GET",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
        })
      .then(handleStatus);
    let responseJson = await response.json();
    return responseJson.items;
  }
  async function cancelInstance(token, calendarId, instance) {
    instance.status = "cancelled";
    await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calendarId}/events/${instance.id}`, {
          method: "PUT",
          headers: new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${token}`
          }),
          body: JSON.stringify(instance),
        })
      .then(handleStatus);
  }

  async function cancelExDates(calendarId, event, exDates) {
    let token = await authenticate(false);
    await Promise.all(exDates.map(async function(exDate) {
      let instances = [];
      instances = await fetchInstancesWithOriginalStart(token,
        calendarId, event.id, exDate);
      if (instances.length === 0 && event.start.hasOwnProperty(
          'dateTime')) {
        // If the event is not an all-day event and we get no exact match
        // for the exDate, we check whether there is a single instance on
        // the day described by exDate (with respect to exDate's time zone).
        // If there is a single instance on this day, we assume that the
        // exDate's time is off and this instance should be cancelled. This
        // appears to be Outlook's standard behavior.
        let startDay = exDate.clone();
        startDay.adjust(/* d */0, -exDate.hour, -exDate.minute, -exDate.second);
        let endDay = startDay.clone();
        endDay.adjust(1, 0, 0, 0);
        instances = await fetchInstancesWithinBounds(token, calendarId, event,
          startDay, endDay);
        // If we still get zero matches or match more than one event, we give up
        // and silently ignore this exDate.
        if (instances.length !== 1)
          return;
      }
      // The iCalendar specification says that duplicate instances must not be
      // generated, so strictly speaking this .all should not be necessary.
      await Promise.all(instances.map(instance => cancelInstance(token,
        calendarId, instance)));
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
      const expanded = new ICAL.RecurExpansion({
        component: event.component,
        dtstart: event.component.getFirstPropertyValue('dtstart'),
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
        }).then(handleStatus);
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
    await chromep.contextMenus.removeAll();
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
      if (item.accessRole !== 'owner' && item.accessRole !== 'writer')
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
