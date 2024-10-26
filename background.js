import ICAL from 'ical.js';
import UUID from 'pure-uuid';
(function() {
  const LINK_MENU_ID = "ics2gcal.contextmenu.link";
  const LINK_MENU_SEPARATOR_ID = "ics2gcal.contextmenu.separator";
  const LINK_MENU_CALENDAR_ID_PREFIX = "ics2gcal.contextmenu.link.calendar/";
  const LINK_MENU_CONTEXT = {
    "contexts": ["link"],
    "targetUrlPatterns": [
      "<all_urls>"
    ]
  };

  const WINDOWS_TO_IANA_TIMEZONES = {
    "Dateline Standard Time": "Etc/GMT+12",
    "UTC-11": "Etc/GMT+11",
    "Aleutian Standard Time": "America/Adak",
    "Hawaiian Standard Time": "Pacific/Honolulu",
    "Marquesas Standard Time": "Pacific/Marquesas",
    "Alaskan Standard Time": "America/Anchorage",
    "UTC-09": "Etc/GMT+9",
    "Pacific Standard Time (Mexico)": "America/Tijuana",
    "UTC-08": "Etc/GMT+8",
    "Pacific Standard Time": "America/Los_Angeles",
    "US Mountain Standard Time": "America/Phoenix",
    "Mountain Standard Time (Mexico)": "America/Chihuahua",
    "Mountain Standard Time": "America/Denver",
    "Central America Standard Time": "America/Guatemala",
    "Central Standard Time": "America/Chicago",
    "Easter Island Standard Time": "Pacific/Easter",
    "Central Standard Time (Mexico)": "America/Mexico_City",
    "Canada Central Standard Time": "America/Regina",
    "SA Pacific Standard Time": "America/Bogota",
    "Eastern Standard Time (Mexico)": "America/Cancun",
    "Eastern Standard Time": "America/New_York",
    "Haiti Standard Time": "America/Port-au-Prince",
    "Cuba Standard Time": "America/Havana",
    "US Eastern Standard Time": "America/Indianapolis",
    "Paraguay Standard Time": "America/Asuncion",
    "Atlantic Standard Time": "America/Halifax",
    "Venezuela Standard Time": "America/Caracas",
    "Central Brazilian Standard Time": "America/Cuiaba",
    "SA Western Standard Time": "America/La_Paz",
    "Pacific SA Standard Time": "America/Santiago",
    "Turks And Caicos Standard Time": "America/Grand_Turk",
    "Newfoundland Standard Time": "America/St_Johns",
    "Tocantins Standard Time": "America/Araguaina",
    "E. South America Standard Time": "America/Sao_Paulo",
    "SA Eastern Standard Time": "America/Cayenne",
    "Argentina Standard Time": "America/Buenos_Aires",
    "Greenland Standard Time": "America/Godthab",
    "Montevideo Standard Time": "America/Montevideo",
    "Magallanes Standard Time": "America/Punta_Arenas",
    "Saint Pierre Standard Time": "America/Miquelon",
    "Bahia Standard Time": "America/Bahia",
    "UTC-02": "Etc/GMT+2",
    "Azores Standard Time": "Atlantic/Azores",
    "Cape Verde Standard Time": "Atlantic/Cape_Verde",
    "UTC": "Etc/GMT",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central Europe Standard Time": "Europe/Budapest",
    "Romance Standard Time": "Europe/Paris",
    "Morocco Standard Time": "Africa/Casablanca",
    "Sao Tome Standard Time": "Africa/Sao_Tome",
    "Central European Standard Time": "Europe/Warsaw",
    "W. Central Africa Standard Time": "Africa/Lagos",
    "Jordan Standard Time": "Asia/Amman",
    "GTB Standard Time": "Europe/Bucharest",
    "Middle East Standard Time": "Asia/Beirut",
    "Egypt Standard Time": "Africa/Cairo",
    "E. Europe Standard Time": "Europe/Chisinau",
    "Syria Standard Time": "Asia/Damascus",
    "West Bank Standard Time": "Asia/Hebron",
    "South Africa Standard Time": "Africa/Johannesburg",
    "FLE Standard Time": "Europe/Kiev",
    "Israel Standard Time": "Asia/Jerusalem",
    "Kaliningrad Standard Time": "Europe/Kaliningrad",
    "Sudan Standard Time": "Africa/Khartoum",
    "Libya Standard Time": "Africa/Tripoli",
    "Namibia Standard Time": "Africa/Windhoek",
    "Arabic Standard Time": "Asia/Baghdad",
    "Turkey Standard Time": "Europe/Istanbul",
    "Arab Standard Time": "Asia/Riyadh",
    "Belarus Standard Time": "Europe/Minsk",
    "Russian Standard Time": "Europe/Moscow",
    "E. Africa Standard Time": "Africa/Nairobi",
    "Iran Standard Time": "Asia/Tehran",
    "Arabian Standard Time": "Asia/Dubai",
    "Astrakhan Standard Time": "Europe/Astrakhan",
    "Azerbaijan Standard Time": "Asia/Baku",
    "Russia Time Zone 3": "Europe/Samara",
    "Mauritius Standard Time": "Indian/Mauritius",
    "Saratov Standard Time": "Europe/Saratov",
    "Georgian Standard Time": "Asia/Tbilisi",
    "Caucasus Standard Time": "Asia/Yerevan",
    "Afghanistan Standard Time": "Asia/Kabul",
    "West Asia Standard Time": "Asia/Tashkent",
    "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
    "Pakistan Standard Time": "Asia/Karachi",
    "India Standard Time": "Asia/Calcutta",
    "Sri Lanka Standard Time": "Asia/Colombo",
    "Nepal Standard Time": "Asia/Katmandu",
    "Central Asia Standard Time": "Asia/Almaty",
    "Bangladesh Standard Time": "Asia/Dhaka",
    "Omsk Standard Time": "Asia/Omsk",
    "Myanmar Standard Time": "Asia/Rangoon",
    "SE Asia Standard Time": "Asia/Bangkok",
    "Altai Standard Time": "Asia/Barnaul",
    "W. Mongolia Standard Time": "Asia/Hovd",
    "North Asia Standard Time": "Asia/Krasnoyarsk",
    "N. Central Asia Standard Time": "Asia/Novosibirsk",
    "Tomsk Standard Time": "Asia/Tomsk",
    "China Standard Time": "Asia/Shanghai",
    "North Asia East Standard Time": "Asia/Irkutsk",
    "Singapore Standard Time": "Asia/Singapore",
    "W. Australia Standard Time": "Australia/Perth",
    "Taipei Standard Time": "Asia/Taipei",
    "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
    "Aus Central W. Standard Time": "Australia/Eucla",
    "Transbaikal Standard Time": "Asia/Chita",
    "Tokyo Standard Time": "Asia/Tokyo",
    "North Korea Standard Time": "Asia/Pyongyang",
    "Korea Standard Time": "Asia/Seoul",
    "Yakutsk Standard Time": "Asia/Yakutsk",
    "Cen. Australia Standard Time": "Australia/Adelaide",
    "AUS Central Standard Time": "Australia/Darwin",
    "E. Australia Standard Time": "Australia/Brisbane",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "West Pacific Standard Time": "Pacific/Port_Moresby",
    "Tasmania Standard Time": "Australia/Hobart",
    "Vladivostok Standard Time": "Asia/Vladivostok",
    "Lord Howe Standard Time": "Australia/Lord_Howe",
    "Bougainville Standard Time": "Pacific/Bougainville",
    "Russia Time Zone 10": "Asia/Srednekolymsk",
    "Magadan Standard Time": "Asia/Magadan",
    "Norfolk Standard Time": "Pacific/Norfolk",
    "Sakhalin Standard Time": "Asia/Sakhalin",
    "Central Pacific Standard Time": "Pacific/Guadalcanal",
    "Russia Time Zone 11": "Asia/Kamchatka",
    "New Zealand Standard Time": "Pacific/Auckland",
    "UTC+12": "Etc/GMT-12",
    "Fiji Standard Time": "Pacific/Fiji",
    "Chatham Islands Standard Time": "Pacific/Chatham",
    "UTC+13": "Etc/GMT-13",
    "Tonga Standard Time": "Pacific/Tongatapu",
    "Samoa Standard Time": "Pacific/Apia",
    "Line Islands Standard Time": "Pacific/Kiritimati",
  };

  let calendarIdToTitle = {};
  let mutexLinkMenuCalendar_onClick = false;

  async function handleStatus(response, token) {
    if (response.status >= 200 && response.status < 300) {
      return Promise.resolve(response);
    } else {
      if (response.status == 401) {
        // The token has become invalid, revoke it.
        await chrome.identity.removeCachedAuthToken({
          token
        });
        await authenticate(true);
      }
      return Promise.reject(new Error(await response.text()));
    }
  }

  async function injectSnackbar(tab) {
    // Load styles first to prevent flashes
    await chrome.scripting.insertCSS({
      target: { tabId: tab.id },
      files: ["snackbar.css"]
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["snackbar.js"]
    });
  }

  function showSnackbar(tabId, text, action_label, callback) {
    chrome.tabs.sendMessage(tabId, {
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
    // ICAL 1.4.0 offers to make the parser more lenient.
    ICAL.design.strict = false;
    let icalData = ICAL.parse(icsContent);
    let icalRoot = new ICAL.Component(icalData);
    // ical.js does not automatically populate its TimezoneService with
    // custom time zones defined in VTIMEZONE components
    icalRoot.getAllSubcomponents("vtimezone").forEach(
      vtimezone => ICAL.TimezoneService.register(vtimezone));
    return icalRoot.getAllSubcomponents("vevent").map(
      vevent => new ICAL.Event(vevent));
  }

  async function fetchDefaultTimeZone(token, calendarId) {
    let response = await fetch(
      `https://www.googleapis.com/calendar/v3/calendars/${calendarId}`, {
      method: "GET",
      headers: new Headers({
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      }),
    })
      .then(response => handleStatus(response, token));
    let responseJson = await response.json();
    return responseJson.timeZone;
  }

  async function linkMenuCalendar_onClick(info) {
    // Temporarily ignore clicks while we process one
    if (mutexLinkMenuCalendar_onClick)
      return;
    mutexLinkMenuCalendar_onClick = true;
    const [activeTab] = await chrome.tabs.query({
      active: true, lastFocusedWindow: true
    });
    const activeTabId = activeTab.id;
    await injectSnackbar(activeTab);
    // If the user disables and then re-enables the app, the context menus will
    // show up magically without installContextMenus ever having been called.
    // In this case, we have to populate the list of calendars here.
    if (Object.keys(calendarIdToTitle).length === 0)
      await fetchCalendars(true);
    const calendarId = info.menuItemId.split("/")[1];
    const calendarTitle = calendarIdToTitle[calendarId];
    let responseText = '';
    try {
      let response = await fetch(info.linkUrl, {
        credentials: 'include'
      })
        .then(handleStatus);
      responseText = await response.text();
    } catch (error) {
      showSnackbar(activeTabId, "Can't fetch iCal file.");
      console.log(error);
      mutexLinkMenuCalendar_onClick = false;
      return;
    }
    let defaultTimeZone = '';
    try {
      let token = await authenticate(false);
      defaultTimeZone = await fetchDefaultTimeZone(token, calendarId);
    } catch (error) {
      showSnackbar(activeTabId, "Can't fetch calendar's default time zone.");
      console.log(error);
      mutexLinkMenuCalendar_onClick = false;
      return;
    }
    let gcalEventsAndExDates = [];
    try {
      gcalEventsAndExDates = parseIcs(responseText).map(
        event => toGcalEvent(event, activeTab, defaultTimeZone));
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
        `Importing event into '${calendarTitle}'...`;
    } else {
      importMessage =
        `Importing ${gcalEventsAndExDates.length} events into '${calendarTitle}'...`;
    }

    // As we implement cancelling via lazy execution, we will notify the user of
    // possible event loss if they try to leave the page while we haven't
    // committed.
    try {
      await chrome.scripting.executeScript({
        target: {tabId: activeTabId},
        function: () => {
          document.addEventListener('DOMContentLoaded', function() {
            window.onbeforeunload = function() { return true; };
          });
        },
      });
    } catch (error) {
      console.log(error);
    }

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
              clicked => clicked ? chrome.tabs.create({
                active: true,
                openerTabId: activeTabId,
                url: eventResponses[0].htmlLink
              }) : {});
          } else {
            showSnackbar(activeTabId,
              `${eventResponses.length} events imported into '${calendarIdToTitle[calendarId]}'.`,
              "View all",
              clicked => clicked ? eventResponses.forEach(response =>
                chrome.tabs.create({
                  active: true,
                  openerTabId: activeTabId,
                  url: response.htmlLink
                })) : {});
          }
        }
      }();
      mutexLinkMenuCalendar_onClick = false;
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

  async function authenticate(isInteractive) {
    let token = '';
    try {
      // Strip the scopes.
      const response = await chrome.identity.getAuthToken({
        interactive: isInteractive
      });
      token = response.token;
    } catch (error) {
      await updateBrowserAction(false);
      throw error;
    }
    if (isInteractive)
      await updateBrowserAction(true);
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
      .then(response => handleStatus(response, token));
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
      .then(response => handleStatus(response, token));
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
      .then(response => handleStatus(response, token));
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
        startDay.adjust( /* d */ 0, -exDate.hour, -exDate.minute, -exDate.second);
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

  function normalizeTimezone(timeZone, defaultTimeZone) {
    if (timeZone == 'floating') {
      return defaultTimeZone;
    }
    if (WINDOWS_TO_IANA_TIMEZONES.hasOwnProperty(timeZone)) {
      return WINDOWS_TO_IANA_TIMEZONES[timeZone];
    }
    return timeZone;
  }

  function toGcalEvent(event, tabInfo, defaultTimeZone) {
    let gcalEvent = {
      'iCalUID': event.uid || new UUID(4),
      'start': {
        'dateTime': event.startDate.toString(),
        'timeZone': normalizeTimezone(event.startDate.zone.toString(), defaultTimeZone)
      },
      'end': {
        'dateTime': event.endDate.toString(),
        'timeZone': normalizeTimezone(event.endDate.zone.toString(), defaultTimeZone)
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
    }).then(response => handleStatus(response, token));
    return response.json();
  }

  async function updateBrowserAction(active) {
    if (active) {
      await chrome.action.setIcon({
        path: "images/logo_active.png"
      });
      await chrome.action.setTitle({
        title: "Update calendar list"
      });
    } else {
      await chrome.action.setIcon({
        path: "images/logo_inactive.png"
      });
      await chrome.action.setTitle({
        title: "Authorize ICS to GCal to access (read/write) your Google Calendar"
      });
    }
  }

  async function fetchCalendars(interactive) {
    chrome.contextMenus.removeAll();
    let token = await authenticate(interactive);
    if (token === '') {
      await updateBrowserAction(false);
    }
    let responseCalendarList = null;
    try {
      let response = await fetch(
        `https://www.googleapis.com/calendar/v3/users/me/calendarList`, {
        method: "GET",
        headers: new Headers({
          "Authorization": `Bearer ${token}`,
        }),
      })
        .then(response => handleStatus(response, token));
      responseCalendarList = await response.json();
    } catch (error) {
      console.log("Failed to fetch calendars.");
      console.log(error);
      await updateBrowserAction(false);
      return;
    }
    await updateBrowserAction(true);
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
  chrome.action.onClicked.addListener(fetchCalendars.bind(null, true));
  chrome.contextMenus.onClicked.addListener(linkMenuCalendar_onClick);
})();
