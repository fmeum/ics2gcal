# ICS to GCal
Found an interesting event online that comes as an iCal file (`*.ics`) and want to import it into one of your Google calendars? With ICS to GCal this will take just two clicks - no need to open the Google Calendar settings and pollute your list of calendars.

ICS to GCal correctly imports recurrent events (to the extent they are supported by Google Calendar) and automatically fixes some deviations from the iCalendar specification that are rather common (such as trailing commas in date lists, exceptional dates not specified with the correct time,...).

![ics2gcal](https://github.com/FabianHenneke/ics2gcal/raw/master/assets/ics2gcal.1.png)

How to use:

0. Before the first use, click on the extension symbol in the browser toolbar to allow access to you calendars.
1. Right-click on a link to an iCal file (`*.ics`).
2. Select a calendar under "Add to calendar".
A popup notification will inform you about the number of events added and allow you to view the event or undo the changes while they have not been committed.

# Install in Chrome
You can install the extension directly from the [Chrome Web Store](https://chrome.google.com/webstore/detail/ics-to-gcal/ljobcbehhifehkmamikmchekbbljopao). If you want to make changes to the extension and test it locally, you have to:

1. Run `yarn install` to install dependencies ([ical.js](https://github.com/mozilla-comm/ical.js/) and [pure-uuid](https://github.com/rse/pure-uuid)).
2. Run `yarn run build` to resolve development dependencies (due to changes in Chrome Manifest 3, libraries using CommonJS were the easiest way to restore functionality).
3. In Chrome, under `More tools -> Extensions`, mark the `Developer mode` checkbox and use `Load unpacked extension...`.
4. Select the build directory.

Examples of both valid and invalid iCalendar files to test on can be found in the `tests` folder. They have to be hosted by a web server since the Chrome extension has no permission to access `file://` URLs.
