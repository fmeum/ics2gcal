#!/usr/bin/env bash
bower install

mkdir -p build
cp background.js manifest.json snackbar.css snackbar.js build
mkdir -p build/bower_components/ical.js/build
cp bower_components/ical.js/build/ical.min.js build/bower_components/ical.js/build
mkdir -p build/bower_components/chrome-promise
cp bower_components/chrome-promise/chrome-promise.js build/bower_components/chrome-promise
mkdir -p build/images
cp images/*.png build/images/
cd build
zip -r ../ics2gcal.zip *
