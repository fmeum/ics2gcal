#!/usr/bin/env bash
yarn install

mkdir -p build
cp background.js manifest.json snackbar.css snackbar.js build
mkdir -p build/node_modules/ical.js/build
cp node_modules/ical.js/build/ical.min.js build/node_modules/ical.js/build
mkdir -p build/node_modules/chrome-promise
cp node_modules/chrome-promise/chrome-promise.js build/node_modules/chrome-promise
mkdir -p build/images
cp images/*.png build/images/
cd build
zip -r ../ics2gcal.zip *
