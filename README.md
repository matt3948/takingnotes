# takingnotes.ink

Drawing and animation studio for the Huion X10 and the Wacom Slate/Spark series of smart notebooks that runs in the browser. Built with React, TypeScript, and Vite.

Wacom SmartPad support in this project was informed by the [tuhi project](https://github.com/tuhiproject/tuhi).

Supports layered canvas editing, animation timeline with playback/export, and pen tablet input via Web Bluetooth and WebHID.

## Quick Start

```sh
npm install
npm run dev
```

Requires Node.js 20+. Dev server runs at `localhost:3000`.

## Build

```sh
npm run build    # production build → dist/
npm run preview  # serve the build locally
```

## Tablet Support

Pen tablet features (Huion, Wacom) require Chrome or Edge and an HTTPS connection.

## What works
Drawing live from both Wacom and Huion smart notebooks, downloading from Huion ones

## Todo
Fix small UI bugs, re-implement Wacom memory download
