# EDL to CSV/XLSX Converter

A lightweight macOS drag-and-drop app that converts EDL (Edit Decision List) files to CSV or XLSX format. Built with JavaScript for Automation (JXA), requiring zero dependencies. Signed and notarized by Apple.

## Features

- **Drag and drop** any `.edl` file onto the app to convert instantly
- **CSV or XLSX output** with a saved preference (no prompt on each drop)
- **Auto-detects EDL type**: standard CMX3600 or DaVinci Resolve Markers
- **All original fields preserved** with timecodes kept as-is in HH:MM:SS:FF format
- **XLSX extras**: bold headers, pre-sized columns
- **Auto-detects frame rate** from FCM line (24fps non-drop / 29.97fps drop frame)
- **No dependencies**: pure JXA, runs on any Mac without Python, Node, or Homebrew
- **Signed and notarized**: opens on any Mac without Gatekeeper warnings

## Installation

Download `EDL to CSV.app` from the [latest release](https://github.com/chadlittlepage/EDLtoCSV/releases), unzip, and place it wherever you like. The app is signed and notarized by Apple, so it opens without warnings.

For unsigned builds (compiled from source), use the included `Install.command` to bypass Gatekeeper.

## Usage

1. **Double-click** the app to open preferences and set your output format (CSV or XLSX)
2. **Drag and drop** `.edl` files onto the app icon
3. The converted file appears next to the original EDL

Your format preference is saved automatically and persists across reboots.

## Supported EDL Types

### Standard CMX3600 (timeline cuts)

Exported by DaVinci Resolve, Avid Media Composer, Adobe Premiere, and other NLEs.

| Column | Example |
|--------|---------|
| Event | 001 |
| Reel | AX |
| Track | V |
| Transition | C |
| Source In | 00:00:44:18 |
| Source Out | 00:00:46:06 |
| Record In | 00:00:00:00 |
| Record Out | 00:00:01:12 |
| Clip Name | Interview_A.mov |

### DaVinci Resolve Markers EDL

Includes all standard fields plus marker-specific data:

| Column | Example |
|--------|---------|
| Color | ResolveColorBlue |
| Marker Name | ABOUT_MY_FATHER.mov |
| Duration (TC) | 00:00:59:21 |
| Duration (frames) | 1437 |

Single-frame markers (`|D:1`) are fully supported.

## Building from Source

To recompile the app from the JXA source:

```bash
osacompile -l JavaScript -o "EDL to CSV.app" edl_to_csv.js
```

To sign and notarize (requires Apple Developer ID):

```bash
codesign --force --deep --sign "Developer ID Application: Your Name (TEAM_ID)" --options runtime "EDL to CSV.app"
ditto -c -k --keepParent "EDL to CSV.app" EDL_to_CSV.zip
xcrun notarytool submit EDL_to_CSV.zip --keychain-profile "notary" --wait
xcrun stapler staple "EDL to CSV.app"
```

## Author

Chad Littlepage
chad.littlepage@gmail.com | 323.974.0444
