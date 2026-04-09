# EDL to CSV/XLSX Converter

A lightweight macOS drag-and-drop app that converts CMX3600 EDL (Edit Decision List) files to CSV or XLSX format. Built with JavaScript for Automation (JXA), requiring zero dependencies.

## Features

- **Drag and drop** any `.edl` file onto the app to convert instantly
- **CSV or XLSX output** with a saved preference (no prompt on each drop)
- **All original EDL fields preserved**: Event, Reel, Track, Transition, Source In/Out, Record In/Out, Clip Name
- **Timecodes kept as-is** in HH:MM:SS:FF format
- **XLSX extras**: bold headers, pre-sized columns (Clip Name column is extra wide for long filenames)
- **Auto-detects frame rate** from FCM line (24fps non-drop / 29.97fps drop frame)
- **No dependencies**: pure JXA, runs on any Mac without Python, Node, or Homebrew
- **Properly signed**: compiled with `osacompile`, passes Gatekeeper on copy

## Usage

1. **Double-click** the app to open preferences and set your output format (CSV or XLSX)
2. **Drag and drop** `.edl` files onto the app icon
3. The converted file appears next to the original EDL

Your format preference is saved automatically and persists across reboots.

## EDL Format Support

Parses standard CMX3600 EDL files as exported by DaVinci Resolve, Avid Media Composer, Adobe Premiere, and other NLEs.

Example EDL input:
```
TITLE: My Timeline
FCM: NON-DROP FRAME

001  AX       V     C        00:00:44:18 00:00:46:06 00:00:00:00 00:00:01:12
* FROM CLIP NAME: Interview_A.mov
```

Example CSV output:
```
Event,Reel,Track,Transition,Source In,Source Out,Record In,Record Out,Clip Name
001,AX,V,C,00:00:44:18,00:00:46:06,00:00:00:00,00:00:01:12,Interview_A.mov
```

## Building from Source

To recompile the app from the JXA source:

```bash
osacompile -l JavaScript -o "EDL to CSV.app" edl_to_csv.js
```

## Author

Chad Littlepage
chad.littlepage@gmail.com | 323.974.0444
