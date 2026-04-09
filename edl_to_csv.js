/*
 * EDL to CSV/XLSX Converter
 * Drag and drop .edl files onto this app to convert them.
 * Outputs a .csv or .xlsx next to the original file.
 *
 * Supports:
 *   - CMX3600 EDL (standard timeline EDL with clip names)
 *   - DaVinci Resolve Markers EDL (with color, marker name, duration)
 *
 * Auto-detects EDL type and includes all fields.
 *
 * Author: Chad Littlepage
 * Contact: chad.littlepage@gmail.com | 323.974.0444
 * Version: 1.3
 */

ObjC.import('Foundation');

var BUNDLE_ID = 'com.chadlittlepage.edl-to-csv';

function getPreferredFormat() {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	try {
		var fmt = app.doShellScript('defaults read ' + BUNDLE_ID + ' outputFormat');
		if (fmt === 'xlsx' || fmt === 'csv') return fmt;
	} catch (e) {}
	return 'csv';
}

function setPreferredFormat(fmt) {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	app.doShellScript('defaults write ' + BUNDLE_ID + ' outputFormat -string ' + fmt);
}

/* ── Detect EDL type ── */

function detectEDLType(text) {
	// Markers EDL has |C: |M: |D: lines
	if (text.indexOf('|C:') !== -1 && text.indexOf('|M:') !== -1 && text.indexOf('|D:') !== -1) {
		return 'markers';
	}
	return 'standard';
}

/* ── Standard CMX3600 parser ── */

var STD_HEADERS = ['Event', 'Reel', 'Track', 'Transition', 'Source In', 'Source Out', 'Record In', 'Record Out', 'Clip Name'];
var STD_COL_WIDTHS = [8, 8, 7, 11, 14, 14, 14, 14, 45];

function parseStandard(text) {
	var lines = text.split(/\r?\n/);
	var events = [];

	var eventRe = /^\s*(\d{3,})\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S*)\s*(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})/;

	for (var i = 0; i < lines.length; i++) {
		var m = eventRe.exec(lines[i]);
		if (m) {
			var clipName = '';
			for (var j = i + 1; j < lines.length; j++) {
				if (lines[j].charAt(0) !== '*' && lines[j].charAt(0) !== '|') break;
				var comment = lines[j].toUpperCase();
				var idx = comment.indexOf('FROM CLIP NAME:');
				if (idx === -1) idx = comment.indexOf('CLIP NAME:');
				if (idx !== -1) {
					clipName = lines[j].substring(lines[j].indexOf(':', idx) + 1).trim();
					break;
				}
			}

			events.push([m[1], m[2], m[3], m[4], m[6], m[7], m[8], m[9], clipName || ('Event ' + m[1] + ' (' + m[2] + ')')]);
		}
	}

	return {headers: STD_HEADERS, colWidths: STD_COL_WIDTHS, rows: events};
}

/* ── Markers EDL parser ── */

var MKR_HEADERS = ['Event', 'Reel', 'Track', 'Transition', 'Source In', 'Source Out', 'Record In', 'Record Out', 'Color', 'Marker Name', 'Duration (frames)'];
var MKR_COL_WIDTHS = [8, 8, 7, 11, 14, 14, 14, 14, 22, 45, 18];

function parseMarkers(text) {
	var lines = text.split(/\r?\n/);
	var events = [];

	var eventRe = /^\s*(\d{3,})\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S*)\s*(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})/;
	var colorRe = /\|C:([^\|]*)/;
	var markerRe = /\|M:([^\|]*)/;
	var durRe = /\|D:([^\|]*)/;

	for (var i = 0; i < lines.length; i++) {
		var m = eventRe.exec(lines[i]);
		if (m) {
			var color = '';
			var markerName = '';
			var duration = '';

			// Look at following comment lines for marker data
			for (var j = i + 1; j < lines.length; j++) {
				var cl = lines[j];
				if (cl.indexOf('|') === -1 && cl.charAt(0) !== '*' && cl.trim() !== '') break;
				if (cl.indexOf('|C:') !== -1 || cl.indexOf('|M:') !== -1 || cl.indexOf('|D:') !== -1) {
					var cm = colorRe.exec(cl);
					if (cm) color = cm[1].trim();
					var mm = markerRe.exec(cl);
					if (mm) markerName = mm[1].trim();
					var dm = durRe.exec(cl);
					if (dm) duration = dm[1].trim();
					break;
				}
			}

			events.push([m[1], m[2], m[3], m[4], m[6], m[7], m[8], m[9], color, markerName, duration]);
		}
	}

	return {headers: MKR_HEADERS, colWidths: MKR_COL_WIDTHS, rows: events};
}

/* ── CSV writer ── */

function escapeCSV(val) {
	if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
		return '"' + val.replace(/"/g, '""') + '"';
	}
	return val;
}

function writeCSV(parsed, outPath) {
	var csvLines = [parsed.headers.join(',')];
	for (var i = 0; i < parsed.rows.length; i++) {
		csvLines.push(parsed.rows[i].map(escapeCSV).join(','));
	}
	var csvText = csvLines.join('\n') + '\n';
	$(csvText).writeToFileAtomicallyEncodingError($(outPath), true, $.NSUTF8StringEncoding, null);
}

/* ── XLSX writer ── */

function escapeXML(val) {
	return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colLetter(idx) {
	if (idx < 26) return String.fromCharCode(65 + idx);
	return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
}

function writeXLSX(parsed, outPath) {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	var tmpDir = app.doShellScript('mktemp -d');

	app.doShellScript('mkdir -p "' + tmpDir + '/_rels" "' + tmpDir + '/xl/_rels" "' + tmpDir + '/xl/worksheets"');

	var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/>' +
		'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
		'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
		'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
		'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
		'</Types>';

	var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
		'</Relationships>';

	var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
		'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
		'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
		'</Relationships>';

	var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
		'<sheets><sheet name="EDL" sheetId="1" r:id="rId1"/></sheets>' +
		'</workbook>';

	var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
		'<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
		'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
		'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
		'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
		'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
		'</styleSheet>';

	// Shared strings
	var ssMap = {};
	var ssList = [];
	function ssIndex(val) {
		if (ssMap[val] === undefined) {
			ssMap[val] = ssList.length;
			ssList.push(val);
		}
		return ssMap[val];
	}

	var numCols = parsed.headers.length;
	for (var h = 0; h < numCols; h++) ssIndex(parsed.headers[h]);
	for (var i = 0; i < parsed.rows.length; i++) {
		for (var c = 0; c < parsed.rows[i].length; c++) ssIndex(parsed.rows[i][c]);
	}

	var ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (numCols + parsed.rows.length * numCols) + '" uniqueCount="' + ssList.length + '">';
	for (var s = 0; s < ssList.length; s++) {
		ssXml += '<si><t>' + escapeXML(ssList[s]) + '</t></si>';
	}
	ssXml += '</sst>';

	// Sheet
	var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

	sheet += '<cols>';
	for (var c = 0; c < numCols; c++) {
		sheet += '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + parsed.colWidths[c] + '" customWidth="1"/>';
	}
	sheet += '</cols>';

	sheet += '<sheetData>';

	// Header row
	sheet += '<row r="1">';
	for (var h = 0; h < numCols; h++) {
		sheet += '<c r="' + colLetter(h) + '1" t="s" s="1"><v>' + ssIndex(parsed.headers[h]) + '</v></c>';
	}
	sheet += '</row>';

	// Data rows
	for (var i = 0; i < parsed.rows.length; i++) {
		var rowNum = i + 2;
		sheet += '<row r="' + rowNum + '">';
		for (var c = 0; c < parsed.rows[i].length; c++) {
			sheet += '<c r="' + colLetter(c) + rowNum + '" t="s"><v>' + ssIndex(parsed.rows[i][c]) + '</v></c>';
		}
		sheet += '</row>';
	}

	sheet += '</sheetData></worksheet>';

	function writeFile(path, content) {
		$(content).writeToFileAtomicallyEncodingError($(path), true, $.NSUTF8StringEncoding, null);
	}

	writeFile(tmpDir + '/[Content_Types].xml', contentTypes);
	writeFile(tmpDir + '/_rels/.rels', rels);
	writeFile(tmpDir + '/xl/_rels/workbook.xml.rels', wbRels);
	writeFile(tmpDir + '/xl/workbook.xml', workbook);
	writeFile(tmpDir + '/xl/styles.xml', styles);
	writeFile(tmpDir + '/xl/sharedStrings.xml', ssXml);
	writeFile(tmpDir + '/xl/worksheets/sheet1.xml', sheet);

	app.doShellScript('cd "' + tmpDir + '" && zip -r "' + outPath + '" . -x ".*"');
	app.doShellScript('rm -rf "' + tmpDir + '"');
}

/* ── Main convert ── */

function convertFile(posixPath, format) {
	var data = $.NSString.stringWithContentsOfFileEncodingError($(posixPath), $.NSUTF8StringEncoding, null);
	if (!data) return {error: 'Could not read file.'};

	var text = ObjC.unwrap(data);
	var edlType = detectEDLType(text);
	var parsed;

	if (edlType === 'markers') {
		parsed = parseMarkers(text);
	} else {
		parsed = parseStandard(text);
	}

	if (parsed.rows.length === 0) {
		return {error: 'No events found in EDL.'};
	}

	var basePath = posixPath.replace(/\.[^.]+$/, '');
	var outPath;

	if (format === 'xlsx') {
		outPath = basePath + '.xlsx';
		writeXLSX(parsed, outPath);
	} else {
		outPath = basePath + '.csv';
		writeCSV(parsed, outPath);
	}

	var typeLabel = edlType === 'markers' ? 'Markers' : 'Standard';
	return {path: outPath, count: parsed.rows.length, type: typeLabel};
}

function showNotification(title, message) {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	app.displayNotification(message, {withTitle: title});
}

// Drag-and-drop handler
function openDocuments(docs) {
	var format = getPreferredFormat();
	var results = [];
	for (var i = 0; i < docs.length; i++) {
		var posixPath = docs[i].toString();
		if (posixPath.indexOf('/') === -1) {
			posixPath = ObjC.unwrap($(posixPath).stringByResolvingSymlinksInPath);
		}
		try {
			var baseName = posixPath.split('/').pop();
			var r = convertFile(posixPath, format);
			if (r.error) {
				results.push(baseName + ': ' + r.error);
			} else {
				results.push(baseName + ': ' + r.count + ' events (' + r.type + ' EDL) exported to ' + format.toUpperCase() + '.');
			}
		} catch (e) {
			results.push(posixPath.split('/').pop() + ': Error - ' + e.message);
		}
	}
	var summary = results.join('\n');
	showNotification('EDL Converter', summary);
}

// Double-click handler
function run() {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	var current = getPreferredFormat().toUpperCase();

	var result = app.displayDialog(
		'Current output format: ' + current + '\n\n' +
		'Choose a new default, or click OK to keep it.\n\n' +
		'Supports:\n' +
		'  \u2022 Standard CMX3600 EDL (timeline cuts)\n' +
		'  \u2022 DaVinci Resolve Markers EDL (color, name, duration)\n\n' +
		'EDL type is auto-detected. Drag and drop .edl files to convert.', {
		withTitle: 'EDL Converter \u2014 Preferences',
		buttons: ['CSV', 'XLSX', 'OK'],
		defaultButton: 'OK'
	});

	var btn = result.buttonReturned;
	if (btn === 'CSV') {
		setPreferredFormat('csv');
		app.displayNotification('Output format set to CSV.', {withTitle: 'EDL Converter'});
	} else if (btn === 'XLSX') {
		setPreferredFormat('xlsx');
		app.displayNotification('Output format set to XLSX.', {withTitle: 'EDL Converter'});
	}
}
