/*
 * EDL to CSV/XLSX Converter
 * Drag and drop .edl files onto this app to convert them.
 * Outputs a .csv or .xlsx next to the original file.
 *
 * Supports CMX3600 EDL format with all original fields preserved.
 *
 * Author: Chad Littlepage
 * Contact: chad.littlepage@gmail.com | 323.974.0444
 * Version: 1.2
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

function parseCMX3600(text) {
	var lines = text.split(/\r?\n/);
	var events = [];
	var title = '';
	var fcm = '';

	var eventRe = /^\s*(\d{3,})\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S*)\s*(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})/;

	for (var i = 0; i < lines.length; i++) {
		var line = lines[i];

		if (line.indexOf('TITLE:') === 0) {
			title = line.substring(6).trim();
			continue;
		}
		if (line.toUpperCase().indexOf('FCM:') !== -1) {
			fcm = line.substring(line.indexOf(':') + 1).trim();
			continue;
		}

		var m = eventRe.exec(line);
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

			events.push({
				event: m[1],
				reel: m[2],
				track: m[3],
				transition: m[4],
				srcIn: m[6],
				srcOut: m[7],
				recIn: m[8],
				recOut: m[9],
				clipName: clipName
			});
		}
	}

	return {events: events, title: title, fcm: fcm};
}

function escapeCSV(val) {
	if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
		return '"' + val.replace(/"/g, '""') + '"';
	}
	return val;
}

function escapeXML(val) {
	return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

var HEADERS = ['Event', 'Reel', 'Track', 'Transition', 'Source In', 'Source Out', 'Record In', 'Record Out', 'Clip Name'];
var COL_WIDTHS = [8, 8, 7, 11, 14, 14, 14, 14, 45];

function eventToRow(e) {
	return [e.event, e.reel, e.track, e.transition, e.srcIn, e.srcOut, e.recIn, e.recOut, e.clipName];
}

function writeCSV(events, outPath) {
	var csvLines = [HEADERS.join(',')];
	for (var i = 0; i < events.length; i++) {
		var row = eventToRow(events[i]);
		csvLines.push(row.map(escapeCSV).join(','));
	}
	var csvText = csvLines.join('\n') + '\n';
	$(csvText).writeToFileAtomicallyEncodingError($(outPath), true, $.NSUTF8StringEncoding, null);
}

function writeXLSX(events, outPath) {
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

	var ssMap = {};
	var ssList = [];
	function ssIndex(val) {
		if (ssMap[val] === undefined) {
			ssMap[val] = ssList.length;
			ssList.push(val);
		}
		return ssMap[val];
	}

	for (var h = 0; h < HEADERS.length; h++) ssIndex(HEADERS[h]);
	for (var i = 0; i < events.length; i++) {
		var row = eventToRow(events[i]);
		for (var c = 0; c < row.length; c++) ssIndex(row[c]);
	}

	var ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (HEADERS.length + events.length * 9) + '" uniqueCount="' + ssList.length + '">';
	for (var s = 0; s < ssList.length; s++) {
		ssXml += '<si><t>' + escapeXML(ssList[s]) + '</t></si>';
	}
	ssXml += '</sst>';

	var colLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

	var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

	sheet += '<cols>';
	for (var c = 0; c < COL_WIDTHS.length; c++) {
		sheet += '<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + COL_WIDTHS[c] + '" customWidth="1"/>';
	}
	sheet += '</cols>';

	sheet += '<sheetData>';

	sheet += '<row r="1">';
	for (var h = 0; h < HEADERS.length; h++) {
		sheet += '<c r="' + colLetters[h] + '1" t="s" s="1"><v>' + ssIndex(HEADERS[h]) + '</v></c>';
	}
	sheet += '</row>';

	for (var i = 0; i < events.length; i++) {
		var rowNum = i + 2;
		var row = eventToRow(events[i]);
		sheet += '<row r="' + rowNum + '">';
		for (var c = 0; c < row.length; c++) {
			sheet += '<c r="' + colLetters[c] + rowNum + '" t="s"><v>' + ssIndex(row[c]) + '</v></c>';
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

function convertFile(posixPath, format) {
	var data = $.NSString.stringWithContentsOfFileEncodingError($(posixPath), $.NSUTF8StringEncoding, null);
	if (!data) return {error: 'Could not read file.'};

	var text = ObjC.unwrap(data);
	var result = parseCMX3600(text);

	if (result.events.length === 0) {
		return {error: 'No events found in EDL.'};
	}

	var basePath = posixPath.replace(/\.[^.]+$/, '');

	if (format === 'xlsx') {
		var outPath = basePath + '.xlsx';
		writeXLSX(result.events, outPath);
	} else {
		var outPath = basePath + '.csv';
		writeCSV(result.events, outPath);
	}

	return {path: outPath, count: result.events.length};
}

function showNotification(title, message) {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	app.displayNotification(message, {withTitle: title});
}

// Drag-and-drop: uses saved preference, no prompt
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
				results.push(baseName + ': ' + r.count + ' events exported to ' + format.toUpperCase() + '.');
			}
		} catch (e) {
			results.push(posixPath.split('/').pop() + ': Error - ' + e.message);
		}
	}
	var summary = results.join('\n');
	showNotification('EDL Converter', summary);
}

// Double-click: show preferences
function run() {
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;
	var current = getPreferredFormat().toUpperCase();

	var result = app.displayDialog(
		'Current output format: ' + current + '\n\n' +
		'Choose a new default, or click OK to keep it.\n\n' +
		'Drag and drop .edl files onto this app to convert them.\n' +
		'The file will be saved next to the original .edl file.', {
		withTitle: 'EDL Converter — Preferences',
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

