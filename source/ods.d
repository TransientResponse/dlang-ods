module ods;

import std.stdio, std.utf, std.file;
import archive.zip;
import dxml.parser;

private enum configSplitYes = makeConfig(SplitEmpty.yes);
alias rangeT = EntityRange!(configSplitYes, string);

private string[string] getAttributeDict(rangeT range) {
	if(range.front.type != EntityType.elementStart) throw new Exception("No attributes on non-start tag");
	auto attrs = range.front.attributes;
	string[string] temp;
	while(!attrs.empty) {
		temp[attrs.front.name] = attrs.front.value;
		attrs.popFront;
	}

	return temp;
}

/**
Parses an ODS file and presents the results as a lazy forward range of rows. 

Usage: 
~~~
auto sheet = new ODSSheet();
sheet.loadFromFile("file.ods", 0); //The first sheet
while(!sheet.empty) { writeln(sheet.front); sheet.popFront; }
~~~
*/
public class ODSSheet {
	private rangeT range;
	private string[] _row;
	private bool _endOfSheet;

	/** Retruns `true` if there are no more rows to read, and false otherwise. */
	public bool empty() {
		return range.empty && !_endOfSheet;
	}

	/** Returns the current row. */
	public string[] front() {
		return _row;
	}

	/** Parses the next avaialable row, if available. */
	public void popFront() {
		if(empty) return;
		_row = parseNextRow;
	}

	/** Returns a copy of this object that can be iterated separately. */
	public ODSSheet save() {
		auto temp = new ODSSheet;
		temp._row = _row;
		temp.range = range.save();
		return temp;
	}

	/** 
	Reads a sheet by index from a given file.

	Params:
	filename = The name of the ODS file to read.
	sheet = The one-based index of the sheet to read.
	*/
	public void readSheet(string filename, int sheet) {
		loadFile(filename);
		runToSheet(sheet);
		_row = parseNextRow();
	}

	/**
	Reads a sheet by name from a given file.

	Params:
	filename = The name of the ODS file to read.
	sheetName = The name of the sheet to read.
	*/
	public void readSheetByName(string filename, string sheetName) {
		loadFile(filename);
		runToSheet(sheetName);
		_row = parseNextRow();
	}

	private void loadFile(string filename) {
		auto zip = new ZipArchive(read(filename));
		auto content = zip.getFile("content.xml");
		if(content is null) throw new Exception("Invalid ODS file (no content.xml)");
		auto data = content.data;
		string xml = cast(string)data;
		validate(xml);
		range = parseXML!configSplitYes(xml);
	}

	private string[] parseNextRow() {
		if(!range.empty) range.popFront;
		string[] row;
		while(!range.empty) {
			if(range.front.type == EntityType.elementStart) {
				//if(range.front.name == "table:table-row") row = new string[];
				if(range.front.name == "text:p") {
					range.popFront;
					row ~= range.front.text;
				}
			}
			else if(range.front.type == EntityType.elementEnd) {
				if(range.front.name == "table:table-row") return row;
				else if(range.front.name == "table:table") _endOfSheet = true;
			}
			range.popFront;
		}
		return null;
	}
	unittest {
		string rowXML = `<table:table-row table:style-name="ro1">
		<table:table-cell office:value-type="string" calcext:value-type="string">
			<text:p>This</text:p>
		</table:table-cell>
		<table:table-cell office:value-type="string" calcext:value-type="string">
			<text:p>Is</text:p>
		</table:table-cell>
		<table:table-cell office:value-type="string" calcext:value-type="string">
			<text:p>A</text:p>
		</table:table-cell>
		<table:table-cell office:value-type="string" calcext:value-type="string">
			<text:p>Test</text:p>
		</table:table-cell>
	</table:table-row>`;

		auto range = parseXML!configSplitYes(rowXML);
		assert(parseNextRow(range) == ["This", "Is", "A", "Test"]);
	}


	private void runToSheet(int sheet) {
		int N = 0;
		while(!range.empty) {
			if((range.front.type == EntityType.elementStart) && (range.front.name == "table:table")) {
				if(++N > sheet) return;
			}
			range.popFront;
		}
	}

	private void runToSheet(string sheetName) {
		while(!range.empty) {
			if((range.front.type == EntityType.elementStart) && (range.front.name == "table:table")) {
				string[string] attrs = getAttributeDict(range);
				if(auto name = "name" in attrs) {
					if(*name == sheetName) return;
				}
			}
			range.popFront;
		}
		throw new Exception("No sheet by that name");
	}
}