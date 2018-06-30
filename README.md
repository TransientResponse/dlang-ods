# D language ODS parser

Parses an ODS (OpenDocument Spreadsheet) file, and returns each row in a particular sheet as a lazy range of dynamic string arrays.

Usage:
```{.d}
import std.stdio;
import ods;

void main()
{
	auto sheet = new ODSSheet();

	sheet.readSheet("test.ods", 1);

	while(!sheet.empty) {
		writeln(sheet.front);
		sheet.popFront;
	}
}
```

Sheet by name:
```{.d}
import std.stdio;
import ods;

void main()
{
	auto sheet = new ODSSheet();

	sheet.readSheetByName("test.ods", "Sheet1");

	while(!sheet.empty) {
		writeln(sheet.front);
		sheet.popFront;
	}
}
```

This project is currently in a very early state. I've done very little testing and have only one unit test in the source. 
It should NOT be considered ready for production use. 