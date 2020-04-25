This program allows you to load and display JSON string of any structure on an Excel worksheet.
Edit it and save back to json file, with the same or another structure

This tool is not quite for the programmer - rather for an IT specialist
It is necessary for processing of considerable volumes of json data (not huge, but not small).

The basic principle: 
Each column of header row contain the path to the json data, the filled cells of the column contain data having the same path which specified in the header of column. Detailed description see below in "Description of format output".

The project provided a ready-to-use Excel file, containing on the "Sheet1" buttons and all the necessary descriptions and of course macros.

Json parsing implemented using library https://github.com/VBA-tools/VBA-JSON

Converting json data to a format for placing on an Excel sheet and vice versa is implemented by its own library JsonSheet.bas

## Description of format output
Elementary type values are data that is not a structure or an array.
For each value of an elementary type in a json string, can be formed a string describing the "path" to it value, recorded by the following rules:
* Recording the path starts from the beginning of the json string, from left to right.
* "$" - indicates the root, located at the beginning of each path string.
*  "[" - descent to a new level, into the array.
Values of elementary type can be embedded in the json array, in this case after the symbol "[" should follow the symbol "e".
* "{" - descent to a new level, into the structure
After the symbol "{" should follow the field name in quotation marks (").

The header of each Excel column contains a "path" to the value of the elementary data type in the json string.
For each array in json string, additional Excel column is allocated. This column will contain the row numbers of the json array.
Header name (path) for this column, built as described above but the character "i" is added to the end of the string (as a result "$...[i").

Json data is displayed below the Excel header row.
Each elementary data type value from a json string is displayed in a column which header and data path match.
For each row of the json array, a separate row of the Excel sheet is allocated, in the corresponding "$...[i" column, displays the row number of the json array (numbering starts from 1).

Sometimes it's necessary, to transpose the output of the array, corresponding columns are allocated for this.
The header of these columns contain the path to the value, with the json array row number added to the end of path (no separators, just a number).

## Public Methods of JsonSheet.bas
* #### Public Function GetJsonHeader(json As Object) As collection
  Returns a collection containing fields path/addresses
  json object can be obtained by: Set json = JsonConverter.ParseJson(jsonString)

* #### Public Sub OutJsonHeader(json As Object, Sheet As Worksheet, Row As Integer, Col As Integer)
  Forms on a sheet columns headings containing field path/addresses

* #### Public Sub OutJsonBody(json As Object, Sheet As Worksheet, Row As Integer, Col As Integer)
  Output on a worksheet content of json. looking at the column headings
  the location of the row of column headers is specified in the call parameters
  data is output below this heading row

* #### Public Function ReadJsonFromSheet(Sheet As Worksheet, Row As Integer, Col As Integer) As Object
  Reads data from a worksheet and fills an object representing Json (Dictionary or Collection in root)
  further, this object can be converted to json string using JsonConverter.ConvertToJson(<object>, Whitespace:=2)
  call parameters must point to the first cell of the heading row
