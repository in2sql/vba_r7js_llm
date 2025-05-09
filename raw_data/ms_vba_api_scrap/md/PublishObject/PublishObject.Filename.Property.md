# PublishObject Filename Property

## Business Description
Returns or sets the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved. Read/write String.

## Behavior
Returns or sets the URL (on the intranet or the Web) or path (local or network) to the location where the specified source object was saved. Read/writeString.

## Example Usage
```vba
ActiveWorkbook.PublishObjects(1).FileName= _ 
 "\\Server2\Q1\StockReport.htm"
```