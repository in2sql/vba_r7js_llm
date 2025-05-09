# Submit time and task status in Project Server 2013

## Business Description
By using the appropriate method, you can easily refer to multiple ranges. Use the Range and Union methods to refer to any group of ranges. Use the Areas property to refer to the group of ranges selected on a worksheet.

## Behavior
By using the appropriate method, you can easily refer to multiple ranges. Use theRangeandUnionmethods to refer to any group of ranges. Use theAreasproperty to refer to the group of ranges selected on a worksheet.

## Example Usage
```vba
Sub ClearRanges() 
 Worksheets("Sheet1").Range("C5:D9,G9:H16,B14:D18"). _ 
 ClearContents 
End Sub
```