# Range Parse Method

## Business Description
Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.

## Behavior
Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.

## Example Usage
```vba
Worksheets("Sheet1").Columns("A").Parse_ 
 parseLine:="[xxx] [xxxxxxxx]", _ 
 destination:=Worksheets("Sheet1").Range("B1")
```