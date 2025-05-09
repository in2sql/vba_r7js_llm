# Application.Worksheets property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
MsgBox Worksheets("Sheet1").Range("A1").Value
```

## Remarks
Using this property without an object qualifier returns all the worksheets in the active workbook.

## Example
```vba
Set newSheet = Worksheets.Add 
newSheet.Name = "current Budget"
```

