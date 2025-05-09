# Application.Sheets property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Set newSheet = Sheets.Add(Type:=xlWorksheet) 
For i = 1 To Sheets.Count 
 newSheet.Cells(i, 1).Value = Sheets(i).Name 
Next i
```

## Remarks
Using this property without an object qualifier is equivalent to using ActiveWorkbook.Sheets.

## Example
```vba
Set newSheet = Sheets.Add(Type:=xlWorksheet) 
For i = 1 To Sheets.Count 
 newSheet.Cells(i, 1).Value = Sheets(i).Name 
Next i
```

