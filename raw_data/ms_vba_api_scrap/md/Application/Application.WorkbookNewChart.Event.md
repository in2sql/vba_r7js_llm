# Application.WorkbookNewChart event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_NewChart(ByVal Wb As Workbook, _ 
 ByVal Ch As Chart) 
 MsgBox ("A new chart was added to: " & Wb.Name & " of type: " & Ch.Type) 
End Sub
```

## Parameters
- **Wb**: Required
- **Ch**: Required

## Return Value
Nothing

## Remarks
The WorkbookNewChart event occurs when a new chart is inserted or pasted on a worksheet, a chart sheet, or other sheet types. If multiple charts are inserted or pasted, the event will occur for each chart in the order they are inserted or pasted.

## Example
```vba
Private Sub App_NewChart(ByVal Wb As Workbook, _ 
 ByVal Ch As Chart) 
 MsgBox ("A new chart was added to: " & Wb.Name & " of type: " & Ch.Type) 
End Sub
```

