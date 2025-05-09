# Workbook BreakLink Method

## Business Description
Converts formulas linked to other Microsoft Excel sources or OLE sources to values.

## Behavior
Converts formulas linked to other Microsoft Excel sources or OLE sources to values.

## Example Usage
```vba
Sub UseBreakLink() 
 
 Dim astrLinks As Variant 
 
 ' Define variable as an Excel link type. 
 astrLinks = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks) 
 
 ' Break the first link in the active workbook. 
 ActiveWorkbook.BreakLink_ 
 Name:=astrLinks(1), _ 
 Type:=xlLinkTypeExcelLinks 
 
End Sub
```