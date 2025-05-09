# PivotCache SaveAsODC Method

## Business Description
Saves the PivotTable cache source as an Microsoft Office Data Connection file.

## Behavior
Saves the PivotTable cache source as an Microsoft Office Data Connection file.

## Example Usage
```vba
Sub UseSaveAsODC() 
 
 Application.ActiveWorkbook.PivotCaches.Item(1).SaveAsODC("ODCFile") 
 
End Sub
```