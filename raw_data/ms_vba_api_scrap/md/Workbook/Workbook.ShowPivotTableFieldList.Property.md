# Workbook ShowPivotTableFieldList Property

## Business Description
True (default) if the PivotTable field list can be shown. Read/write Boolean.

## Behavior
True(default) if the PivotTable field list can be shown. Read/writeBoolean.

## Example Usage
```vba
Sub UseShowPivotTableFieldList() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.ActiveWorkbook 
 
 'Determine PivotTable field list setting. 
 If wkbOne.ShowPivotTableFieldList= True Then 
 MsgBox "The PivotTable field list can be viewed." 
 Else 
 MsgBox "The PivotTable field list cannot be viewed." 
 End If 
 
End Sub
```