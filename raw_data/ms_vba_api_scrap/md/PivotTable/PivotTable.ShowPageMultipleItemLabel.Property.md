# PivotTable ShowPageMultipleItemLabel Property

## Business Description
When set to True (default), "(Multiple Items)" will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view. Read/write Boolean.

## Behavior
When set toTrue(default), "(Multiple Items)" will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view. Read/writeBoolean.

## Example Usage
```vba
Sub UseShowPageMultipleItemLabel() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if multiple items are allowed. 
 If pvtTable.ShowPageMultipleItemLabel= True Then 
 MsgBox "The words 'Multiple Items' can be displayed." 
 Else 
 MsgBox "The words 'Multiple Items' cannot be displayed." 
 End If 
 
End Sub
```