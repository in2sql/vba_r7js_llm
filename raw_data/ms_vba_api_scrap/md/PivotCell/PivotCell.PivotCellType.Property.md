# PivotCell PivotCellType Property

## Business Description
Returns one of the XlPivotCellType constants that identifies the PivotTable entity the cell corresponds to. Read-only.

## Behavior
Returns one of theXlPivotCellTypeconstants that identifies the PivotTable entity the cell corresponds to. Read-only.

## Example Usage
```vba
Sub CheckPivotCellType() 
 
 On Error GoTo Not_In_PivotTable 
 
 ' Determine if cell A5 is a data item in the PivotTable. 
 If Application.Range("A5").PivotCell.PivotCellType= xlPivotCellValue Then 
 MsgBox "The cell at A5 is a data item." 
 Else 
 MsgBox "The cell at A5 is not a data item." 
 End If 
 Exit Sub 
 
Not_In_PivotTable: 
 MsgBox "The chosen cell is not in a PivotTable." 
 
End Sub
```