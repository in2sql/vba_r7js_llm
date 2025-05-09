# PivotCell Object

## Business Description
Represents a cell in a PivotTable report.

## Behavior
Represents a cell in a PivotTable report.

## Example Usage
```vba
Sub CheckPivotCellType() 
 
 On Error GoTo Not_In_PivotTable 
 
 ' Determine if cell A5 is a data item in the PivotTable. 
 If Application.Range("A5").PivotCell.PivotCellType = xlPivotCellValue Then 
 MsgBox "The PivotCell at A5 is a data item." 
 Else 
 MsgBox "The PivotCell at A5 is not a data item." 
 End If 
 Exit Sub 
 
Not_In_PivotTable: 
 MsgBox "The chosen cell is not in a PivotTable." 
 
End Sub
```