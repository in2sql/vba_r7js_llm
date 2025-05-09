# PivotValueCell PivotCell Property

## Business Description
Returns the PivotCell that specifies the location of the PivotValueCell. Read-only

## Behavior
Returns thePivotCell Object (Excel)that specifies the location of thePivotValueCell. Read-only

## Example Usage
```vba
Sub GetMDX()
   'Get the MDX query for a particular PivotCell in a workbook level PivotTable
   MsgBox "The MDX for the PivotCell 1, 1: " & _
   ThisWorkbook.PivotTables(1).PivotValueCell(1, 1).PivotCell.MDX
End Sub
```