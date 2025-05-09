# PivotField IsMemberProperty Property

## Business Description
Returns True when the PivotField contains member properties. Read-only Boolean.

## Behavior
ReturnsTruewhen the PivotField contains member properties. Read-onlyBoolean.

## Example Usage
```vba
Sub CheckForMembers() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(1) 
 
 ' Determine if member properties exist and notify user. 
 If pvtField.IsMemberProperty= True Then 
 MsgBox "The PivotField contains member properties." 
 Else 
 MsgBox "The PivotField does not contain member properties." 
 End If 
 
End Sub
```