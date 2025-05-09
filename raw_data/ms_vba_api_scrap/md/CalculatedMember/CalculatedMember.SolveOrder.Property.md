# CalculatedMember SolveOrder Property

## Business Description
Returns a Long specifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.

## Behavior
Returns aLongspecifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.

## Example Usage
```vba
Sub CheckSolveOrder() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine solve order and notify user. 
 If pvtTable.CalculatedMembers.Item(1).SolveOrder= 0 Then 
 MsgBox "The solve order is set to the default value." 
 Else 
 MsgBox "The solve order is not set to the default value." 
 End If 
 
End Sub
```