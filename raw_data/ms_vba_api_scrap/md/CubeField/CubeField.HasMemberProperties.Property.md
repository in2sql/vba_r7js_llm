# CubeField HasMemberProperties Property

## Business Description
Returns True when there are member properties specified to be displayed for the cube field. Read-only Boolean.

## Behavior
ReturnsTruewhen there are member properties specified to be displayed for the cube field. Read-onlyBoolean.

## Example Usage
```vba
Sub UseHasMemberProperties() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine if there are member properties to be displayed. 
 If cbeField.HasMemberProperties= True Then 
 MsgBox "There are member properties to be displayed." 
 Else 
 MsgBox "There are no member properties to be displayed." 
 End If 
 
End Sub
```