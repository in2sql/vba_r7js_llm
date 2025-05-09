# PivotTable MDX Property

## Business Description
Returns a String indicating the Multidimensional Expression (MDX) that would be sent to the provider to populate the current PivotTable view. Read-only.

## Behavior
Returns aStringindicating the Multidimensional Expression (MDX) that would be sent to the provider to populate the current PivotTable view. Read-only.

## Example Usage
```vba
Sub CheckMDX() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 MsgBox "The MDX string for the PivotTable is: " & _ 
 pvtTable.MDX 
 
End Sub
```