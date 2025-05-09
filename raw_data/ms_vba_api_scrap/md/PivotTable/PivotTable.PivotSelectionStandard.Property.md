# PivotTable PivotSelectionStandard Property

## Business Description
Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English (United States) settings. Read/write.

## Behavior
Returns or sets aStringindicating the PivotTable selection in standard PivotTable report format using English (United States) settings. Read/write.

## Example Usage
```vba
Sub CheckPivotSelectionStandard() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 pvtTable.PivotSelectionStandard= "1.57" 
 Selection.Insert 
 
End Sub
```