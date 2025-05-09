# PivotTable TableStyle2 Property

## Business Description
The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable. Read/write.

## Behavior
TheTableStyle2property specifies the PivotTable style currently applied to the PivotTable. Read/write.

## Example Usage
```vba
Sub ApplyingStyle() 
 
 ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight17" 
 
End Sub
```