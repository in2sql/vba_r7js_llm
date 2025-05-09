# PivotFormula StandardFormula Property

## Business Description
Returns or sets a String specifying formulas with standard English (United States) formatting. Read/write.

## Behavior
Returns or sets aStringspecifying formulas with standard English (United States) formatting. Read/write.

## Example Usage
```vba
Sub UseStandardFomula() 
 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Change calculated field of decimals by adding '10'. 
 pvtTable.CalculatedFields.Item(1).StandardFormula= "Decimals + 10" 
 
End Sub
```