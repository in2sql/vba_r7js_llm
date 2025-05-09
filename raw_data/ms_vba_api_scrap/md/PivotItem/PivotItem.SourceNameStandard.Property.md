# PivotItem SourceNameStandard Property

## Business Description
Returns a String that represents the PivotTable items' source name in standard English (United States) format settings. Read-only.

## Behavior
Returns aStringthat represents the PivotTable items' source name in standard English (United States) format settings. Read-only.

## Example Usage
```vba
Sub CheckSourceNameStandard() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 Dim pvtItem As PivotItem 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(5) 
 Set pvtItem = pvtField.PivotItems(6) 
 
 ' Display source name. 
 MsgBox "The source name is: " & pvtItem.SourceNameStandardEnd Sub
```