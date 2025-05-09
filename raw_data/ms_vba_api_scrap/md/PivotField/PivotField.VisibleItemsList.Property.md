# PivotField VisibleItemsList Property

## Business Description
Returns or sets a Variant specifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.

## Behavior
Returns or sets aVariantspecifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Country]").VisibleItemsList = Array("[Customer].[Customer Geography].[Country].&[Australia]") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[State-Province]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[City]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Postal Code]").VisibleItemsList = Array("") 
ActiveSheet.PivotTables("PivotTable2").PivotFields("[Customer].[Customer Geography] & _ 
.[Full Name]").VisibleItemsList = Array("")
```