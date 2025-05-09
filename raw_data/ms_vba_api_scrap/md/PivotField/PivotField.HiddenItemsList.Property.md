# PivotField HiddenItemsList Property

## Business Description
Returns or sets a Variant specifying an array of strings that are hidden items for a PivotTable field. Read/write.

## Behavior
Returns or sets aVariantspecifying an array of strings that are hidden items for a PivotTable field. Read/write.

## Example Usage
```vba
Sub UseHiddenItemsList() 
 
 ActiveSheet.PivotTables(1).PivotFields(1).HiddenItemsList= _ 
 Array("[Product].[All Products].[Food]", _ 
 "[Product].[All Products].[Drink]") 
 
End Sub
```