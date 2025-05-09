# PivotField CurrentPageList Property

## Business Description
Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/write Variant.

## Behavior
Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/writeVariant.

## Example Usage
```vba
Sub UseCurrentPageList() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields("[Product]") 
 
 ' To avoid run-time errors set the following property to True. 
 pvtTable.CubeFields("[Product]").EnableMultiplePageItems = True 
 
 ' Set the page list to "Food". 
 pvtField.CurrentPageList= "[Product].[All Products].[Food]" 
 
End Sub
```