# PivotField AddPageItem Method

## Business Description
Adds an additional item to a multiple item page field.

## Behavior
Adds an additional item to a multiple item page field.

## Example Usage
```vba
Sub UseAddPageItem() 
 
 ' The source is an OLAP database and you can manually reorder items. 
 ActiveSheet.PivotTables(1).CubeFields("[Product]"). _ 
 EnableMultiplePageItems = True 
 
 ' Add the page item titled "[Product].[All Products].[Food].[Eggs]". 
 ActiveSheet.PivotTables(1).PivotFields("[Product]").AddPageItem( _ 
 "[Product].[All Products].[Food].[Eggs]") 
 
End Sub
```