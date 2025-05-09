# CubeField CreatePivotFields Method

## Business Description
The CreatePivotFields method enables users to apply a filter to PivotFields not yet added to the PivotTable by creating the corresponding PivotField object.

## Behavior
TheCreatePivotFieldsmethod enables users to apply a filter to PivotFields not yet added to the PivotTable by creating the correspondingPivotFieldobject.

## Example Usage
```vba
Sub FilterFieldBeforeAddingItToPivotTable() 
 ActiveSheet.PivotTables("PivotTable1").CubeFields("[Date].[Fiscal]").CreatePivotFields 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Fiscal Year]").VisibleItemsList = 
 
 Array("[Date].[Fiscal].[Fiscal Year].&[2003]", "[Date].[Fiscal].[Fiscal Year].&[2004]", "[Date].[Fiscal].[Fiscal Year].&[2005]") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields( _ 
 "[Date].[Fiscal].[Fiscal Semester]").VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields( _ 
 "[Date].[Fiscal].[Fiscal Quarter]").VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Month]"). _ 
 VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Date]"). _ 
 VisibleItemsList = Array("") 
End Sub
```