# PivotField DatabaseSort Property

## Business Description
When set to True, manual repositioning of items in a PivotTable field is allowed. Returns True, if the field has no manually positioned items. Read/write Boolean.

## Behavior
When set toTrue, manual repositioning of items in a PivotTable field is allowed. ReturnsTrue, if the field has no manually positioned items. Read/writeBoolean.

## Example Usage
```vba
Sub UseDatabaseSort() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields("[Product].[Product Family]") 
 
 ' Determine source type for the PivotTable report. 
 If pvtField.DatabaseSort= True Then 
 MsgBox "The source is OLAP; you can manually reorder items." 
 Else 
 MsgBox "The data source might not be OLAP." 
 End If 
 
End Sub
```