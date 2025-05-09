# ListDataFormat DefaultValue Property

## Business Description
Returns Variant representing the default data type value for a new row in a column. The Nothing object is returned when the schema does not specify a default value. Read-only Variant.

## Behavior
ReturnsVariantrepresenting the default data type value for a new row in a  column. TheNothingobject is returned when the schema does not specify a default value. Read-onlyVariant.

## Example Usage
```vba
Sub ShowDefaultSetting() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 If IsNull(objListCol.ListDataFormat.DefaultValue) Then 
 MsgBox "No default value specified." 
 Else 
 MsgBox objListCol.ListDataFormat.DefaultValueEnd If 
End Sub
```