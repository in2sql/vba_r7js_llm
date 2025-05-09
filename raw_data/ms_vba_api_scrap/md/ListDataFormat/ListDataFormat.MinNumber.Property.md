# ListDataFormat MinNumber Property

## Business Description
Returns a Variant containing the minimum value allowed in this field in the list column. This can be a negative floating point number. Read-only Variant.

## Behavior
Returns aVariantcontaining the minimum value allowed in this field in the list column.  This can be a negative floating point number. Read-onlyVariant.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MinNumber
```