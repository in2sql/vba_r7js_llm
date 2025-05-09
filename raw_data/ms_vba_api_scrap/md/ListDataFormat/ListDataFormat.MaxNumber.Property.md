# ListDataFormat MaxNumber Property

## Business Description
Returns a Variant containing the maximum value allowed in this field in the list column. Read-only Variant.

## Behavior
Returns aVariantcontaining the maximum value allowed in this field in the list column. Read-onlyVariant.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MaxNumber
```