# Workbook BuiltinDocumentProperties Property

## Business Description
Returns a DocumentPropertieshttp://msdn.microsoft.com/library/90d42786-7d9a-b604-dbdf-88db41cbe69b(Office.15).aspx collection that represents all the built-in document properties for the specified workbook. Read-only.

## Behavior
Returns aDocumentPropertiescollection that represents all the built-in document properties for the specified workbook. Read-only.

## Example Usage
```vba
rw = 1 
Worksheets(1).Activate 
For Each p In ActiveWorkbook.BuiltinDocumentPropertiesCells(rw, 1).Value = p.Name 
    rw = rw + 1 
Next
```