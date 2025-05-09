# Workbook CustomDocumentProperties Property

## Business Description
Returns or sets a DocumentPropertieshttp://msdn.microsoft.com/library/90d42786-7d9a-b604-dbdf-88db41cbe69b(Office.15).aspx collection that represents all the custom document properties for the specified workbook.

## Behavior
Returns or sets aDocumentPropertiescollection that represents all the custom document properties for the specified workbook.

## Example Usage
```vba
rw = 1 
Worksheets(1).Activate 
For Each p In ActiveWorkbook.CustomDocumentPropertiesCells(rw, 1).Value = p.Name 
    Cells(rw, 2).Value = p.Value 
    rw = rw + 1 
Next
```