# Workbook FileFormat Property

## Business Description
Returns the file format and/or type of the workbook. Read-only XlFileFormat.

## Behavior
Returns the file format and/or type of the workbook.  Read-onlyXlFileFormat.

## Example Usage
```vba
If ActiveWorkbook.FileFormat= xlExcel9795 Then 
 ActiveWorkbook.SaveAs fileFormat:=xlExcel12 
End If
```