# Workbook ExportAsFixedFormat Method

## Business Description
The ExportAsFixedFormat method is used to publish a workbook to either the PDF or XPS format.

## Behavior
TheExportAsFixedFormatmethod is used to publish a workbook to either the PDF or XPS format.

## Example Usage
```vba
ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF FileName:="sales.pdf" Quality:=xlQualityStandard DisplayFileAfterPublish:=True
```