# Saving Documents as Web Pages

## Business Description
In Microsoft Excel, you can save a workbook, worksheet, chart, range, query table, PivotTable report, print area, or AutoFilter range to a Web page. You can also edit HTML files directly in Excel.

## Behavior
In Microsoft Excel, you can save a workbook, worksheet, chart, range, query table, PivotTable report, print area, or AutoFilter range to a Web page. You can also edit HTML files directly in Excel.

## Example Usage
```vba
ActiveWorkbook.SaveAs _ 
 Filename:="C:\Reports\myfile.htm", _ 
 FileFormat:=xlHTML
```