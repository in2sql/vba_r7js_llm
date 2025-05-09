# Using Events with Excel Objects

## Business Description
You can write event procedures in Microsoft Excel at the worksheet, chart, query table, workbook, or application level.

## Behavior
You can write event procedures in Microsoft Excel at the worksheet, chart, query table, workbook, or application level. For example, theActivateevent occurs at the sheet level, and theSheetActivateevent is available at both the workbook and application levels. TheSheetActivateevent for a workbook occurs when any sheet in the workbook is activated. At the application level, theSheetActivateevent occurs when any sheet in any open workbook is activated.

## Example Usage
```vba
Application.EnableEvents = False 
ActiveWorkbook.Save 
Application.EnableEvents = True
```