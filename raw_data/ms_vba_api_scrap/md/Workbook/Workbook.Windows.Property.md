# Workbook Windows Property

## Business Description
Returns a Windows collection that represents all the windows in the specified workbook. Read-only Windows object.

## Behavior
Returns aWindowscollection that represents all the windows in the specified workbook. Read-onlyWindowsobject.

## Example Usage
```vba
ActiveWorkbook.Windows(1).Caption = "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```