# Workbook RunAutoMacros Method

## Business Description
Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility.

## Behavior
Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility. For new Visual Basic code, you should use the Open, Close, Activate and Deactivate events instead of these macros.

## Example Usage
```vba
Workbooks.Open "ANALYSIS.XLS" 
ActiveWorkbook.RunAutoMacrosxlAutoOpen
```