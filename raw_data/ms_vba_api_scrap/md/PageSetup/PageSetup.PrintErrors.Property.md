# PageSetup PrintErrors Property

## Business Description
Sets or returns an XlPrintErrors contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write .

## Behavior
Sets or returns anXlPrintErrorscontstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write .

## Example Usage
```vba
Sub UsePrintErrors() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Create a formula that returns an error value. 
 Range("A1").Value = 1 
 Range("A2").Value = 0 
 Range("A3").Formula = "=A1/A2" 
 
 ' Change print errors to display dashes. 
 wksOne.PageSetup.PrintErrors= xlPrintErrorsDash 
 
 ' Use the Print Preview window to see the dashes used for print errors. 
 ActiveWindow.SelectedSheets.PrintPreview 
 
End Sub
```