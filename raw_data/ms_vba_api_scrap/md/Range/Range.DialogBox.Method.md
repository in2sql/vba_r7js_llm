# Range DialogBox Method

## Business Description
Displays a dialog box defined by a dialog box definition table on a Microsoft Excel 4.0 macro sheet. Returns the number of the chosen control, or returns False if the user clicks the Cancel button.

## Behavior
Displays a dialog box defined by a dialog box definition table on a Microsoft Excel 4.0 macro sheet. Returns the number of the chosen control, or returnsFalseif the user clicks theCancelbutton.

## Example Usage
```vba
Set dialogRange = Excel4MacroSheets("Macro1").Range("myDialogBox") 
result = dialogRange.DialogBoxMsgBox result
```