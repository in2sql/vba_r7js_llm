# Worksheet Object Events

## Business Description
Events on sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and click View Code on the shortcut menu. Select one of the following events from the Procedure drop-down list box.

## Behavior
Events on sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and clickView Codeon the shortcut menu. Select one of the following events from theProceduredrop-down list box.

## Example Usage
```vba
Private Sub Worksheet_Calculate() 
    Columns("A:F").AutoFit 
End Sub
```