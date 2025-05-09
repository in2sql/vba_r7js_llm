# How to: Create a Workbook

## Business Description
To create a workbook in Visual Basic, use the Add method. The following procedure creates a workbook. Microsoft Excel automatically names the workbook BookN, where N is the next available number. The new workbook becomes the active workbook.

## Behavior
To create a workbook in Visual Basic, use theAddmethod. The following procedure creates a workbook. Microsoft Excel automatically names the workbook BookN, whereNis the next available number. The new workbook becomes the active workbook.

## Example Usage
```vba
Sub AddOne() 
 Workbooks.Add 
End Sub
```