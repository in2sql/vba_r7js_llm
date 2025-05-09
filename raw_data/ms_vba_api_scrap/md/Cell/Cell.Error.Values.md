# Cell Error Values

## Business Description
You can insert a cell error value into a cell or test the value of a cell for an error value by using the CVErr function. The cell error values can be one of the following XlCVError constants.

## Behavior
You can insert a cell error value into a cell or test the value of a cell for an error value by using theCVErrfunction. The cell error values can be one of the followingXlCVErrorconstants.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
If IsError(ActiveCell.Value) Then 
 errval = ActiveCell.Value 
 Select Case errval 
 Case CVErr(xlErrDiv0) 
 MsgBox "#DIV/0! error" 
 Case CVErr(xlErrNA) 
 MsgBox "#N/A error" 
 Case CVErr(xlErrName) 
 MsgBox "#NAME? error" 
 Case CVErr(xlErrNull) 
 MsgBox "#NULL! error" 
 Case CVErr(xlErrNum) 
 MsgBox "#NUM! error" 
 Case CVErr(xlErrRef) 
 MsgBox "#REF! error" 
 Case CVErr(xlErrValue) 
 MsgBox "#VALUE! error" 
 Case Else 
 MsgBox "This should never happen!!" 
 End Select 
End If
```