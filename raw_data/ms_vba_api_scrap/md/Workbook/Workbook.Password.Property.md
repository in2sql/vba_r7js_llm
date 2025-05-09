# Workbook Password Property

## Business Description
Returns or sets the password that must be supplied to open the specified workbook. Read/write String.

## Behavior
Returns or sets the password that must be supplied to open the specified workbook. Read/writeString.

## Example Usage
```vba
Sub UsePassword() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.Workbooks.Open("C:\Password.xls") 
 
 wkbOne.Password= InputBox ("Enter Password") 
 wkbOne.Close 
 
End Sub
```