# Workbook WritePassword Property

## Business Description
Returns or sets a String for the write password of a workbook. Read/write.

## Behavior
Returns or sets aStringfor the write password of a workbook. Read/write.

## Example Usage
```vba
Sub UseWritePassword() 
 
 Dim strPassword As String 
 
 strPassword = InputBox ("Enter the password") 
 
 ' Set password to a string if allowed. 
 If ActiveWorkbook.WriteReserved = False Then 
 ActiveWorkbook.WritePassword= strPassword 
 End If 
 
End Sub
```