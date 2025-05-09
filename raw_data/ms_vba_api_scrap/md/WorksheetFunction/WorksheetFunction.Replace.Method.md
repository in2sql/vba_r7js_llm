# WorksheetFunction Replace Method

## Business Description
Replaces part of a text string, based on the number of characters you specify, with a different text string.

## Behavior
Replaces part of a text string, based on the number of characters you specify, with a different text string.

## Example Usage
```vba
Sub UseReplace() 
 
 Dim strCurrent As String 
 Dim strReplaced As String 
 
 strCurrent = "abcdef" 
 
 ' Notify user and display current string. 
 MsgBox "The current string is: " & strCurrent 
 
 ' Replace "cd" with "-". 
 strReplaced = Application.WorksheetFunction.Replace_ 
 (Arg1:=strCurrent, Arg2:=3, _ 
 Arg3:=2, Arg4:="-") 
 
 ' Notify user and display replaced string. 
 MsgBox "The replaced string is: " & strReplaced 
 
End Sub
```