# Range AutoComplete Method

## Business Description
Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.

## Behavior
Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.

## Example Usage
```vba
s = Worksheets(1).Range("A5").AutoComplete("Ap") 
If Len(s) > 0 Then 
 MsgBox "Completes to " & s 
Else 
 MsgBox "Has no completion" 
End If
```