# SpellingOptions IgnoreMixedDigits Property

## Business Description
False instructs Microsoft Excel to check for mixed digits, True instructs Excel to ignore mixed digits when checking spelling. Read/write Boolean.

## Behavior
Falseinstructs Microsoft Excel to check for mixed digits,Trueinstructs Excel to ignore mixed digits when checking spelling. Read/writeBoolean.

## Example Usage
```vba
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreMixedDigits= True Then 
 MsgBox "Spelling options for checking mixed digits is disabled." 
 Else 
 MsgBox "Spelling options for checking mixed digits is enabled." 
 End If 
 
End Sub
```