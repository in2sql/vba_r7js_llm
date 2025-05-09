# SpellingOptions HebrewModes Property

## Business Description
Returns or sets the mode for the Hebrew spelling checker. Read/write XlHebrewModes.

## Behavior
Returns or sets the mode for the Hebrew spelling checker. Read/writeXlHebrewModes.

## Example Usage
```vba
Sub CheckHebrewMode() 
 
 ' Determine the Hebrew spelling mode setting and notify user. 
 Select Case Application.SpellingOptions.HebrewModesCase xlHebrewFullScript 
 MsgBox "The Hebrew spelling mode setting is Full Script." 
 Case xlHebrewMixedAuthorizedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Authorized Script." 
 Case xlHebrewMixedScript 
 MsgBox "The Hebrew spelling mode setting is Mixed Script." 
 Case xlHebrewPartialScript 
 MsgBox "The Hebrew spelling mode setting is Partial Script." 
 End Select 
 
End Sub
```