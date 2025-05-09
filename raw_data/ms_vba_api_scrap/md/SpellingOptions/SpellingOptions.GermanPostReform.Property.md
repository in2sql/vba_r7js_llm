# SpellingOptions GermanPostReform Property

## Business Description
True to check the spelling of words using the German post-reform rules. False cancels this feature. Read/write Boolean.

## Behavior
Trueto check the spelling of words using the German post-reform rules.Falsecancels this feature. Read/writeBoolean.

## Example Usage
```vba
Sub SpellingCheck() 
 
 ' Determine if spelling check for German words is using post-reform rules. 
 If Application.SpellingOptions.GermanPostReform= False Then 
 Application.SpellingOptions.GermanPostReform= True 
 MsgBox "German words will now use post-reform rules." 
 Else 
 MsgBox "German words using post-reform rules has already been set." 
 End If 
 
End Sub
```