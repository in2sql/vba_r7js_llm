# SpellingOptions DictLang Property

## Business Description
Selects the dictionary language used when Microsoft Excel performs spelling checks. Read/write Long.

## Behavior
Selects the dictionary language used when Microsoft Excel performs spelling checks. Read/writeLong.

## Example Usage
```vba
Sub LanguageSpellCheck() 
 
 With Application.SpellingOptions 
 .DictLang= 1033 ' United States English language number. 
 .UserDict = "CUSTOM.DIC" 
 End With 
 
End Sub
```