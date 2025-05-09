# SpellingOptions UserDict Property

## Business Description
Instructs Microsoft Excel to create a custom dictionary to which new words can be added to, when performing spelling checks on a worksheet. Read/write String.

## Behavior
Instructs Microsoft Excel to create a custom dictionary to which new words can be added to, when performing spelling checks on a worksheet. Read/writeString.

## Example Usage
```vba
Sub SpecialWord() 
 
 Application.SpellingOptions.UserDict= "Custom1.dic" 
 MsgBox "The custom dictionary is currently set to: " _ 
 & Application.SpellingOptions.UserDictEnd Sub
```