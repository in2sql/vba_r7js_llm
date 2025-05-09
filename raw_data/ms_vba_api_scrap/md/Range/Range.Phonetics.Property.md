# Range Phonetics Property

## Business Description
Returns the Phonetics collection of the range. Read only.

## Behavior
Returns thePhoneticscollection of the range. Read only.

## Example Usage
```vba
Set objPhon = ActiveCell.PhoneticsWith objPhon 
 For Each objPhonItem in objPhon 
 MsgBox "Phonetic object: " & .Text 
 Next 
End With
```