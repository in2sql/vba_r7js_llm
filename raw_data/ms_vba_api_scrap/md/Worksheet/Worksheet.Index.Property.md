# Worksheet Index Property

## Business Description
Returns a Long value that represents the index number of the object within the collection of similar objects.

## Behavior
Returns aLongvalue that represents the index number of the object within the collection of similar objects.

## Example Usage
```vba
Sub DisplayTabNumber() 
 Dim strSheetName as String 
 
 strSheetName = InputBox("Type a sheet name, such as Sheet4.") 
 
 MsgBox "This sheet is tab number " & Sheets(strSheetName).Index 
End Sub
```