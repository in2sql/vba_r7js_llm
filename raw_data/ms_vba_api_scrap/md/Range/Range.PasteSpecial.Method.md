# Range PasteSpecial Method

## Business Description
Pastes a Range from the Clipboard into the specified range.

## Behavior
Pastes aRangefrom the Clipboard into the specified range.

## Example Usage
```vba
With Worksheets("Sheet1") 
 .Range("C1:C5").Copy 
 .Range("D1:D5").PasteSpecial_ 
 Operation:=xlPasteSpecialOperationAdd 
End With
```