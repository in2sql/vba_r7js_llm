# Phonetics Length Property

## Business Description
Returns a Long value that represents the number of characters of phonetic text from the position you've specified with the Start property.

## Behavior
Returns aLongvalue that represents the number of characters of phonetic text from the position you've specified with theStartproperty.

## Example Usage
```vba
ActiveCell.FormulaR1C1 = "東京都渋谷区代々木" 
ActiveCell.Phonetics.Add Start:=1, Length:=3, Text:="トウキョウト" 
ActiveCell.Phonetics.Add Start:=4, Length:=3, Text:="シブヤク" 
MsgBox ActiveCell.Phonetics(2).Length
```