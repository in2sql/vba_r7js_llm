# Phonetics Start Property

## Business Description
Returns the position that represents the first character of a phonetic text string in the specified cell. Read-only Long.

## Behavior
Returns the position that represents the first character of a phonetic text string in the specified cell. Read-onlyLong.

## Example Usage
```vba
ActiveCell.FormulaR1C1 = "東京都渋谷区代々木" 
ActiveCell.Phonetics.Add Start:=1, Length:=3, Text:="トウキョウト" 
ActiveCell.Phonetics.Add Start:=4, Length:=3, Text:="シブヤク" 
MsgBox ActiveCell.Phonetics(2).Start
```