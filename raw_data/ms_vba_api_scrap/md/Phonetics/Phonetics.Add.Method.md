# Phonetics Add Method

## Business Description
Adds phonetic text to the specified cell.

## Behavior
Adds phonetic text to the specified cell.

## Example Usage
```vba
ActiveCell.FormulaR1C1 = "東京都渋谷区代々木" 
ActiveCell.Phonetics.AddStart:=1, Length:=3, Text:="トウキョウト" 
ActiveCell.Phonetics.AddStart:=4, Length:=3, Text:="シブヤク" 
ActiveCell.Phonetics.CharacterType = xlHiragana 
ActiveCell.Phonetics.Font.Color = vbBlue 
ActiveCell.Phonetics.Visible = True
```