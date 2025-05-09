# Top10 Object

## Business Description
Represents a top ten visual of a conditional formatting rule. Applying a color to a range helps you see the value of a cell relative to other cells.

## Behavior
Represents a top ten visual of a conditional formatting rule. Applying a color to a range helps you see the value of a cell relative to other cells.

## Example Usage
```vba
Sub Top10CF() 
 
' Building data 
 Range("A1").Value = "Name" 
 Range("B1").Value = "Number" 
 Range("A2").Value = "Agent1" 
 Range("A2").AutoFill Destination:=Range("A2:A26"), Type:=xlFillDefault 
 Range("B2:B26").FormulaArray = "=INT(RAND()*101)" 
 Range("B2:B26").Select 
 
' Applying Conditional Formatting Top 10 
 Selection.FormatConditions.AddTop10 
 Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 With Selection.FormatConditions(1) 
 .TopBottom = xlTop10Top 
 .Rank = 10 
 .Percent = False 
 End With 
 
' Applying color fill 
 With Selection.FormatConditions(1).Font 
 .Color = -16752384 
 .TintAndShade = 0 
 End With 
 With Selection.FormatConditions(1).Interior 
 .PatternColorIndex = xlAutomatic 
 .Color = 13561798 
 .TintAndShade = 0 
 End With 
MsgBox "Added Top10 Conditional Format. Press F9 to update values.", vbInformation 
 
End Sub
```