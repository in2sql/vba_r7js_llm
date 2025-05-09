# IconCriterion Icon Property

## Business Description
Returns or specifies the icon for a criterion in an icon set conditional formatting rule. Read/write

## Behavior
Returns or specifies the icon for a criterion in an icon set conditional formatting rule.  Read/write

## Example Usage
```vba
Range("A1:A10").Select 
Selection.FormatConditions.AddIconSetCondition 
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 
With Selection.FormatConditions(1) 
 .ReverseOrder = False 
 .ShowIconOnly = False 
 .IconSet = ActiveWorkbook.IconSets(xl4Arrows) 
End With 
 
With Selection.FormatConditions(1).IconCriteria(1) 
 .Icon= xlIconRedCross 
End With 
 
With Selection.FormatConditions(1).IconCriteria(2) 
 .Type = xlConditionValuePercent 
 .Value = 25 
 .Operator = 7 
End With 
 
With Selection.FormatConditions(1).IconCriteria(3) 
 .Type = xlConditionValuePercent 
 .Value = 50 
 .Operator = 7 
 .Icon= xlIconYellowTrafficLight 
End With 
 
With Selection.FormatConditions(1).IconCriteria(4) 
 .Type = xlConditionValuePercent 
 .Value = 75 
 .Operator = 7 
End With
```