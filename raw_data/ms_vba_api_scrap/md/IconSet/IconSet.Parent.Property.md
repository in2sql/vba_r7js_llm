# IconSet object (Excel)

## Description
This browser is no longer supported.

## Remarks
The IconSet object is a child object of the IconSets collection.

## Example
```vba
Sub CreateIconSetCF() 
 
 Dim cfIconSet As IconSetCondition 
 
 'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 55 
 .Range("C2") = 92 
 .Range("C3") = 88 
 .Range("C4") = 77 
 .Range("C5") = 66 
 .Range("C6") = 93 
 .Range("C7") = 76 
 .Range("C8") = 80 
 .Range("C9") = 79 
 .Range("C10") = 83 
 .Range("C11") = 66 
 .Range("C12") = 74 
 End With 
 
 Range("C1:C12").Select 
 
 'Create an icon set conditional format for the created sample data range 
 Set cfIconSet = Selection.FormatConditions.AddIconSetCondition 
 
 'Change the icon set to a 5-arrow icon set 
 cfIconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows) 
 
 'The IconCriterion collection contains all of IconCriteria 
 'By indexing into the collection you can modify each criteria 
 
 With cfIconSet.IconCriteria(1) 
 .Type = xlConditionValueNumber 
 .Value = 0 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(2) 
 .Type = xlConditionValueNumber 
 .Value = 60 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(3) 
 .Type = xlConditionValueNumber 
 .Value = 70 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(4) 
 .Type = xlConditionValueNumber 
 .Value = 80 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(5) 
 .Type = xlConditionValueNumber 
 .Value = 90 
 .Operator = 7 
 End With 
 
End Sub
```

