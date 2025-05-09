# ConditionValue Object

## Business Description
Represents how the shortest bar or longest bar is evaluated for a data bar conditional formatting rule.

## Behavior
Represents how the shortest bar or longest bar is evaluated for a data bar conditional formatting rule.

## Example Usage
```vba
Sub CreateDataBarCF() 
 
 Dim cfDataBar As Databar 
 
 'Create a range of data with a couple of extreme values 
 With ActiveSheet 
 .Range("D1") = 1 
 .Range("D2") = 45 
 .Range("D3") = 50 
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8") 
 .Range("D9") = 500 
 End With 
 
 Range("D1:D9").Select 
 
 'Create a data bar with default behavior 
 Set cfDataBar = Selection.FormatConditions.AddDatabar 
 MsgBox "Because of the extreme values, middle data bars are very similar" 
 
 'The MinPoint and MaxPoint properties return aConditionValueobject 
 'which you can use to change threshold parameters 
 cfDataBar.MinPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=5 
 cfDataBar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=75 
 
End Sub
```