# Sort Object

## Business Description
Represents a sort of a range of data.

## Behavior
Represents a sort of a range of data.

## Example Usage
```vba
Sub SortData() 
 
 'Building data to sort on the active sheet. 
 Range("A1").Value = "Name" 
 Range("A2").Value = "Bill" 
 Range("A3").Value = "Rod" 
 Range("A4").Value = "John" 
 Range("A5").Value = "Paddy" 
 Range("A6").Value = "Kelly" 
 Range("A7").Value = "William" 
 Range("A8").Value = "Janet" 
 Range("A9").Value = "Florence" 
 Range("A10").Value = "Albert" 
 Range("A11").Value = "Mary" 
 MsgBox "The list is out of order. Hit Ok to continue...", vbInformation 
 
 'Selecting a cell within the range. 
 Range("A2").Select 
 
 'Applying sort. 
 With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort 
 .SortFields.Clear 
 .SortFields.Add Key:=Range("A2:A11"), _ 
 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 
 .SetRange Range("A1:A11") 
 .Header = xlYes 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
 End With 
 MsgBox "Sort complete.", vbInformation 
 
End Sub
```