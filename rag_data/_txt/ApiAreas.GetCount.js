**English:** This code accesses the active worksheet, sets the value "1" in cells B1 to D1, selects the range, retrieves the areas, counts them, and updates cells A5 and B5 with a message and the count respectively.

**Russian:** Этот код обращается к активному листу, устанавливает значение "1" в ячейки B1 до D1, выбирает диапазон, получает области, считает их количество и обновляет ячейки A5 и B5 сообщением и количеством соответственно.

```javascript
// This example shows how to get a value that represents the number of objects in the collection.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1:D1");
oRange.SetValue("1");
oRange.Select();
var oAreas = oRange.GetAreas();
var nCount = oAreas.GetCount();
oRange = oWorksheet.GetRange('A5');
oRange.SetValue("The number of ranges in the areas: ");
oRange.AutoFit(false, true);
oWorksheet.GetRange('B5').SetValue(nCount); 
```

```vba
' This example shows how to get a value that represents the number of objects in the collection.
Sub CountRanges()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oAreas As Areas
    Dim nCount As Long
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1:D1
    Set oRange = oWorksheet.Range("B1:D1")
    
    ' Set the value "1" in the range
    oRange.Value = "1"
    
    ' Select the range
    oRange.Select
    
    ' Get the areas (non-contiguous ranges)
    Set oAreas = oRange.Areas
    
    ' Get the count of areas
    nCount = oAreas.Count
    
    ' Get range A5 and set a message
    Set oRange = oWorksheet.Range("A5")
    oRange.Value = "The number of ranges in the areas: "
    
    ' AutoFit columns
    oRange.EntireColumn.AutoFit
    
    ' Set the count in B5
    oWorksheet.Range("B5").Value = nCount
End Sub
```