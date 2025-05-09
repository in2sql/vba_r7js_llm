# Code Description / Описание кода

**English:** This code demonstrates how to get the parent object for a specified range in an Excel worksheet, set values, select ranges, autofit columns, and retrieve and display the parent object and its class type.

**Russian:** Этот код демонстрирует, как получить родительский объект для указанного диапазона на листе Excel, установить значения, выбрать диапазоны, автоматически подогнать столбцы и получить и отобразить родительский объект и его тип класса.

```vba
' VBA code equivalent to the OnlyOffice JS example
Sub GetParentObjectExample()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1:D1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1:D1")
    
    ' Set value to "1"
    oRange.Value = "1"
    
    ' Select the range
    oRange.Select
    
    ' Get the areas of the range
    Dim oAreas As Areas
    Set oAreas = oRange.Areas
    
    ' Get the parent of the areas
    Dim oParent As Object
    Set oParent = oAreas.Parent
    
    ' Get the class type of the parent
    Dim sType As String
    sType = TypeName(oParent)
    
    ' Set value in A4
    Set oRange = oWorksheet.Range("A4")
    oRange.Value = "The areas parent:"
    
    ' Autofit columns A and B
    oWorksheet.Columns("A:B").AutoFit
    
    ' Paste the parent object reference in B4
    oWorksheet.Range("B4").Value = oParent.Name
    
    ' Set value in A5
    Set oRange = oWorksheet.Range("A5")
    oRange.Value = "The type of the areas parent:"
    
    ' Autofit columns A and B again
    oWorksheet.Columns("A:B").AutoFit
    
    ' Set the type in B5
    oWorksheet.Range("B5").Value = sType
End Sub
```

```javascript
// This example shows how to get the parent object for the specified collection.
/* Этот пример показывает, как получить родительский объект для указанной коллекции. */
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1:D1"); // Get the range B1:D1
oRange.SetValue("1"); // Set value to "1"
oRange.Select(); // Select the range
var oAreas = oRange.GetAreas(); // Get the areas of the range
var oParent = oAreas.GetParent(); // Get the parent of the areas
var sType = oParent.GetClassType(); // Get the class type of the parent
oRange = oWorksheet.GetRange('A4'); // Get the range A4
oRange.SetValue("The areas parent: "); // Set value in A4
oRange.AutoFit(false, true); // Autofit columns
oWorksheet.GetRange('B4').Paste(oParent); // Paste the parent object reference in B4
oRange = oWorksheet.GetRange('A5'); // Get the range A5
oRange.SetValue("The type of the areas parent: "); // Set value in A5
oRange.AutoFit(false, true); // Autofit columns
oWorksheet.GetRange('B5').SetValue(sType); // Set the type in B5
```