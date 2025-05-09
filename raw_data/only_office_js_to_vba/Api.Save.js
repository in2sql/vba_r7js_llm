**Description / Описание**

This code saves changes to the active worksheet by setting a value in cell A1 and then saving the document.  
Этот код сохраняет изменения в активном листе, устанавливая значение в ячейку A1 и затем сохраняя документ.

```vba
' This VBA code saves changes to the active worksheet by setting a value in cell A1 and saving the workbook.
Sub SaveChanges()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ws.Range("A1").Value = "This sample text is saved to the worksheet." ' Set value in A1
    
    ThisWorkbook.Save ' Save the workbook
End Sub
```

```javascript
// This JavaScript code saves changes to the active worksheet by setting a value in cell A1 and saving the document.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
oWorksheet.GetRange("A1").SetValue("This sample text is saved to the worksheet."); // Set value in A1
Api.Save(); // Save the document
```