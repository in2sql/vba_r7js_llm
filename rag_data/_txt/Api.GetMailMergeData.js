# Description / Описание

This code retrieves mail merge data and populates the worksheet accordingly.
Этот код получает данные слияния почты и заполняет рабочий лист соответственно.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the width of the first column to 20
oWorksheet.SetColumnWidth(0, 20);

// Set header values
oWorksheet.GetRange("A1").SetValue("Email address");
oWorksheet.GetRange("B1").SetValue("Greeting");
oWorksheet.GetRange("C1").SetValue("First name");
oWorksheet.GetRange("D1").SetValue("Last name");

// Set data for first row
oWorksheet.GetRange("A2").SetValue("user1@example.com");
oWorksheet.GetRange("B2").SetValue("Dear");
oWorksheet.GetRange("C2").SetValue("John");
oWorksheet.GetRange("D2").SetValue("Smith");

// Set data for second row
oWorksheet.GetRange("A3").SetValue("user2@example.com");
oWorksheet.GetRange("B3").SetValue("Hello");
oWorksheet.GetRange("C3").SetValue("Kate");
oWorksheet.GetRange("D3").SetValue("Cage");

// Get mail merge data
var aMailMergeData = Api.GetMailMergeData(0);

// Set mail merge data in cell A5
oWorksheet.GetRange("A5").SetValue("Mail merge data: " + aMailMergeData);
```

```vba
' Excel VBA equivalent code

Sub GetMailMergeData()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set the width of the first column to 20
    oWorksheet.Columns(1).ColumnWidth = 20
    
    ' Set header values
    oWorksheet.Range("A1").Value = "Email address"
    oWorksheet.Range("B1").Value = "Greeting"
    oWorksheet.Range("C1").Value = "First name"
    oWorksheet.Range("D1").Value = "Last name"
    
    ' Set data for first row
    oWorksheet.Range("A2").Value = "user1@example.com"
    oWorksheet.Range("B2").Value = "Dear"
    oWorksheet.Range("C2").Value = "John"
    oWorksheet.Range("D2").Value = "Smith"
    
    ' Set data for second row
    oWorksheet.Range("A3").Value = "user2@example.com"
    oWorksheet.Range("B3").Value = "Hello"
    oWorksheet.Range("C3").Value = "Kate"
    oWorksheet.Range("D3").Value = "Cage"
    
    ' Get mail merge data (assuming a function to retrieve it)
    Dim aMailMergeData As String
    aMailMergeData = GetMailMergeDataFunction(0) ' Placeholder for actual mail merge data retrieval
    
    ' Set mail merge data in cell A5
    oWorksheet.Range("A5").Value = "Mail merge data: " & aMailMergeData
End Sub

' Placeholder function for GetMailMergeData
Function GetMailMergeDataFunction(index As Integer) As String
    ' Implement mail merge data retrieval logic here
    GetMailMergeDataFunction = "Sample Mail Merge Data"
End Function
```