```vba
' Change the title of a protected range in the worksheet
' Изменение названия защищенного диапазона в листе

Sub ChangeProtectedRangeTitle()
    Dim ws As Worksheet
    Dim nm As Name
    
    ' Set worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Unlock all cells
    ws.Cells.Locked = False
    
    ' Lock specific range
    ws.Range("A1:B1").Locked = True
    
    ' Protect the sheet
    ws.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Add named range
    On Error Resume Next
    Set nm = ThisWorkbook.Names("protectedRange")
    If nm Is Nothing Then
        ThisWorkbook.Names.Add Name:="protectedRange", RefersTo:=ws.Range("A1:B1")
    End If
    On Error GoTo 0
    
    ' Rename the named range
    nm.Name = "protectedRangeNew"
End Sub
```

```javascript
// Change the title of a protected range in the worksheet
// Изменение названия защищенного диапазона в листе

function changeProtectedRangeTitle() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Add a protected range named "protectedRange" for the range A1:B1 in Sheet1
    oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");
    
    // Get the protected range by name
    var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
    
    // Set the new title for the protected range
    protectedRange.SetTitle("protectedRangeNew");
}
```