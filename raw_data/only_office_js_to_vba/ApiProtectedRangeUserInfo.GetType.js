### Description / Описание

**English:**  
This code adds a protected range to the active worksheet, assigns a user with view permissions, retrieves the user type, and sets a cell value indicating the user type.

**Russian:**  
Этот код добавляет защищенный диапазон на активный лист, назначает пользователю права на просмотр, получает тип пользователя и устанавливает значение ячейки, указывающее тип пользователя.

```vba
' VBA Code Equivalent

Sub ManageProtectedRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Protect the worksheet to lock cells
    ws.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Lock the range A1:B1
    ws.Range("A1:B1").Locked = True
    
    ' Unlock other cells if necessary
    ws.Range("A1:B1").Locked = True
    ws.Range("A3").Locked = False
    
    ' Set value in cell A3
    ws.Range("A3").Value = "Type: ViewOnly" ' Placeholder since VBA does not support user types
End Sub
```

```javascript
// JavaScript Code Equivalent
// This example changes the user protected range.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
var userInfo = protectedRange.GetUser("userId");
var userType = userInfo.GetType();
oWorksheet.GetRange("A3").SetValue("Type: " + userType); 
```