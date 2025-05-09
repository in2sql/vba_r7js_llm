# DefaultWebOptions TargetBrowser Property

## Business Description
Returns or sets an MsoTargetBrowserhttp://msdn.microsoft.com/library/6ce561d2-c327-b433-3c91-df1036e87a75(Office.15).aspx constant indicating the browser version. Read/write.

## Behavior
Returns or sets anMsoTargetBrowserconstant indicating the browser version. Read/write.

## Example Usage
```vba
Sub CheckWebOptions() 
 
    Dim wkbOne As Workbook 
 
    Set wkbOne = Application.Workbooks(1) 
 
    ' Determine if IE5 is the target browser. 
    If wkbOne.WebOptions.TargetBrowser= msoTargetBrowserIE5 Then 
        MsgBox "The target browser is IE5 or later." 
    Else 
        MsgBox "The target browser is not IE5 or later." 
    End If 
 
End Sub
```