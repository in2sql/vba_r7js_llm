# DefaultWebOptions SaveNewWebPagesAsWebArchives Property

## Business Description
True if new Web pages can be saved as Web archives. Read/write Boolean.

## Behavior
Trueif new Web pages can be saved as Web archives. Read/writeBoolean.

## Example Usage
```vba
Sub DetermineSettings() 
 
 ' Determine settings and notify user. 
 If Application.DefaultWebOptions.SaveNewWebPagesAsWebArchives= True Then 
 MsgBox "New Web pages will be saved as Web archives." 
 Else 
 MsgBox "New Web pages will not be saved as Web archives." 
 End If 
 
End Sub
```