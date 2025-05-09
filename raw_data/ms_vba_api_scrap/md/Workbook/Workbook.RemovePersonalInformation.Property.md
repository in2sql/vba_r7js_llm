# Workbook RemovePersonalInformation Property

## Business Description
True if personal information can be removed from the specified workbook. The default value is False. Read/write Boolean.

## Behavior
Trueif personal information can be removed from the specified workbook. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
Sub UsePersonalInformation() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.ActiveWorkbook 
 
 ' Determine settings and notify user. 
 If wkbOne.RemovePersonalInformation= True Then 
 MsgBox "Personal information can be removed." 
 Else 
 MsgBox "Personal information cannot be removed." 
 End If 
 
End Sub
```