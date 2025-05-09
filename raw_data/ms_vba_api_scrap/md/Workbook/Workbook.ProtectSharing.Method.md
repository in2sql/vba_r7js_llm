# Workbook ProtectSharing Method

## Business Description
Saves the workbook and protects it for sharing.

## Behavior
Saves the workbook and protects it for sharing.

## Example Usage
```vba
Sub ProtectWorkbook() 
 
    Dim wbAWB As Workbook 
    Dim strPwd As String 
    Dim strSharePwd As String 
 
    Set wbAWB = Application.ActiveWorkbook 
 
    strPwd = InputBox("Enter password for the file") 
    strSharePwd = InputBox("Enter password for sharing") 
 
    wbAWB.ProtectSharingPassword:=strPwd, _ 
        SharingPassword:=strSharePwd 
 
End Sub
```