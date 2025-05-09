# Application.FeatureInstall property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Dim WordApp As New Word.Application, Reply As Integer 
Application.ActivateMicrosoftApp xlMicrosoftWord With WordApp 
    If .FeatureInstall = msoFeatureInstallNone Then 
        Reply = MsgBox("Uninstalled features for this " _ 
            & "application " & vbCrLf _ 
            & "may cause a run-time error when called." & vbCrLf _ 
            & vbCrLf _ 
            & "Would you like to change this setting" & vbCrLf _ 
            & "to automatically install missing features?" _ 
            , 52, "Feature Install Setting") 
        If Reply = 6 Then 
            .FeatureInstall = msoFeatureInstallOnDemand 
        End If 
    End If 
End With
```

## Remarks
MsoFeatureInstall can be one of these constants:

## Example
```vba
Dim WordApp As New Word.Application, Reply As Integer 
Application.ActivateMicrosoftApp xlMicrosoftWord With WordApp 
    If .FeatureInstall = msoFeatureInstallNone Then 
        Reply = MsgBox("Uninstalled features for this " _ 
            & "application " & vbCrLf _ 
            & "may cause a run-time error when called." & vbCrLf _ 
            & vbCrLf _ 
            & "Would you like to change this setting" & vbCrLf _ 
            & "to automatically install missing features?" _ 
            , 52, "Feature Install Setting") 
        If Reply = 6 Then 
            .FeatureInstall = msoFeatureInstallOnDemand 
        End If 
    End If 
End With
```

