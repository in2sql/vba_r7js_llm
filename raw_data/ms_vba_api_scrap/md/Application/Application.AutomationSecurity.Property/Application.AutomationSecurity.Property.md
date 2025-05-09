# Application.AutomationSecurity property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub Security() 
    Dim secAutomation As MsoAutomationSecurity 
 
    secAutomation = Application.AutomationSecurity 
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable 
    Application.FileDialog(msoFileDialogOpen).Show 
 
    Application.AutomationSecurity = secAutomation 
 
End Sub
```

## Remarks
This property is automatically set to msoAutomationSecurityLow when the application is started. Therefore, to avoid breaking solutions that rely on the default setting, you should be careful to reset this property to msoAutomationSecurityLow after programmatically opening a file. Also, this property should be set immediately before and after opening a file programmatically to avoid malicious subversion.

## Example
```vba
Sub Security() 
    Dim secAutomation As MsoAutomationSecurity 
 
    secAutomation = Application.AutomationSecurity 
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable 
    Application.FileDialog(msoFileDialogOpen).Show 
 
    Application.AutomationSecurity = secAutomation 
 
End Sub
```

