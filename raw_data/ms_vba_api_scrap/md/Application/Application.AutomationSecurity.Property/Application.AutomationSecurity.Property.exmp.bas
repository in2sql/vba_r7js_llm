Sub Security() 
    Dim secAutomation As MsoAutomationSecurity 
 
    secAutomation = Application.AutomationSecurity 
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable 
    Application.FileDialog(msoFileDialogOpen).Show 
 
    Application.AutomationSecurity = secAutomation 
 
End Sub