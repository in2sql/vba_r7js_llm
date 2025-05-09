# Application.PrintCommunication property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.PrintCommunication = False 
 With ActiveSheet.PageSetup 
 .PrintTitleRows = "" 
 .PrintTitleColumns = "" 
 End With 
Application.PrintCommunication = True
```

## Return Value
True if communication with the printer is turned on; otherwise, False.

## Remarks
Set the PrintCommunication property to False to speed up the execution of code that sets PageSetup properties.

## Example
No VBA example available.
