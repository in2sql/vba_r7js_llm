# Application.CommandBars property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
For Each bar In Application.CommandBars 
    If Not bar.BuiltIn And Not bar.Visible Then bar.Delete 
Next
```

## Remarks
Used with the Application object, this property returns the set of built-in and custom command bars available to the application.

## Example
No VBA example available.
