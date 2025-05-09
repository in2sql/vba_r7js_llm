# AutoFilter.Range property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
rAddress = Worksheets("Crew").AutoFilter.Range.Address
```

## Example
```vba
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).Range 
ActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```

