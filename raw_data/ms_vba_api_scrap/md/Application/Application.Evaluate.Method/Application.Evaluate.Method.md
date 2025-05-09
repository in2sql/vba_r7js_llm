# Application.Evaluate method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
[a1].Value = 25 
Evaluate("A1").Value = 25 
 
trigVariable = [SIN(45)] 
trigVariable = Evaluate("SIN(45)") 
 
Set firstCellInSheet = Workbooks("BOOK1.XLS").Sheets(4).[A1] 
Set firstCellInSheet = _ 
    Workbooks("BOOK1.XLS").Sheets(4).Evaluate("A1")
```

## Parameters
- **Name**: Required

## Return Value
Variant

## Remarks
The following types of names in Microsoft Excel can be used with this method:

## Example
```vba
[a1].Value = 25 
Evaluate("A1").Value = 25 
 
trigVariable = [SIN(45)] 
trigVariable = Evaluate("SIN(45)") 
 
Set firstCellInSheet = Workbooks("BOOK1.XLS").Sheets(4).[A1] 
Set firstCellInSheet = _ 
    Workbooks("BOOK1.XLS").Sheets(4).Evaluate("A1")
```

