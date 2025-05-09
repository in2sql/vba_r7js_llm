# Application.AddCustomList method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
On Error Resume Next  ' if the list already exists, don'thing
Application.AddCustomList Array("cogs", "sprockets", _ 
 "widgets", "gizmos")
On Error Goto 0       ' resume regular error handling
```

## Parameters
- **ListArray**: Required
- **ByRow**: Optional

## Remarks
If the list that you are trying to add already exists, this method throws a run-time error 1004. Catch the error with an On Error statement.

## Example
No VBA example available.
