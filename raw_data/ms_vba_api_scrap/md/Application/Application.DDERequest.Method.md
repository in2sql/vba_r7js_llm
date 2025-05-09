# Application.DDERequest method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="System") 
returnList = Application.DDERequest(channelNumber, "Topics") 
For i = LBound(returnList) To UBound(returnList) 
 Worksheets("Sheet1").Cells(i, 1).Formula = returnList(i) 
Next i 
Application.DDETerminate channelNumber
```

## Parameters
- **Channel**: Required
- **Item**: Required

## Return Value
Variant

## Example
No VBA example available.
