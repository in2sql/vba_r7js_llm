# Application.DDEPoke method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\SALES.DOC") 
Set rangeToPoke = Worksheets("Sheet1").Range("A1") 
Application.DDEPoke channelNumber, "\StartOfDoc", rangeToPoke 
Application.DDETerminate channelNumber
```

## Parameters
- **Channel**: Required
- **Item**: Required
- **Data**: Required

## Remarks
An error occurs if the method call doesn't succeed.

## Example
```vba
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\SALES.DOC") 
Set rangeToPoke = Worksheets("Sheet1").Range("A1") 
Application.DDEPoke channelNumber, "\StartOfDoc", rangeToPoke 
Application.DDETerminate channelNumber
```

