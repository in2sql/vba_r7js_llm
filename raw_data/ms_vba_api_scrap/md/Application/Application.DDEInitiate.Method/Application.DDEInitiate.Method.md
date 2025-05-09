# Application.DDEInitiate method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber
```

## Parameters
- **App**: Required
- **Topic**: Required

## Return Value
Long

## Remarks
If successful, the DDEInitiate method returns the number of the open channel. All subsequent DDE functions use this number to specify the channel.

## Example
No VBA example available.
