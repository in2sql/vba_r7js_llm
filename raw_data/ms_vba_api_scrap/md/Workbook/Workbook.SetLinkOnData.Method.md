# Workbook SetLinkOnData Method

## Business Description
Sets the name of a procedure that runs whenever a DDE link is updated.

## Behavior
Sets the name of a procedure that runs whenever a DDE link is updated.

## Example Usage
```vba
ActiveWorkbook.SetLinkOnData_ 
 "WinWord|'C:\MSGFILE.DOC'!DDE_LINK1", _ 
 "my_Link_Update_Macro"
```