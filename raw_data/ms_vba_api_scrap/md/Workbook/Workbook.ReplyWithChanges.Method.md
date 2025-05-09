# Workbook ReplyWithChanges Method

## Business Description
Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.

## Behavior
Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.

## Example Usage
```vba
Sub ReplyMsg() 
 
 ActiveWorkbook.ReplyWithChangesShowMessage:=False 
 
End Sub
```