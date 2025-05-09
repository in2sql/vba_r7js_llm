# Workbook SendForReview Method

## Business Description
Sends a workbook in an e-mail message for review to the specified recipients.

## Behavior
Sends a workbook in an e-mail message for review to the specified recipients.

## Example Usage
```vba
Sub WebReview() 
 
 ActiveWorkbook.SendForReview_ 
 Recipients:="someone@example.com; amy jones; lewjudy", _ 
 Subject:="Please review this document.", _ 
 ShowMessage:=False, _ 
 IncludeAttachment:=True 
 
End Sub
```