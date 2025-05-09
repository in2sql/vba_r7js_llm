# Workbook UpdateLinks Property

## Business Description
Returns or sets an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links. Read/write.

## Behavior
Returns or sets anXlUpdateLinkconstant indicating a workbook's setting for updating embedded OLE links. Read/write.

## Example Usage
```vba
Sub UseUpdateLinks() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.Workbooks(1) 
 
 Select Case wkbOne.UpdateLinksCase xlUpdateLinksAlways 
 MsgBox "Links will always be updated " & _ 
 "for the specified workbook." 
 Case xlUpdateLinksNever 
 MsgBox "Links will never be updated " & _ 
 "for the specified workbook." 
 Case xlUpdateLinksUserSetting 
 MsgBox "Links will update according " & _ 
 "to user settting for the specified workbook." 
 End Select 
 
End Sub
```