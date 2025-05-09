# Opening a Workbook

## Business Description
When you open a workbook using the Open method, it becomes a member of the Workbooks collection. The following procedure opens a workbook named MyBook.xls located in the folder named MyFolder on drive C.

## Behavior
When you open a workbook using theOpenmethod, it becomes a member of theWorkbookscollection. The following procedure opens a workbook named MyBook.xls located in the folder named MyFolder on drive C.

## Example Usage
```vba
Sub OpenUp() 
 Workbooks.Open("C:\MyFolder\MyBook.xls") 
End Sub
```