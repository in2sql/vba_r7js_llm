# PivotField DataType Property

## Business Description
Returns a XlPivotFieldDataType value that represents the type of data in the PivotTable field.

## Behavior
Returns aXlPivotFieldDataTypevalue that represents the type of data in the PivotTable field.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Select Case pvtTable.PivotFields("ORDER_DATE").DataTypeCase Is = xlText 
 MsgBox "The field contains text data" 
 Case Is = xlNumber 
 MsgBox "The field contains numeric data" 
 Case Is = xlDate 
 MsgBox "The field contains date data" 
End Select
```