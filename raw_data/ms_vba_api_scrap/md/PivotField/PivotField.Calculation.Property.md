# PivotField Calculation Property

## Business Description
Returns or sets a XlPivotFieldCalculation value that represents the type of calculation performed by the specified field. This property is valid only for data fields.

## Behavior
Returns or sets aXlPivotFieldCalculationvalue that represents the type of calculation performed by the specified field. This property is valid only for data fields.

## Example Usage
```vba
With Worksheets("Sheet1").Range("A3").PivotField 
    .Calculation = xlDifferenceFrom 
    .BaseField = "ORDER_DATE" 
    .BaseItem = "5/16/89" 
End With
```