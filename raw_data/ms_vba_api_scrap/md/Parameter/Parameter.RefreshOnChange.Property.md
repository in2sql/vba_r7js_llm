# Parameter RefreshOnChange Property

## Business Description
True if the specified query table is refreshed whenever you change the parameter value of a parameter query. Read/write Boolean.

## Behavior
Trueif the specified query table is refreshed whenever you change the parameter value of a parameter query. Read/writeBoolean.

## Example Usage
```vba
Set objQT = Worksheets("Sheet1").QueryTables(1) 
objQT.CommandText = "Select * From Customers Where (ContactTitle=?)" 
Set objParam1 = objQT.Parameters _ 
 .Add("Contact Title", xlParamTypeVarChar) 
objParam1.RefreshOnChange= True 
objParam1.SetParam xlRange, Range("D4")
```