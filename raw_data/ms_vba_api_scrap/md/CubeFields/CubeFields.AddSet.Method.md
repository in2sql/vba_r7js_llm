# CubeFields AddSet Method

## Business Description
Adds a new CubeField object to the CubeFields collection. The CubeField object corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.

## Behavior
Adds a newCubeFieldobject to theCubeFieldscollection. TheCubeFieldobject corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.

## Example Usage
```vba
Sub UseAddSet() 
 
 Dim pvtOne As PivotTable 
 Dim strAdd As String 
 Dim strFormula As String 
 Dim cbfOne As CubeField 
 
 Set pvtOne = Sheet1.PivotTables(1) 
 
 strAdd = "[MySet]" 
 strFormula = "'{[Product].[All Products].[Food].children}'" 
 
 ' Establish connection with data source if necessary. 
 If Not pvtOne.PivotCache.IsConnected Then pvtOne.PivotCache.MakeConnection 
 
 ' Add a calculated member titled "[MySet]" 
 pvtOne.CalculatedMembers.Add Name:=strAdd, _ 
 Formula:=strFormula, Type:=xlCalculatedSet 
 
 ' Add a set to the CubeField object. 
 Set cbfOne = pvtOne.CubeFields.AddSet(Name:="[MySet]", _ 
 Caption:="My Set") 
 
End Sub
```