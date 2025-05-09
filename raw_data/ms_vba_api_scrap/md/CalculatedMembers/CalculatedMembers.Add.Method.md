# CalculatedMembers Add Method

## Business Description
Adds a calculated field or calculated item to a PivotTable. Returns a CalculatedMember object.

## Behavior
Adds a calculated field or calculated item to a PivotTable. Returns aCalculatedMemberobject.

## Example Usage
```vba
Sub UseAddSet() 
 
 Dim pvtOne As PivotTable 
 Dim strAdd As String 
 Dim strFormula As String 
 Dim cbfOne As CubeField 
 
 Set pvtOne = ActiveSheet.PivotTables(1) 
 
 strAdd = "[MySet]" 
 strFormula = "'{[Product].[All Products].[Food].children}'" 
 
 ' Establish connection with data source if necessary. 
 If Not pvtOne.PivotCache.IsConnected Then pvtOne.PivotCache.MakeConnection 
 
 ' Add a calculated member titled "[MySet]" 
 pvtOne.CalculatedMembers.AddName:=strAdd, _ 
 Formula:=strFormula, Type:=xlCalculatedSet 
 
 ' Add a set to the CubeField object. 
 Set cbfOne = pvtOne.CubeFields.AddSet(Name:="[MySet]", _ 
 Caption:="My Set") 
 
End Sub
```