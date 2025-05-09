# PivotCache ADOConnection Property

## Business Description
Returns an ADO Connection object if the PivotTable cache is connected to an OLE DB data source.

## Behavior
Returns anADO Connectionobject if the PivotTable cache is connected to an OLE DB data source. TheADOConnectionproperty exposes the Microsoft Excel connection to the data provider, allowing the user to write code within the context of the same session that Excel is using with ADO (relational source) or ADO MD (OLAP source). Read-only.

## Example Usage
```vba
Sub UseADOConnection() 
 
 Dim ptOne As PivotTable 
 Dim cmdOne As New ADODB.Command 
 Dim cfOne As CubeField 
 
 Set ptOne = Sheet1.PivotTables(1) 
 ptOne.PivotCache.MaintainConnection = True 
 Set cmdOne.ActiveConnection = ptOne.PivotCache.ADOConnectionptOne.PivotCache.MakeConnection 
 
 ' Create a set. 
 cmdOne.CommandText = "Create Set [Warehouse].[My Set] as '{[Product].[All Products].Children}'" 
 cmdOne.CommandType = adCmdUnknown 
 cmdOne.Execute 
 
 ' Add a set to the CubeField. 
 Set cfOne = ptOne.CubeFields.AddSet("My Set", "My Set") 
 
End Sub
```