# PivotTable CreateCubeFile Method

## Business Description
Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.

## Behavior
Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.

## Example Usage
```vba
Sub UseCreateCubeFile() 
 
 ActiveSheet.PivotTables(1).CreateCubeFile_ 
 File:="C:\CustomCubeFile", Properties:=False 
 
End Sub
```