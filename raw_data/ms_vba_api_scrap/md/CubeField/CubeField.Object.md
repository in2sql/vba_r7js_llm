# CubeField Object

## Business Description
Represents a hierarchy or measure field from an OLAP cube. In a PivotTable report, the CubeField object is a member of the CubeFields collection.

## Behavior
Represents a hierarchy or measure field from an OLAP

			cube. In a PivotTable report, theCubeFieldobject is a member of theCubeFieldscollection.

## Example Usage
```vba
strAlphaName = _ 
 ActiveSheet.PivotTables(1).CubeFields(2).Name
```