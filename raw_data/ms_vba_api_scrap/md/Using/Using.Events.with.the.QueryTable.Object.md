# Using Events with the QueryTable Object

## Business Description
Before you can use events with the QueryTable object, you must first create a class module and declare a QueryTable object with events. For example, assume that you have created a class module and named it ClsModQT. This module contains the following code:

## Behavior
Before you can use events with theQueryTableobject, you must first create a  class module and declare aQueryTableobject with events. For example, assume that you have created a  class module and named itClsModQT. This module contains the following code:

## Example Usage
```vba
Sub InitQueryEvent(QT as Object) 
 Set qtQueryTable = QT 
End Sub
```