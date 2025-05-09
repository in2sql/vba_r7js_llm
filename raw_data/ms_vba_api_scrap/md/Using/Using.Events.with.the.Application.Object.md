# Using Events with the Application Object

## Business Description
Before you can use events with the Application object, you must create a class module and declare an object of type Application with events. For example, assume that a new class module is created and called EventClassModule.

## Behavior
Before you can use events with theApplicationobject, you must create a class module and declare an object of typeApplicationwith events. For example, assume that a new class module is created and called EventClassModule. The new class module contains the following code:

## Example Usage
```vba
Dim X As New EventClassModule 
 
Sub InitializeApp() 
 Set X.App = Application 
End Sub
```