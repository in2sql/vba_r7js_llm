Attribute VB_Name = "Module1"
Option Explicit
'This project is used to demonstrate variable SCOPE. The rules of scope determine
'how widely a variable can be "seen" within a program. Variables should always be
'declared with the narrowest possible scope suitable for the problem at hand. This
'is the most resource-conservative and most robust practice.
'
'The narrowest scope is "Local" meaning you use a regular Dim statement within
'the sub. A local variable is only good inside the sub where it is declared.
'
'The next level of scope is "Module-level". A module level variable can be shared
'between any Subs within the same code module, and will retain its value between subs.
'They are declared using a dim statement at the top of the module,
'or with the keyword "Private" (meaning private to that module) like so...

'Dim bookCount as integer           'Module level declaration
'Private bookCount as integer       'This means the same thing

'The broadest level of scope is "Global" meaning the variable can be seen by all Subs
'in ALL MODULES. You can declare a global variable using the keyword "Global" or
'the keyword "Public" (meaning Public to all modules) like so...

Global bookCount As Integer         'Global declaration
'Public bookCount as integer        'This would also work

Public Sub CreateBooks()
Dim x As Integer
'Use two new variables to allow the user to choose how many workbooks to create
'We use a string to receive the result from the Inputbox
'Then attempt to convert it to a number if possible for the numBooks integer variable
Dim numString As String
Dim numBooks As Integer


numString = InputBox("How many new workbooks?")

'Check to see if we got a number, if not, or empty in the case of Cancel, bail out..
If Not IsNumeric(numString) Then Exit Sub

numBooks = CInt(numString)      'covert to integer

For x = 1 To numBooks
    bookCount = bookCount + 1   'keep track of how many workbooks we created.
    Workbooks.Add
Next

'This will tile all the new workbook windows so we can see all of
'them on the screen.

Application.Windows.Arrange xlArrangeStyleTiled
End Sub

Public Sub CloseAllExcept()
'Close all the open workbooks except this one containing the code
Dim wb As Workbook

For Each wb In Workbooks
    If Not wb Is ThisWorkbook Then wb.Close False
Next

'this will zoom the remaining window to fill the screen

Application.Windows.Arrange xlArrangeStyleVertical

End Sub
'-------------------------------------------------------------------------
'Note that Subs can also be declared as Public or Private. This has essentially
'the same meaning. Private Subs can only be called by other Subs in the same
'module. Public Subs can be called by any Sub in any module. Subs are by default
'Public, which is fine for our purposes. Although variable scope is important in
'VBA, Sub scope is less important for casual programmers.
'--------------------------------------------------------------------------




