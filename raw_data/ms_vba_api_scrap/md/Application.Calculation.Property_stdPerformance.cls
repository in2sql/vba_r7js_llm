VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'This class has been designed to meet your performance testing and optimisation needs. stdPerformance uses the `Sentry Object` design pattern
'which allows for cleaner more maintainable code.
'Functions implemented on the class
'CONSTRUCTORS
'    [X] Create   - With Cache
'    [X] init       #PROTECTED
'    [X] Measure  - Create a performance measuring Sentry Object
'    [X] Optimise - Create a object which toggles runtime options for optimisation. Currently sets: `ScreenUpdating`, `EnableEvents` and `XLCalculation`.
'                   This is intended to be application agnostic.
'
'STATIC PROPERTIES
'    [X] Get MeasureKeys - Get an array of procs/blocks which have been measured
'    [x] Get Measurement(sProcOrBlock) - Get the average time it took to execute a block.
'    [ ] Get MeasuresStr
'    [ ] Get MeasuresHtml
'
'STATIC methods
'    [x] MeasuresClear() - Clear the performance stack.
'
'OUT-OF-SCOPE
'    * Anything performance related which is specific, should realistically be honed to a specific class for that thing.
'
'EXAMPLES
'# 1 - Usage of Optimser
'
'   'Disable numerous options for performance
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   With stdPerformance.Optimiser()
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4135
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   
'   'Disable everything BUT Calculation
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   With stdPerformance.Optimiser(Calculation:=xlCalculation.xlCalculationAutomatic)
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'
'# 2 - measuring performance:
'
'   With stdPerformance.measure("#1 Select then set")
'     For i = 1 to C_MAX
'       cells(1,1).select
'       selection.value = "hello"
'     Next
'   End With
'   
'   With stdPerformance.measure("#2 Set directly")
'     For i = 1 to C_MAX
'       cells(1,1).value = "hello"
'     next
'   End With
'
'Declares for performance counters
#If Mac Then
   #If MAC_OFFICE_VERSION >= 15 Then
      Private Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
   #Else
      Private Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
   #Else
      Private Declare Function GetTickCount Lib "kernel32" () As Long
   #End If
#End If

'Enum for sentry object type
Public Enum EPerfObjType
  iMeasure=1
  iOptimiser=2
End Enum


'For Performance Optimisation sentries
Private Type TOptimiser
  EnableEvents as boolean
  ScreenUpdating as boolean
  Calculation as long
End Type

'For Performance Measurement
Private Type TMeasure
  StartTime as long
  KeyName as string
  Divider as double
End Type

'For Measurement storage
Private Type FakeDictItem
  key as string
  val as variant
End Type

Private Type TSingleton
  FakeDict() as FakeDictItem   'was Private FakeDict() as FakeDictItem
End Type

Private Type TThis
  Singleton as TSingleton

  ObjectType as EPerfObjType
  Measure as TMeasure
  Optimiser as TOptimiser
End Type
Private This as TThis




'Create
'@constructor
'@param objType - Type of performance object to create. iMeasure - used for measuring performance, iOptimiser - used for optimising performance 
'@param params as Variant<Array<Variant>> - Additional params supplied as array.
'@returns stdPerformance<EPerfObjType> - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@remark 2023-11-23 This function is not intended to be called directly and thus has been made private. Use the Measure or Optimise methods instead. If absolutely required use `protInit`.
Private Function Create(ByVal objType as EPerfObjType, ByVal params as Variant) as stdPerformance
  Set Create = new stdPerformance
  Call Create.protInit(objType, params)
End Function

'Initialises the class
'@protected
'@param objType - Type of performance object to initialise.
'
'* iMeasure - used for measuring performance
'* iOptimiser - used for optimising performance 
'@param params as Variant<Array<Variant>> - Additional params supplied as array.
Public Sub protInit(ByVal objType as EPerfObjType, ByVal params as variant)
  this.ObjectType = objType
  select case objType
    case iMeasure
      this.Measure.KeyName = params(0)
      this.Measure.Divider = params(1)
      this.Measure.StartTime = GetTickCount()
    case iOptimiser
      'Store vals
      this.Optimiser.ScreenUpdating = Application.ScreenUpdating
      this.Optimiser.EnableEvents = Application.EnableEvents
      
      'Set vals
      Application.ScreenUpdating = params(0)
      Application.EnableEvents = params(1)
      
      'Different options for different applications
      select case Application.Name 
        case "Microsoft Excel"
          this.Optimiser.Calculation = Application.Calculation
          Application.Calculation = params(2)
      end select
  end select
End Sub

'Measure
'@constructor
'@param sProc   - Name of method or block to measure
'@param nCount  - Number of times the block will be ran (used to calculate average time), default = 1.
'@returns stdPerformance<iMeasure> - Object which upon termination, adds measurement of block to global cache
'@example
'```vb
'With stdPerformance.Measure("Some test expression")
'  Dim v
'  For i = 1 to 1000
'    v = expressionToTest()
'  next
'End With
'```
'@example
'```vb
'Const C_MAX as Long = 1000
'With stdPerformance.Measure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'```
Public Function CreateMeasure(ByVal sProc as string, Optional ByVal nCount as double=1) as stdPerformance
  set CreateMeasure = Create(iMeasure, Array(sProc,nCount))
End Function

'Optimise
'@constructor
'@param ScreenUpdating  - ScreenUpdating set value
'@param EnableEvents    - EnableEvents set value
'@param Calculation     - Calculation set value
'@returns stdPerformance<iOptimiser> - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@note Calculation is defined as long instead of xlCalculation so the function continues to work without compile error in Word, Powerpoint etc.
'@example
'```vb
'With stdPerformance.Optimise
'  'some heavy code here
'End With
'```
Public Function CreateOptimiser(Optional ByVal ScreenUpdating as boolean = false, Optional ByVal EnableEvents as boolean = false, Optional ByVal Calculation as long = -4135) as stdPerformance
  set CreateOptimiser = Create(iOptimiser, Array(ScreenUpdating,EnableEvents,Calculation))
End Function


'Obtain a measurement
'@param sKey - Name of measurement to get
'@returns - Average measurement time
Public Property Get Measurement(ByVal sKey As String) As Double
  If Me Is stdPerformance Then
    Dim v: v = getDictItem(sKey)
    If TypeName(v) = "Variant()" Then
        Measurement = getDictItem(sKey)(0)
    Else
        Measurement = Empty
    End If
  End If
End Property

'AddMeasurement
'@protected
'@param sKey   - Name of measurement to add to global cache
'@param time   - time to add to global cache
'@param nCount - number of operations (divisor)
'@remark If a time is added that was previously also added then the average of the times is calculated.
Public Sub AddMeasurement(ByVal sKey as string, ByVal time as Double, ByVal nCount as Double)
  if Me is stdPerformance then
    Debug.Print sKey & ": " & time & " ms" & iif(nCount>1," (" & (1000*time/nCount) & chr(181) & "s per operation)","")
    Dim ind as long: ind = getDictIndex(sKey)
    if ind = -1 then
      Call setDictItem(sKey, Array(time,1))
    else
      Dim vItem: vItem = getDictItem(sKey)
      Dim average as long: average = vItem(0)
      Dim count as long: count = vItem(1) + 1
      average = average + (time - average)/count
      Call setDictItem(sKey, Array(average,count))
    end if
  end if
End Sub

'MeasuresClear - Clears all procedures/blocks and times that have been measured
Public Sub MeasuresClear()
  ReDim this.Singleton.FakeDict(0 to 0)
End Sub

'MeasuresKeys
'@returns Array<String> - Array containing the procedures or blocks that have been measured.
Public Property Get MeasuresKeys() as string()
  if Me is stdPerformance then
    if ubound(this.Singleton.FakeDict) = 0 then
      MeasuresKeys = Split("")
    else
      'Define return array
      Dim sOut() as string
      Redim Preserve sOut(0 to ubound(this.Singleton.FakeDict)-1)

      'Fill keys array 
      Dim i as long
      For i = 0 to ubound(this.Singleton.FakeDict)-1
        sOut(i) = this.Singleton.FakeDict(i).key
      next

      'return data
      MeasuresKeys = sOut
    end if
  end if 
End Property

'Used by static class only
'@constructor
Private Sub Class_Initialize()
  if me is stdPerformance then
    Redim this.Singleton.FakeDict(0 to 0)
  end if
End Sub

'Used by instance objects only
'@destructor
Private Sub Class_Terminate()
  if not me is stdPerformance then
    select case this.ObjectType
      case iMeasure
        Dim pEndTime as long: pEndTime = GetTickCount()
        Call stdPerformance.AddMeasurement(this.Measure.KeyName, Abs(pEndTime - this.Measure.StartTime),this.Measure.Divider)
      case iOptimiser
        'Store vals
        Application.ScreenUpdating = this.Optimiser.ScreenUpdating
        Application.EnableEvents = this.Optimiser.EnableEvents
        
        'Different options for different applications
        select case Application.Name 
          case "Microsoft Excel"
            Application.Calculation = this.Optimiser.Calculation
        end select
    end select
  end if
End Sub




'FakeDict Helpers
'==========================================================================================================================================
'NOTE: These functions are completely unoptimised and are largely in use for the purpose of making this class multi-platform friendly.
'These will be unlikely to be optimised given that this is largely a debugging library.

'getDictIndex
'Returns the index where a particular key is stored
'@param key - Key to find in dictionary
'@returns   - Index of key in dictionary
Private Function getDictIndex(ByVal key as string) as Long
  On Error GoTo ErrorOccurred
    Dim i as long
    For i = 0 to ubound(this.Singleton.FakeDict)
      if this.Singleton.FakeDict(i).key = key then
        getDictIndex = i
        Exit Function
      end if
    next
  On Error Goto 0
ErrorOccurred:
  getDictIndex = -1
End Function

'setDictItem
'Set an item within a dictionary to a particular value
'@param key - Key to find in dictionary
'@param v   - Value to set dictionary too
'@param ind - Index of item. If not given `getDictIndex()` is used
Private Sub setDictItem(ByVal key as string, ByVal v as variant, Optional ByVal ind as long = -1)
  'get index of item in fake dict
  if ind = -1 then ind = getDictIndex(key)
  
  'If item not exist, add it
  if ind = -1 then
    ind = getUB(this.Singleton.FakeDict)
    this.Singleton.FakeDict(ind).key = key
    Redim Preserve this.Singleton.FakeDict(0 to ind+1)
  end if

  'Assign value to index
  if isObject(v) then
    set this.Singleton.FakeDict(ind).val = v
  else
    this.Singleton.FakeDict(ind).val = v
  end if
End Sub

'getUB
'gets the upper bound of an array, if the array is uninitialised return -1
'@param items as Array<FakeDictItem> - Array of dict items.
'@returns - Upper bound of array OR -1 if not initialised
Private Function getUB(ByRef items() As FakeDictItem) As Long
    On Error GoTo ErrorOccurred
        getUB = UBound(items)
        Exit Function
ErrorOccurred:
    getUB = -1
End Function

'getDictIndex
'Returns the item paired with some key
'@param key - Key to find in dictionary
'@returns - Item stored at key
Private Function getDictItem(ByVal key as string) as variant
  Dim ind as Long: ind = getDictIndex(key)
  if ind <> -1 then 
    if isObject(this.Singleton.FakeDict(ind).val) then
      set getDictItem = this.Singleton.FakeDict(ind).val
    else
      getDictItem = this.Singleton.FakeDict(ind).val
    end if
  else
    getDictItem = Empty
  end if
End Function
