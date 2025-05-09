Attribute VB_Name = "modPolynomReg"

'@Folder("PolynomReg")

Option Explicit

'module originally from Gerhard Krucker which was created 17.08.1995, 22.09.1996
'and 2004 for VBA in EXCEL7. It was highly modified by Stefan Pinnow in 2016
'and 2017.
'<http://www.krucker.ch/skripten-uebungen/IAMSkript/IAMKap3.pdf>

'==============================================================================
'requires the functions
'- ChangeBoundsOfVector
'- CopyArray
'- GetColumn
'- GetRow
'- IsArrayAllNumeric
'- NumberOfArrayDimensions
'from the revised 'modArraySupport2' module from Chip Pearson originally
'available at
'  <http://www.cpearson.com/excel/VBAArrays.htm>
'and the functions
'- RangeToArray
'- VariableType
'from the 'modUsefulFunctions' module
'==============================================================================

Public Sub AddUDFToCustomCategory()
    
    '==========================================================================
    'how should the category be named?
'    Const vCategory As String = "Math. & Trigonom."
    Const vCategory As Long = 3   '"Math. & Trigonom."
    '==========================================================================
    
    With Application
        .MacroOptions _
            Category:=vCategory, _
            Macro:="Polynom", _
            Description:="Calculates polynomial expression " & _
                "f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n", _
            ArgumentDescriptions:=Array( _
                "Coefficients (a0, a1, a2, ...)", _
                "Independent variable (x)", _
                "(Optional) TRUE = interpret #NA's as 0's" _
            )
        .MacroOptions _
            Category:=vCategory, _
            Macro:="PolynomReg", _
            Description:="Calculates polynomial coefficients (a0,...,an)", _
            ArgumentDescriptions:=Array( _
                "Array of 'x' values", _
                "Array of 'y' values", _
                "Polynomial degree", _
                "(Optional) TRUE = return coefficients vertically", _
                "(Optional) TRUE = ignore #NA entries" _
            )
        .MacroOptions _
            Category:=vCategory, _
            Macro:="PolynomRegRel", _
            Description:="Calculates polynomial coefficients (a0,...,an)", _
            ArgumentDescriptions:=Array( _
                "Array of 'x' values", _
                "Array of 'y' values", _
                "Polynomial degree", _
                "(Optional) TRUE = return coefficients vertically", _
                "(Optional) TRUE = ignore #NA entries" _
            )
    End With
    
End Sub


'@Description("Calculates polynomial expression f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n")
Public Function Polynom( _
    ByVal Coefficients As Variant, _
    ByVal x As Variant, _
    Optional ByVal IgnoreNA As Variant _
        ) As Variant
Attribute Polynom.VB_Description = "Calculates polynomial expression f(x) = a0 + a1*x + a2*x^2 + ... + an*x^n"
    
    '---
    'IgnoreNA' must be a boolean
    If IsMissing(IgnoreNA) Or IsEmpty(IgnoreNA) Then
        IgnoreNA = False
    ElseIf Not VariableType(IgnoreNA) = "Boolean" Then
        GoTo errHandler
    End If
    '---
    
    
    'convert possible range to array
    Coefficients = Coefficients
    
    Dim arrCoeffs() As Variant
    If Not ExtractVector(Coefficients, arrCoeffs) Then GoTo errHandler
    
    Dim xArr As Variant
    xArr = ConvertToArray(x)
    
    'if 'IgnoreNA' is 'TRUE' then remove all trailing 'NAs' lines
    '(this only makes sense if more than one coefficient is given)
    If IgnoreNA = True Then
        If Not RemoveNALines(arrCoeffs) Then GoTo errHandler
    End If
    If Not IsArrayAllNumeric(arrCoeffs) Then GoTo errHandler
    
    Dim sum() As Double
    ReDim sum(LBound(xArr) To UBound(xArr))
    
    Dim j As Long
    For j = LBound(xArr) To UBound(xArr)
        'apply Horner scheme
        Dim i As Long
        For i = UBound(arrCoeffs) To LBound(arrCoeffs) Step -1
            sum(j) = arrCoeffs(i) + sum(j) * xArr(j)
        Next
    Next
    
    If TypeName(x) = "Range" Then
        If x.Columns.Count = 1 Then
            Polynom = Application.WorksheetFunction.Transpose(sum)
        Else
            Polynom = sum
        End If
    'to avoid a type mismatch error in case only a single return value is
    'allowed like for 'Debug.Print'
    ElseIf Not IsArray(x) Then
        Polynom = sum(1)
    Else
        Polynom = sum
    End If
    Exit Function
    
    
errHandler:
    Polynom = CVErr(xlErrNA)
    
End Function


'==============================================================================
'calculate the polynomial coefficients a0,...,an for the polynomial trend
'function n-th degree for m data points using the method of least squares.
'Parameter:
'- x                = array of x values (number of points: m, any)
'- y                = array of y values (number of points: m, any)
'- PolynomialDegree = degree of to generate polynomial trend function
'- VerticalOutput   = optional argument to allow a vertical output of the
'                     polynomial coefficients
'- IgnoreNAs        = optional argument to ignore "NA" data points
'The result will be returned as array (vector)
'@Description("Calculates polynomial coefficients (a0,...,an)")
Public Function PolynomReg( _
    ByVal x As Variant, _
    ByVal y As Variant, _
    ByVal PolynomialDegree As Long, _
    Optional ByVal VerticalOutput As Variant, _
    Optional ByVal IgnoreNAs As Variant _
        ) As Variant
Attribute PolynomReg.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
    
    '---
    ''VerticalOutput' must be a boolean
    If IsMissing(VerticalOutput) Or IsEmpty(VerticalOutput) Then
        VerticalOutput = False
    ElseIf Not VariableType(VerticalOutput) = "Boolean" Then
        GoTo errHandler
    End If
    
    ''IgnoreNAs' must be a boolean
    If IsMissing(IgnoreNAs) Or IsEmpty(IgnoreNAs) Then
        IgnoreNAs = False
    ElseIf Not VariableType(IgnoreNAs) = "Boolean" Then
        GoTo errHandler
    End If
    '---
    
    
    PolynomReg = MasterPolynomReg( _
            x, y, _
            PolynomialDegree, _
            CBool(VerticalOutput), _
            CBool(IgnoreNAs), _
            False _
    )
    Exit Function
    
    
errHandler:
    PolynomReg = CVErr(xlErrNA)
    
End Function


'calculate the polynomial coefficients a0,...,an for the polynomial trend
'function n-th degree for m data points using the method of least relative
'squares.
'Parameter:
'- x                = array of x values (number of points: m, any)
'- y                = array of y values (number of points: m, any)
'- PolynomialDegree = degree of to generate polynomial trend function
'- VerticalOutput   = optional argument to allow a vertical output of the
'                     polynomial coefficients
'- IgnoreNAs        = optional argument to ignore "NA" data points
'The result will be returned as array (vector)
'@Description("Calculates polynomial coefficients (a0,...,an)")
Public Function PolynomRegRel( _
    ByVal x As Variant, _
    ByVal y As Variant, _
    ByVal PolynomialDegree As Long, _
    Optional ByVal VerticalOutput As Variant, _
    Optional ByVal IgnoreNAs As Variant _
        ) As Variant
Attribute PolynomRegRel.VB_Description = "Calculates polynomial coefficients (a0,...,an)"
Attribute PolynomRegRel.VB_ProcData.VB_Invoke_Func = " \n3"
    
    '---
    ''VerticalOutput' must be a boolean
    If IsMissing(VerticalOutput) Or IsEmpty(VerticalOutput) Then
        VerticalOutput = False
    ElseIf Not VariableType(VerticalOutput) = "Boolean" Then
        GoTo errHandler
    End If
    
    ''IgnoreNAs' must be a boolean
    If IsMissing(IgnoreNAs) Or IsEmpty(IgnoreNAs) Then
        IgnoreNAs = False
    ElseIf Not VariableType(IgnoreNAs) = "Boolean" Then
        GoTo errHandler
    End If
    '---
    
    
    PolynomRegRel = MasterPolynomReg( _
            x, y, _
            PolynomialDegree, _
            CBool(VerticalOutput), _
            CBool(IgnoreNAs), _
            True _
    )
    Exit Function
    
    
errHandler:
    PolynomRegRel = CVErr(xlErrNA)
    
End Function


Private Function MasterPolynomReg( _
    ByVal x As Variant, _
    ByVal y As Variant, _
    ByVal PolynomialDegree As Long, _
    ByVal VerticalOutput As Boolean, _
    ByVal IgnoreNAs As Boolean, _
    ByVal UseRelativeVersion As Boolean _
        ) As Variant
    
    '---
    ''PolynomialDegree' has to be an integer >= 0
    If PolynomialDegree < 0 Then
        GoTo errHandler
    End If
    '---
    
    
    'convert 'x' and 'y' to arrays if they are ranges
    If TypeName(x) = "Range" Then
        Dim xArr As Variant
        xArr = RangeToArray(x)
    Else
        xArr = x
    End If
    If TypeName(y) = "Range" Then
        Dim yArr As Variant
        yArr = RangeToArray(y)
    Else
        yArr = y
    End If
    
    'count number of data points in given arrays
    Dim CountX As Long
    CountX = UBound(xArr) - LBound(xArr) + 1
    Dim CountY As Long
    CountY = UBound(yArr) - LBound(yArr) + 1
    
    'the number of points has to be identical for 'xArr' and 'yArr'
    If CountX <> CountY Then GoTo errHandler
    
    'the polynomial coefficient must be smaller than the number of given points
    If CountX <= PolynomialDegree Then GoTo errHandler
    
    
    'prepare vectors 'xWithoutNAs' and 'yWithoutNAs'
    '(which are then used to calculate the polynomial coefficients)
    If IgnoreNAs = False Then
        If Not IsArrayAllNumeric(xArr) Then GoTo errHandler
        If Not IsArrayAllNumeric(yArr) Then GoTo errHandler
        
        Dim xWithoutNAs() As Double
        If Not ExtractVector(xArr, xWithoutNAs) Then GoTo errHandler
        Dim yWithoutNAs() As Double
        If Not ExtractVector(yArr, yWithoutNAs) Then GoTo errHandler
    Else
        'else copy 'xArr' to 'xAsVector' and 'yArr' to 'yAsVector'
        Dim xAsVector() As Variant
        If Not ExtractVector(xArr, xAsVector) Then GoTo errHandler
        Dim yAsVector() As Variant
        If Not ExtractVector(yArr, yAsVector) Then GoTo errHandler
        
        If Not CopyOnlyNonNALines( _
                xAsVector, yAsVector, _
                xWithoutNAs, yWithoutNAs, _
                PolynomialDegree _
        ) Then GoTo errHandler
    End If
    
    '--------------------------------------------------------------------------
    
    'transfer (new) number of 'x' elements to 'NoOfNonNADataPoints'
    Dim NoOfNonNADataPoints As Long
    NoOfNonNADataPoints = UBound(xWithoutNAs) - LBound(xWithoutNAs) + 1
    'check again, if number of (real) data points is smaller than the given
    'polynomial degree
    If NoOfNonNADataPoints <= PolynomialDegree Then GoTo errHandler
    
    Dim CoefficientMatrix() As Double
    CoefficientMatrix = Calculate_CoefficientMatrix( _
            xWithoutNAs, yWithoutNAs, _
            PolynomialDegree, _
            UseRelativeVersion _
    )
    Dim VectorOfConstants() As Double
    VectorOfConstants = Calculate_VectorOfConstants( _
            xWithoutNAs, yWithoutNAs, _
            PolynomialDegree _
    )
    
    'invert coefficient matrix 'CoefficientMatrix'
    '(MINVERSE can't write back to 'CoefficientMatrix')
    '(please note that the resulting array starts with index 1)
    '(it has to be a variant to be able to use 'WorksheetFunction.MInverse')
    Dim InverseCoefficientMatrix As Variant
    InverseCoefficientMatrix = Application.WorksheetFunction.MInverse(CoefficientMatrix)
    
    'dynamic array for the polynomial coefficients a0,...,an
    '(it has to be of type Variant' because of the special handler for
    ' 'PolynomialDegree = 0')
    Dim a() As Variant
    a = Calculate_PolynomialCoefficients( _
            InverseCoefficientMatrix, _
            VectorOfConstants, _
            PolynomialDegree _
    )
    
    If PolynomialDegree = 0 Then
        Call HandleSpecialCaseForPolynomialDegreeEqualsZero(a)
    End If
    
    'return coefficient vector a_0,...,a_n
    If VerticalOutput = True Then
        MasterPolynomReg = Application.WorksheetFunction.Transpose(a)
    Else
        MasterPolynomReg = a
    End If
    
    Exit Function
    
    
errHandler:
    MasterPolynomReg = CVErr(xlErrNA)
    
End Function


'==============================================================================
Private Function Calculate_CoefficientMatrix( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Long, _
    ByVal UseRelativeVersion As Boolean _
        ) As Double()
    
    Dim SumOfPowersXK() As Double
    SumOfPowersXK = Calculate_SumOfPowersXK( _
            x, y, _
            PolynomialDegree, _
            UseRelativeVersion _
    )
    
    Dim CoefficientMatrix() As Double
    ReDim CoefficientMatrix(0 To PolynomialDegree, 0 To PolynomialDegree)
    
    Dim i As Long
    For i = LBound(CoefficientMatrix, 1) To UBound(CoefficientMatrix, 1)
        Dim j As Long
        For j = LBound(CoefficientMatrix, 2) To i
            CoefficientMatrix(i, j) = SumOfPowersXK(i + j)
            CoefficientMatrix(j, i) = SumOfPowersXK(i + j)
        Next
    Next
    
    Calculate_CoefficientMatrix = CoefficientMatrix
    
End Function


'calculate sum of powers 'xk' and store it in a corresponding array
Private Function Calculate_SumOfPowersXK( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Long, _
    ByVal UseRelativeVersion As Boolean _
        ) As Double()
    
    Dim SumOfPowersXK() As Double
    ReDim SumOfPowersXK(0 To 2 * PolynomialDegree)
    
    If UseRelativeVersion = True Then
        Dim i As Long
        For i = LBound(SumOfPowersXK) To UBound(SumOfPowersXK)
            SumOfPowersXK(i) = 0
            Dim k As Long
            For k = LBound(x) To UBound(x)
                SumOfPowersXK(i) = SumOfPowersXK(i) + x(k) ^ i / y(k) ^ 2
            Next
        Next
    Else
        For i = LBound(SumOfPowersXK) To UBound(SumOfPowersXK)
            SumOfPowersXK(i) = 0
            For k = LBound(x) To UBound(x)
                SumOfPowersXK(i) = SumOfPowersXK(i) + x(k) ^ i
            Next
        Next
    End If
    
    Calculate_SumOfPowersXK = SumOfPowersXK
    
End Function


Private Function Calculate_VectorOfConstants( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal PolynomialDegree As Long _
        ) As Double()
    
    'dynamic array for the sum of powers for 'xk*yk'
    Dim SumOfPowersXKYK() As Double
    ReDim SumOfPowersXKYK(0 To PolynomialDegree)
    
    'calculate sum of powers 'xk*yk' and store it in a corresponding array
    Dim i As Long
    For i = LBound(SumOfPowersXKYK) To UBound(SumOfPowersXKYK)
        SumOfPowersXKYK(i) = 0
        Dim k As Long
        For k = LBound(x) To UBound(x)
            SumOfPowersXKYK(i) = SumOfPowersXKYK(i) + x(k) ^ i * y(k)
        Next
    Next
    
    'the sum of powers 'xk*yk' is the vector of constants
    Calculate_VectorOfConstants = SumOfPowersXKYK
    
End Function


'solve system of equations 'CoefficientMatrix * a = c' with matrix inversion
Private Function Calculate_PolynomialCoefficients( _
    ByVal InverseCoefficientMatrix As Variant, _
    ByRef VectorOfConstants() As Double, _
    ByVal PolynomialDegree As Long _
        ) As Variant
    
    'polynomial coefficients a0,...,an (a(0) = a0)
    Dim a() As Variant
    ReDim a(0 To PolynomialDegree)
    
    'matrix multiplication 'a = G_inverse * VectorOfConstants'
'---
'<https://stackoverflow.com/a/7307992>
'   a = WorksheetFunction.MMult(InverseCoefficientMatrix, VectorOfConstants)
'---
    'as a reminder: 'InverseCoefficientMatrix' is a 1-based array
    'which is coming from 'WorksheetFunction.MInverse'
    'it is also needed a special handler for 'PolynomialDegree = 0' because of
    'an anomaly (bug?) in Excel, where a 1D array (z(1 to ..., 1 to 1)) is
    'returned as vector (z(1 to ...)) after inversion with 'WorksheetFunction.MInverse'
    '(see <https://stackoverflow.com/a/28800474/5776000>)
    If PolynomialDegree = 0 Then
        a(0) = InverseCoefficientMatrix(1) * VectorOfConstants(0)
    Else
        Dim i As Long
        For i = LBound(a) To UBound(a)
            a(i) = 0
            Dim j As Long
            For j = LBound(a) To UBound(a)
                a(i) = a(i) + InverseCoefficientMatrix(i + 1, j + 1) * VectorOfConstants(j)
            Next
        Next
    End If
    
    Calculate_PolynomialCoefficients = a
    
End Function


'needed because otherwise the returned array will consist of the a(0) value in
'*all* cells (and not '#NA' values for the "unused coefficients)
Private Sub HandleSpecialCaseForPolynomialDegreeEqualsZero( _
    ByRef a() As Variant _
)
    
    ReDim Preserve a(0 To 1)
    a(1) = CVErr(xlErrNA)
    
End Sub


'==============================================================================
'convert 'target' to an array
Private Function ConvertToArray(target As Variant) As Variant
    
    If Not IsArray(target) Then
        Dim Scalar(1 To 1) As Variant
        Scalar(1) = target
        
        Dim Arr As Variant
        Arr = Scalar
    ElseIf TypeName(target) = "Range" Then
        Arr = RangeToArray(target)
    Else
        Arr = target
    End If
    
    ConvertToArray = Arr
    
End Function


'function to make vectors of the ranges/arrays and optionally only transfer
'non-NA values
Private Function ExtractVector( _
    ByVal Source As Variant, _
    ByRef DestVector As Variant _
        ) As Boolean
    
    Select Case NumberOfArrayDimensions(Source)
        Case 2
            If UBound(Source, 1) > 1 And UBound(Source, 2) = 1 Then
                If Not GetColumn(Source, DestVector, 1) Then Exit Function
            ElseIf UBound(Source, 1) = 1 And UBound(Source, 2) > 1 Then
                If Not GetRow(Source, DestVector, 1) Then Exit Function
            Else
                Exit Function
            End If
        Case 1
            If Not CopyArray(Source, DestVector, False) Then Exit Function
            Dim N As Long
            N = UBound(DestVector) - LBound(DestVector) + 1
            If Not ChangeBoundsOfVector(DestVector, 1, N) Then Exit Function
        Case 0
            ReDim DestVector(0 To 0)
            DestVector(0) = Source
        Case Else
    End Select
    
    ExtractVector = True
    
End Function


Private Function CopyOnlyNonNALines( _
    ByVal xSource As Variant, _
    ByVal ySource As Variant, _
    ByRef xDest As Variant, _
    ByRef yDest As Variant, _
    ByVal PolynomialDegree As Long _
        ) As Boolean
    
    'instantiate 'xDest' and 'yDest'
    ReDim xDest(1 To UBound(xSource) - LBound(xSource) + 1)
    ReDim yDest(1 To UBound(xSource) - LBound(xSource) + 1)
    
    'cycle through each entry
    Dim i As Long
    For i = LBound(xSource) To UBound(xSource)
        'if both values are of numeric type then transfer them to 'xDest' and 'yDest'
        If IsNumeric(xSource(i)) And IsNumeric(ySource(i)) Then
            Dim j As Long
            j = j + 1
            xDest(j) = xSource(i)
            yDest(j) = ySource(i)
        'if not, it is allowed that the values are of the error type 'NA'
        ElseIf Application.WorksheetFunction.IsNA(xSource(i)) Or _
                Application.WorksheetFunction.IsNA(ySource(i)) Then
        'else at least one of the 'xSource' or 'ySource' points contains a
        'not allowed value
        Else
            CopyOnlyNonNALines = False
            Exit Function
        End If
    Next
    
    'ReDim 'xDest' and 'yDest' to only the populated values
    ReDim Preserve xDest(1 To j)
    ReDim Preserve yDest(1 To j)
    
    
    'check again, if the polynomial coefficient is smaller than the number of
    'given points
    If j > PolynomialDegree Then
        CopyOnlyNonNALines = True
    End If
    
End Function


Private Function RemoveNALines(ByRef Arr As Variant) As Boolean
    
    Dim i As Long
    For i = UBound(Arr) To LBound(Arr) Step -1
        'if the actual coefficient is not a number ...
        If Not IsNumeric(Arr(i)) Then
            '... and not the error value 'NA' then exit the function
            If Arr(i) <> CVErr(xlErrNA) Then
                RemoveNALines = False
                Exit Function
            End If
        'if it is a number stop cycling
        '(because then only numbers will follow)
        Else
            Exit For
        End If
    Next
    
    'it could be the case that *all* lines are removed
    If i >= LBound(Arr) Then
        'ReDim 'Arr' to the numeric values only
        ReDim Preserve Arr(LBound(Arr) To i)
        RemoveNALines = True
    Else
        RemoveNALines = False
    End If
    
End Function
