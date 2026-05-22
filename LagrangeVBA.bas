Attribute VB_Name = "LagrangeRegression"
Option Explicit

' ===========================================================================
' Lagrange Polynomial Regression  -  Excel Add-in  (.xlam)
' ===========================================================================
' PUBLIC WORKSHEET FUNCTIONS
'   LAGRANGECOEFFS(x_range, y_range, degree)
'       Enter with Ctrl+Shift+Enter; returns a vertical array of coefficients
'       (highest power first) that are written into the selected cells.
'
'   LAGRANGEFIT(x_range, y_range, degree, x_new)
'       Returns the fitted polynomial value at x_new (single cell).
'
' INTERACTIVE MACRO
'   RunLagrangeFit   (run from Developer > Macros, or assign to a button)
'       Prompts for x/y ranges and degree; writes coefficients to column E
'       starting at E1 on the active sheet.
' ===========================================================================


' ---------------------------------------------------------------------------
' INTERNAL: Gaussian-elimination least-squares polynomial fit (polyfit).
'           Returns coefficients(0..degree), highest power at index 0.
' ---------------------------------------------------------------------------
Private Function PolyFit(x() As Double, y() As Double, _
                          ByVal deg As Long) As Double()
    Dim n As Long, m As Long, i As Long, j As Long, k As Long
    n = UBound(x) - LBound(x) + 1
    m = deg + 1

    ' Build Vandermonde matrix A (n x m)
    Dim A() As Double
    ReDim A(0 To n - 1, 0 To m - 1)
    For i = 0 To n - 1
        For j = 0 To m - 1
            A(i, j) = x(LBound(x) + i) ^ (deg - j)
        Next j
    Next i

    ' Normal equations:  ATA = A'*A,  ATy = A'*y
    Dim ATA() As Double, ATy() As Double
    ReDim ATA(0 To m - 1, 0 To m - 1)
    ReDim ATy(0 To m - 1)
    Dim s As Double

    For i = 0 To m - 1
        s = 0
        For k = 0 To n - 1: s = s + A(k, i) * y(LBound(y) + k): Next k
        ATy(i) = s
        For j = 0 To m - 1
            s = 0
            For k = 0 To n - 1: s = s + A(k, i) * A(k, j): Next k
            ATA(i, j) = s
        Next j
    Next i

    ' Augmented matrix [ATA | ATy]
    Dim aug() As Double
    ReDim aug(0 To m - 1, 0 To m)
    For i = 0 To m - 1
        For j = 0 To m - 1: aug(i, j) = ATA(i, j): Next j
        aug(i, m) = ATy(i)
    Next i

    ' Forward elimination with partial pivoting
    Dim pivRow As Long, tmp As Double
    For i = 0 To m - 1
        pivRow = i
        For k = i + 1 To m - 1
            If Abs(aug(k, i)) > Abs(aug(pivRow, i)) Then pivRow = k
        Next k
        If pivRow <> i Then
            For j = 0 To m
                tmp = aug(i, j): aug(i, j) = aug(pivRow, j): aug(pivRow, j) = tmp
            Next j
        End If
        If Abs(aug(i, i)) > 1E-15 Then
            Dim fac As Double
            For k = i + 1 To m - 1
                fac = aug(k, i) / aug(i, i)
                For j = i To m: aug(k, j) = aug(k, j) - fac * aug(i, j): Next j
            Next k
        End If
    Next i

    ' Back substitution
    Dim c() As Double
    ReDim c(0 To m - 1)
    For i = m - 1 To 0 Step -1
        s = aug(i, m)
        For j = i + 1 To m - 1: s = s - aug(i, j) * c(j): Next j
        If Abs(aug(i, i)) > 1E-15 Then c(i) = s / aug(i, i)
    Next i

    PolyFit = c
End Function


' ---------------------------------------------------------------------------
' INTERNAL: Evaluate polynomial (Horner) at a single x value.
' ---------------------------------------------------------------------------
Private Function EvalPoly(c() As Double, ByVal deg As Long, ByVal xv As Double) As Double
    Dim i As Long
    Dim v As Double
    v = c(0)
    For i = 1 To deg: v = v * xv + c(i): Next i
    EvalPoly = v
End Function


' ===========================================================================
' WORKSHEET FUNCTION 1: LAGRANGECOEFFS
' ===========================================================================
' Returns a vertical array of polynomial coefficients (highest power first).
' Select degree+1 cells in a column, type the formula, press Ctrl+Shift+Enter.
'
'   x_range  Single-column range of x (independent) data values.
'   y_range  Single-column range of y (dependent) data values,
'            same size as x_range.
'   degree   Integer degree of the polynomial to fit (1 = linear,
'            2 = quadratic, 3 = cubic, 4 = quartic, ...).
'            Must be < number of data points.
' ===========================================================================
Public Function LAGRANGECOEFFS(x_range As Range, y_range As Range, _
                                ByVal degree As Long) As Variant
    Application.Volatile False

    Dim n As Long
    n = x_range.Cells.Count

    If n < 2 Then
        LAGRANGECOEFFS = CVErr(xlErrValue): Exit Function
    End If
    If y_range.Cells.Count <> n Then
        LAGRANGECOEFFS = CVErr(xlErrNA): Exit Function
    End If
    If degree < 1 Or degree >= n Then
        LAGRANGECOEFFS = CVErr(xlErrNum): Exit Function
    End If

    Dim x() As Double, y() As Double
    ReDim x(1 To n), y(1 To n)
    Dim i As Long
    For i = 1 To n
        If Not IsNumeric(x_range.Cells(i).Value) Or _
           Not IsNumeric(y_range.Cells(i).Value) Then
            LAGRANGECOEFFS = CVErr(xlErrValue): Exit Function
        End If
        x(i) = CDbl(x_range.Cells(i).Value)
        y(i) = CDbl(y_range.Cells(i).Value)
    Next i

    Dim c() As Double
    c = PolyFit(x, y, degree)

    Dim m As Long
    m = degree + 1
    Dim res() As Variant
    ReDim res(1 To m, 1 To 1)
    For i = 1 To m
        res(i, 1) = c(i - 1)
    Next i

    LAGRANGECOEFFS = res
End Function


' ===========================================================================
' WORKSHEET FUNCTION 2: LAGRANGEFIT
' ===========================================================================
' Evaluates the fitted polynomial at a single new x value.
'
'   x_range  Single-column range of x (training) data values.
'   y_range  Single-column range of y (training) data values.
'   degree   Integer degree of the polynomial (must be < number of points).
'   x_new    The x value at which to evaluate the fitted polynomial.
' ===========================================================================
Public Function LAGRANGEFIT(x_range As Range, y_range As Range, _
                             ByVal degree As Long, _
                             ByVal x_new As Double) As Variant
    Application.Volatile False

    Dim n As Long
    n = x_range.Cells.Count

    If n < 2 Or y_range.Cells.Count <> n Then
        LAGRANGEFIT = CVErr(xlErrNA): Exit Function
    End If
    If degree < 1 Or degree >= n Then
        LAGRANGEFIT = CVErr(xlErrNum): Exit Function
    End If

    Dim x() As Double, y() As Double
    ReDim x(1 To n), y(1 To n)
    Dim i As Long
    For i = 1 To n
        x(i) = CDbl(x_range.Cells(i).Value)
        y(i) = CDbl(y_range.Cells(i).Value)
    Next i

    Dim c() As Double
    c = PolyFit(x, y, degree)

    LAGRANGEFIT = EvalPoly(c, degree, x_new)
End Function


' ===========================================================================
' INTERACTIVE MACRO: RunLagrangeFit
' ===========================================================================
' Prompts for x range, y range, and polynomial degree, then writes:
'   E1  -> label "Coeff (deg N, high→low)"
'   E2  -> coefficient of x^N  (highest power)
'   E3  -> coefficient of x^(N-1)
'   ...
'   E(N+2) -> constant term (x^0)
' ===========================================================================
Public Sub RunLagrangeFit()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- X range ---
    Dim rngX As Range
    On Error Resume Next
    Set rngX = Application.InputBox( _
        "Select the X data range (single column of numeric values)." & vbLf & _
        "Example:  A2:A22", _
        "Step 1 of 3  —  X Data Range", Type:=8)
    On Error GoTo 0
    If rngX Is Nothing Then Exit Sub

    ' --- Y range ---
    Dim rngY As Range
    On Error Resume Next
    Set rngY = Application.InputBox( _
        "Select the Y data range (single column, same number of rows as X)." & vbLf & _
        "Example:  B2:B22", _
        "Step 2 of 3  —  Y Data Range", Type:=8)
    On Error GoTo 0
    If rngY Is Nothing Then Exit Sub

    ' --- Validate sizes ---
    Dim n As Long
    n = rngX.Cells.Count
    If rngY.Cells.Count <> n Then
        MsgBox "X and Y ranges must have the same number of cells." & vbLf & _
               "X has " & n & " cells; Y has " & rngY.Cells.Count & " cells.", _
               vbExclamation, "Size Mismatch"
        Exit Sub
    End If

    ' --- Degree ---
    Dim maxDeg As Long
    maxDeg = n - 1
    Dim degStr As String
    degStr = InputBox( _
        "Enter the polynomial degree to fit." & vbLf & vbLf & _
        "  Valid range : 1 to " & maxDeg & vbLf & _
        "  1  =  linear" & vbLf & _
        "  2  =  quadratic" & vbLf & _
        "  3  =  cubic" & vbLf & _
        "  4  =  quartic" & vbLf & vbLf & _
        "Data points loaded: " & n, _
        "Step 3 of 3  —  Polynomial Degree", "4")
    If degStr = "" Then Exit Sub

    Dim deg As Long
    If Not IsNumeric(degStr) Then
        MsgBox "Please enter a whole number.", vbExclamation: Exit Sub
    End If
    deg = CLng(CDbl(degStr))
    If deg < 1 Or deg > maxDeg Then
        MsgBox "Degree must be between 1 and " & maxDeg & ".", vbExclamation: Exit Sub
    End If

    ' --- Read data ---
    Dim xArr() As Double, yArr() As Double
    ReDim xArr(1 To n), yArr(1 To n)
    Dim i As Long
    For i = 1 To n
        If Not IsNumeric(rngX.Cells(i).Value) Or _
           Not IsNumeric(rngY.Cells(i).Value) Then
            MsgBox "Non-numeric value found at row " & i & " of the selected ranges.", _
                   vbExclamation: Exit Sub
        End If
        xArr(i) = CDbl(rngX.Cells(i).Value)
        yArr(i) = CDbl(rngY.Cells(i).Value)
    Next i

    ' --- Fit ---
    Dim c() As Double
    c = PolyFit(xArr, yArr, deg)

    ' --- Write to column E ---
    ws.Cells(1, 5).Value = "Coeff (deg " & deg & ", high" & Chr(8594) & "low)"

    Dim m As Long
    m = deg + 1
    For i = 0 To m - 1
        ws.Cells(i + 2, 5).Value = c(i)
    Next i

    ' Clear stale rows below
    If m + 2 <= 200 Then
        ws.Range(ws.Cells(m + 2, 5), ws.Cells(200, 5)).ClearContents
    End If

    MsgBox m & " coefficients written to E1:E" & m + 1 & " on sheet """ & _
           ws.Name & """." & vbLf & vbLf & _
           "Polynomial: p(x) = c0*x^" & deg & " + ... + c" & deg, _
           vbInformation, "Lagrange Regression — Done"
End Sub


' ===========================================================================
' Register functions in Excel Function Wizard (runs when add-in loads)
' ===========================================================================
Private Sub RegisterFunctions()
    On Error Resume Next
    Application.MacroOptions _
        Macro:="LAGRANGECOEFFS", _
        Description:="Fits a least-squares polynomial of the chosen degree " & _
                     "to x/y data (Lagrange basis). Returns a vertical array " & _
                     "of degree+1 coefficients, highest power first. " & _
                     "Enter with Ctrl+Shift+Enter.", _
        Category:="Statistical", _
        ArgumentDescriptions:=Array( _
            "x_range — Single-column range containing the x (independent) data values.", _
            "y_range — Single-column range containing the y (dependent) data values; " & _
                      "must be the same length as x_range.", _
            "degree  — Polynomial degree to fit (integer >= 1, must be less than the " & _
                      "number of data points). 1=linear, 2=quadratic, 3=cubic, 4=quartic.")

    Application.MacroOptions _
        Macro:="LAGRANGEFIT", _
        Description:="Evaluates the Lagrange least-squares polynomial fit at a " & _
                     "single new x value.", _
        Category:="Statistical", _
        ArgumentDescriptions:=Array( _
            "x_range — Single-column range of the original x (training) data.", _
            "y_range — Single-column range of the original y (training) data; " & _
                      "must match x_range length.", _
            "degree  — Polynomial degree to fit (integer >= 1, < number of data points).", _
            "x_new   — The x value at which to evaluate the fitted polynomial.")
    On Error GoTo 0
End Sub

Public Sub Auto_Open()
    Call RegisterFunctions
End Sub
