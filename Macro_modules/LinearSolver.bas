Attribute VB_Name = "LinearSolver"
Option Explicit

' Solves G · X = B for X using Gauss elimination with partial pivoting.
' G is (n x n), B is (n x m) (multiple right-hand sides at once).
' Returns X as an (n x m) 1-based Variant array.
' This REPLACES WorksheetFunction.MInverse-followed-by-multiplication,
' which is numerically unreliable in Excel for matrices > ~50x50 with
' mixed-scale entries (exactly what our G_ass becomes at fine
' discretization: ~O(1) force rows mixed with ~O(Delta_s) moment rows).
Public Function SolveLinearSystem(G As Variant, b As Variant) As Variant
    Dim n As Long, m As Long
    n = UBound(G, 1)
    m = UBound(b, 2)

    Debug.Print "SolveLinearSystem: n=" & n & " m=" & m   ' <-- INSIDE the function

    ' Build augmented matrix [G | B]
    Dim A() As Double
    ReDim A(1 To n, 1 To n + m)
    Dim i As Long, j As Long, k As Long
    For i = 1 To n
        For j = 1 To n
            A(i, j) = CDbl(G(i, j))
        Next j
        For j = 1 To m
            A(i, n + j) = CDbl(b(i, j))
        Next j
    Next i

    ' Forward elimination with partial pivoting
    Dim maxVal As Double, maxRow As Long, tmp As Double, factor As Double
    For i = 1 To n
        maxVal = Abs(A(i, i))
        maxRow = i
        For k = i + 1 To n
            If Abs(A(k, i)) > maxVal Then
                maxVal = Abs(A(k, i))
                maxRow = k
            End If
        Next k

        If maxVal < 0.000000000000001 Then
            Err.Raise 5, "SolveLinearSystem", _
                      "Matrix is singular or numerically singular at row " & i
        End If

        If maxRow <> i Then
            For j = i To n + m
                tmp = A(i, j)
                A(i, j) = A(maxRow, j)
                A(maxRow, j) = tmp
            Next j
        End If

        For k = i + 1 To n
            If A(k, i) <> 0 Then
                factor = A(k, i) / A(i, i)
                For j = i To n + m
                    A(k, j) = A(k, j) - factor * A(i, j)
                Next j
            End If
        Next k
    Next i

    ' --- diagnostic: smallest pivot ---
    Dim minPivot As Double: minPivot = 1E+30
    For i = 1 To n
        If Abs(A(i, i)) < minPivot Then minPivot = Abs(A(i, i))
    Next i
    Debug.Print "  smallest pivot = " & minPivot
    ' ----------------------------------

    ' Back-substitution
    Dim X() As Double
    ReDim X(1 To n, 1 To m)
    Dim s As Double
    For k = 1 To m
        For i = n To 1 Step -1
            s = A(i, n + k)
            For j = i + 1 To n
                s = s - A(i, j) * X(j, k)
            Next j
            X(i, k) = s / A(i, i)
        Next i
    Next k

    SolveLinearSystem = X
End Function

Public Function SolveLinearVector(G As Variant, b As Variant) As Variant
    SolveLinearVector = SolveLinearSystem(G, b)
End Function

