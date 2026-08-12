Attribute VB_Name = "Module4"
Sub CreateAinSheet()
    Dim wsGass As Worksheet, wsGassIn As Worksheet, wsFass As Worksheet, wsAin As Worksheet
    Dim wbUser As Workbook
    Dim lastRow As Long
    Dim G1_ass As Variant, G2_ass As Variant
    Dim G1_ass_in As Variant, G2_ass_in As Variant
    Dim F_ass_lambda As Variant
    Dim A_in_Excel As Variant
    Dim numCols As Long

    On Error GoTo ErrorHandler

    Set wbUser = ActiveWorkbook
    Set wsGass = wbUser.Sheets("Gass")
    Set wsGassIn = wbUser.Sheets("Gass_in")
    Set wsFass = wbUser.Sheets("Fass")

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("Ain").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsAin = wbUser.Sheets.Add
    wsAin.Name = "Ain"

    With wsGass
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        numCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        G1_ass = .Range(.Cells(1, 1), .Cells(lastRow, numCols - 3)).Value
        G2_ass = .Range(.Cells(1, numCols - 2), .Cells(lastRow, numCols)).Value
    End With

    With wsGassIn
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        numCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        G1_ass_in = .Range(.Cells(1, 1), .Cells(lastRow, numCols - 3)).Value
        G2_ass_in = .Range(.Cells(1, numCols - 2), .Cells(lastRow, numCols)).Value
    End With

    With wsFass
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        F_ass_lambda = .Range("B2:B" & lastRow).Value
    End With

    ' ============================================================
    ' CRITICAL CHANGE: solve G1_ass * Y = [G2_ass | F_lambda] via LU
    ' decomposition instead of forming MInverse(G1_ass) explicitly.
    ' MInverse silently corrupts large (~165x165) mixed-scale matrices,
    ' which is why refinement from 11 to 55 voussoirs breaks the analysis.
    ' ============================================================

    ' Build combined RHS: [G2_ass | F_ass_lambda] with (3*NB) rows
    Dim nRows As Long, nCols_G2 As Long
    nRows = UBound(G1_ass, 1)
    nCols_G2 = UBound(G2_ass, 2)  ' should be 3

    Dim combinedRHS() As Double
    ReDim combinedRHS(1 To nRows, 1 To nCols_G2 + 1)
    Dim i As Long, j As Long
    For i = 1 To nRows
        For j = 1 To nCols_G2
            combinedRHS(i, j) = CDbl(G2_ass(i, j))
        Next j
        combinedRHS(i, nCols_G2 + 1) = CDbl(F_ass_lambda(i, 1))
    Next i

    ' Solve G1_ass * Ysol = combinedRHS in ONE call
    Dim Ysol As Variant
    Ysol = SolveLinearSystem(G1_ass, combinedRHS)

    ' Split Ysol back into Y_G2 (first 3 cols) and Y_lam (last col)
    Dim Y_G2() As Double, Y_lam() As Double
    ReDim Y_G2(1 To nRows, 1 To nCols_G2)
    ReDim Y_lam(1 To nRows, 1 To 1)
    For i = 1 To nRows
        For j = 1 To nCols_G2
            Y_G2(i, j) = Ysol(i, j)
        Next j
        Y_lam(i, 1) = Ysol(i, nCols_G2 + 1)
    Next i

    ' temp1 = -G1_ass_in * Y_G2 + G2_ass_in
    Dim temp_neg As Variant, temp_add As Variant
    temp_neg = MultiplyMatrixByScalar(MultiplyMatrices(G1_ass_in, Y_G2), -1)
    temp_add = AddMatrices(temp_neg, G2_ass_in)

    ' temp2 = G1_ass_in * Y_lam
    Dim temp_lam As Variant
    temp_lam = MultiplyMatrices(G1_ass_in, Y_lam)

    ' A_in_Excel = [temp_add | temp_lam]
    A_in_Excel = CombineMatricesHorizontal(temp_add, temp_lam)

    With wsAin
        .Range("A1").Value = "A_in_Excel"
        WriteMatrixToRange .Range("A2"), A_in_Excel
        .Columns.AutoFit
        With .Range("A1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(200, 200, 200)
        End With
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Error in CreateAinSheet: " & Err.Description, vbCritical
End Sub

Sub CreateBinSheet()
    Dim wsGass As Worksheet, wsGassIn As Worksheet, wsFass As Worksheet, wsBin As Worksheet
    Dim wbUser As Workbook
    Dim lastRow As Long
    Dim G1_ass As Variant
    Dim G1_ass_in As Variant
    Dim F_ass_W As Variant
    Dim b_ass_in As Variant
    Dim b_in_Excel As Variant
    Dim numCols As Long

    On Error GoTo ErrorHandler

    Set wbUser = ActiveWorkbook
    Set wsGass = wbUser.Sheets("Gass")
    Set wsGassIn = wbUser.Sheets("Gass_in")
    Set wsFass = wbUser.Sheets("Fass")

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("Bin").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsBin = wbUser.Sheets.Add
    wsBin.Name = "Bin"

    With wsGass
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        numCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        G1_ass = .Range(.Cells(1, 1), .Cells(lastRow, numCols - 3)).Value
    End With

    With wsGassIn
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        numCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
        G1_ass_in = .Range(.Cells(1, 1), .Cells(lastRow, numCols - 3)).Value
        b_ass_in = .Range(.Cells(1, numCols + 1), .Cells(lastRow, numCols + 1)).Value
    End With

    With wsFass
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        F_ass_W = .Range("A2:A" & lastRow).Value
    End With

    ' ============================================================
    ' CRITICAL CHANGE: solve G1_ass * Y = F_ass_W via LU instead of MInverse.
    ' ============================================================
    Dim Y_W As Variant
    Y_W = SolveLinearSystem(G1_ass, F_ass_W)

    ' b_in_Excel = b_ass_in - G1_ass_in * Y_W
    Dim temp As Variant
    temp = MultiplyMatrixByScalar(MultiplyMatrices(G1_ass_in, Y_W), -1)
    b_in_Excel = AddMatrices(b_ass_in, temp)

    With wsBin
        .Range("A1").Value = "b_in_Excel"
        WriteMatrixToRange .Range("A2"), b_in_Excel
        .Columns.AutoFit
        With .Range("A1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(200, 200, 200)
        End With
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Error in CreateBinSheet: " & Err.Description, vbCritical
End Sub

Function MultiplyMatrices(mat1 As Variant, mat2 As Variant) As Variant
    Dim result() As Double
    Dim i As Long, j As Long, k As Long
    Dim sum As Double
    ReDim result(1 To UBound(mat1, 1), 1 To UBound(mat2, 2))
    For i = 1 To UBound(mat1, 1)
        For j = 1 To UBound(mat2, 2)
            sum = 0
            For k = 1 To UBound(mat1, 2)
                sum = sum + mat1(i, k) * mat2(k, j)
            Next k
            result(i, j) = sum
        Next j
    Next i
    MultiplyMatrices = result
End Function

Function AddMatrices(mat1 As Variant, mat2 As Variant) As Variant
    Dim result() As Double
    Dim i As Long, j As Long
    ReDim result(1 To UBound(mat1, 1), 1 To UBound(mat1, 2))
    For i = 1 To UBound(mat1, 1)
        For j = 1 To UBound(mat1, 2)
            result(i, j) = mat1(i, j) + mat2(i, j)
        Next j
    Next i
    AddMatrices = result
End Function

Function MultiplyMatrixByScalar(mat As Variant, scalar As Double) As Variant
    Dim result() As Double
    Dim i As Long, j As Long
    ReDim result(1 To UBound(mat, 1), 1 To UBound(mat, 2))
    For i = 1 To UBound(mat, 1)
        For j = 1 To UBound(mat, 2)
            result(i, j) = mat(i, j) * scalar
        Next j
    Next i
    MultiplyMatrixByScalar = result
End Function

Function CombineMatricesHorizontal(mat1 As Variant, mat2 As Variant) As Variant
    Dim result() As Double
    Dim i As Long, j As Long
    ReDim result(1 To UBound(mat1, 1), 1 To UBound(mat1, 2) + UBound(mat2, 2))
    For i = 1 To UBound(mat1, 1)
        For j = 1 To UBound(mat1, 2)
            result(i, j) = mat1(i, j)
        Next j
        For j = 1 To UBound(mat2, 2)
            result(i, UBound(mat1, 2) + j) = mat2(i, j)
        Next j
    Next i
    CombineMatricesHorizontal = result
End Function

Sub WriteMatrixToRange(rng As Range, matrix As Variant)
    Dim outputRange As Range
    Set outputRange = rng.Resize(UBound(matrix, 1), UBound(matrix, 2))
    outputRange.Value = matrix
End Sub
