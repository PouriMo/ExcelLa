Attribute VB_Name = "Module2"
Sub GenerateGassMatrix()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim wbUser As Workbook
    Dim i As Long, matrixCount As Long
    Dim theta_i As Double, theta_i_1 As Double, delta_s As Double
    Dim totalRows As Long, totalCols As Long, rad As Double
    Dim r As Long, c As Long, val As Double
    Dim fullMatrix() As Double

    Set wbUser = ActiveWorkbook
    Set wsSource = wbUser.ActiveSheet

    On Error Resume Next
    Set wsTarget = wbUser.Worksheets("Gass")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Set wsTarget = wbUser.Worksheets.Add
        wsTarget.Name = "Gass"
    Else
        wsTarget.Cells.Clear
    End If

    rad = WorksheetFunction.PI() / 180
    matrixCount = 0
    
    ' Count Ds (Column G) to find the exact number of blocks
    Do While wsSource.Range("G" & (2 + matrixCount)).Value <> ""
        matrixCount = matrixCount + 1
    Loop

    If matrixCount = 0 Then Exit Sub

    totalRows = 3 * matrixCount
    totalCols = 3 * (matrixCount + 1)
    ReDim fullMatrix(1 To totalRows, 1 To totalCols)

    i = 1
    Do While i <= matrixCount
        If IsNumeric(wsSource.Range("H" & (i + 1)).Value) Then theta_i_1 = wsSource.Range("H" & (i + 1)).Value Else theta_i_1 = 0
        If IsNumeric(wsSource.Range("H" & (i + 2)).Value) Then theta_i = wsSource.Range("H" & (i + 2)).Value Else theta_i = 0
        If IsNumeric(wsSource.Range("G" & (i + 1)).Value) Then delta_s = wsSource.Range("G" & (i + 1)).Value Else delta_s = 0

        theta_i_1 = theta_i_1 * rad
        theta_i = theta_i * rad

        Dim startRow As Long, startCol As Long
        startRow = (i - 1) * 3 + 1
        startCol = (i - 1) * 3 + 1

        fullMatrix(startRow, startCol) = -Cos(theta_i_1)
        fullMatrix(startRow, startCol + 1) = -Sin(theta_i_1)
        If startCol + 3 <= totalCols Then fullMatrix(startRow, startCol + 3) = Cos(theta_i)
        If startCol + 4 <= totalCols Then fullMatrix(startRow, startCol + 4) = Sin(theta_i)

        fullMatrix(startRow + 1, startCol) = -Sin(theta_i_1)
        fullMatrix(startRow + 1, startCol + 1) = Cos(theta_i_1)
        If startCol + 3 <= totalCols Then fullMatrix(startRow + 1, startCol + 3) = Sin(theta_i)
        If startCol + 4 <= totalCols Then fullMatrix(startRow + 1, startCol + 4) = -Cos(theta_i)

        If startCol + 1 <= totalCols Then fullMatrix(startRow + 2, startCol + 1) = -delta_s / 2
        If startCol + 2 <= totalCols Then fullMatrix(startRow + 2, startCol + 2) = -1
        If startCol + 4 <= totalCols Then fullMatrix(startRow + 2, startCol + 4) = -delta_s / 2
        If startCol + 5 <= totalCols Then fullMatrix(startRow + 2, startCol + 5) = 1

        For r = startRow To startRow + 2
            For c = startCol To startCol + 6
                If c <= totalCols Then
                    val = fullMatrix(r, c)
                    If Abs(val) < 0.0000000001 Then val = 0
                    fullMatrix(r, c) = val
                End If
            Next c
        Next r
        i = i + 1
    Loop
    wsTarget.Range("A1").Resize(totalRows, totalCols).Value = fullMatrix
End Sub
