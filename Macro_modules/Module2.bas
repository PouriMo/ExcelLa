Attribute VB_Name = "Module2"
Option Explicit

Sub GenerateGassMatrix()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim wbUser As Workbook
    Dim i As Long, matrixCount As Long
    Dim totalRows As Long, totalCols As Long
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

    ' Extract block coordinates using the existing geometry reader
    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant
    Dim NB As Long
    Call ReadCSVGeometry(wsSource, bx, by, fx, fy, NB)
    
    If NB = 0 Then Exit Sub
    matrixCount = NB
    
    totalRows = 3 * matrixCount
    totalCols = 3 * (matrixCount + 1)
    ReDim fullMatrix(1 To totalRows, 1 To totalCols)

    ' Read interface angles (Fi) from column H of the CSV data
    Dim FiArr() As Double
    ReDim FiArr(1 To NB + 1)
    For i = 1 To NB + 1
        If IsNumeric(wsSource.Cells(i + 1, 8).Value) And wsSource.Cells(i + 1, 8).Value <> "" Then
            FiArr(i) = wsSource.Cells(i + 1, 8).Value
        Else
            FiArr(i) = 0
        End If
    Next i

    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    ReDim p1x(1 To NB + 1), p1y(1 To NB + 1), p2x(1 To NB + 1), p2y(1 To NB + 1)
    
    Dim curX() As Double, curY() As Double
    Dim prevX() As Double, prevY() As Double

    ' Locate the exact geometric endpoints of each interface
    For i = 1 To NB + 1
        If i = 1 Then
            curX = bx(1): curY = by(1)
            Call FindEdgeByNormal(curX, curY, FiArr(1), p1x(1), p1y(1), p2x(1), p2y(1))
        ElseIf i = NB + 1 Then
            curX = bx(NB): curY = by(NB)
            Call FindEdgeByNormal(curX, curY, FiArr(NB + 1), p1x(NB + 1), p1y(NB + 1), p2x(NB + 1), p2y(NB + 1))
        Else
            prevX = bx(i - 1): prevY = by(i - 1)
            curX = bx(i): curY = by(i)
            Call FindSharedEdge(prevX, prevY, curX, curY, p1x(i), p1y(i), p2x(i), p2y(i))
        End If
        Call OrientEdgeByFi(p1x(i), p1y(i), p2x(i), p2y(i), FiArr(i))
    Next i

    Dim rad As Double
    rad = WorksheetFunction.PI() / 180
    
    For i = 1 To NB
        Dim theta_i_1 As Double, theta_i As Double
        theta_i_1 = FiArr(i) * rad
        theta_i = FiArr(i + 1) * rad
        
        Dim npts As Long, j As Long
        Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, cross As Double
        Dim signedArea As Double, cxA As Double, cyA As Double
        
        curX = bx(i): curY = by(i)
        npts = UBound(curX)
        signedArea = 0: cxA = 0: cyA = 0
        
        ' Compute the exact two-dimensional centroid of the structural block
        For j = 1 To npts
            x0 = curX(j): y0 = curY(j)
            If j = npts Then
                x1 = curX(1): y1 = curY(1)
            Else
                x1 = curX(j + 1): y1 = curY(j + 1)
            End If
            cross = x0 * y1 - x1 * y0
            signedArea = signedArea + cross
            cxA = cxA + (x0 + x1) * cross
            cyA = cyA + (y0 + y1) * cross
        Next j
        signedArea = signedArea / 2
        If Abs(signedArea) > 0.000000001 Then
            cxA = cxA / (6 * signedArea)
            cyA = cyA / (6 * signedArea)
        Else
            cxA = 0: cyA = 0
            For j = 1 To npts
                cxA = cxA + curX(j): cyA = cyA + curY(j)
            Next j
            cxA = cxA / npts: cyA = cyA / npts
        End If
        
        ' Calculate the exact midpoints of the contiguous interfaces
        Dim midX_i_1 As Double, midY_i_1 As Double
        Dim midX_i As Double, midY_i As Double
        
        midX_i_1 = (p1x(i) + p2x(i)) / 2
        midY_i_1 = (p1y(i) + p2y(i)) / 2
        
        midX_i = (p1x(i + 1) + p2x(i + 1)) / 2
        midY_i = (p1y(i + 1) + p2y(i + 1)) / 2
        
        ' Compute geometric offset vectors relative to the centroid
        Dim dx_k_1 As Double, dy_k_1 As Double
        Dim dx_k As Double, dy_k As Double
        
        dx_k_1 = midX_i_1 - cxA
        dy_k_1 = midY_i_1 - cyA
        
        dx_k = midX_i - cxA
        dy_k = midY_i - cyA
        
        ' Derive the EXACT moment lever arms via 2D cross product of the force components
        Dim m_N_i_1 As Double, m_V_i_1 As Double
        Dim m_N_i As Double, m_V_i As Double
        
        m_N_i_1 = dx_k_1 * (-Sin(theta_i_1)) - dy_k_1 * (-Cos(theta_i_1))
        m_V_i_1 = dx_k_1 * Cos(theta_i_1) - dy_k_1 * (-Sin(theta_i_1))
        m_N_i = dx_k * Sin(theta_i) - dy_k * Cos(theta_i)
        m_V_i = dx_k * (-Cos(theta_i)) - dy_k * Sin(theta_i)

        Dim startRow As Long, startCol As Long
        startRow = (i - 1) * 3 + 1
        startCol = (i - 1) * 3 + 1

        ' Horizontal equilibrium
        fullMatrix(startRow, startCol) = -Cos(theta_i_1)
        fullMatrix(startRow, startCol + 1) = -Sin(theta_i_1)
        If startCol + 3 <= totalCols Then fullMatrix(startRow, startCol + 3) = Cos(theta_i)
        If startCol + 4 <= totalCols Then fullMatrix(startRow, startCol + 4) = Sin(theta_i)

        ' Vertical equilibrium
        fullMatrix(startRow + 1, startCol) = -Sin(theta_i_1)
        fullMatrix(startRow + 1, startCol + 1) = Cos(theta_i_1)
        If startCol + 3 <= totalCols Then fullMatrix(startRow + 1, startCol + 3) = Sin(theta_i)
        If startCol + 4 <= totalCols Then fullMatrix(startRow + 1, startCol + 4) = -Cos(theta_i)

        ' Rotational equilibrium
        fullMatrix(startRow + 2, startCol) = m_N_i_1           ' N_{i-1}
        fullMatrix(startRow + 2, startCol + 1) = m_V_i_1       ' V_{i-1}
        fullMatrix(startRow + 2, startCol + 2) = -1            ' M_{i-1}
        
        If startCol + 3 <= totalCols Then fullMatrix(startRow + 2, startCol + 3) = m_N_i  ' N_i
        If startCol + 4 <= totalCols Then fullMatrix(startRow + 2, startCol + 4) = m_V_i  ' V_i
        If startCol + 5 <= totalCols Then fullMatrix(startRow + 2, startCol + 5) = 1      ' M_i

        ' Clean near-zero floating point artifacts
        For r = startRow To startRow + 2
            For c = startCol To startCol + 6
                If c <= totalCols Then
                    val = fullMatrix(r, c)
                    If Abs(val) < 0.0000000001 Then val = 0
                    fullMatrix(r, c) = val
                End If
            Next c
        Next r
    Next i
    
    wsTarget.Range("A1").Resize(totalRows, totalCols).Value = fullMatrix
End Sub
