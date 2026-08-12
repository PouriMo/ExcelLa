Attribute VB_Name = "Module7"
Option Explicit

Sub PlotThrustLine(Optional wsDataIn As Worksheet)
    Dim wsData As Worksheet, wsForces As Worksheet, wsPlot As Worksheet
    Dim wbUser As Workbook
    Dim i As Long, j As Long, NI As Long, NB As Long, lam As Double

    If Not wsDataIn Is Nothing Then
        Set wsData = wsDataIn
        Set wbUser = wsDataIn.Parent
    Else
        Set wbUser = ActiveWorkbook
        Set wsData = wbUser.ActiveSheet
    End If

    On Error Resume Next
    Set wsForces = wbUser.Sheets("ForcesTable")
    On Error GoTo ErrorHandler
    If wsForces Is Nothing Then MsgBox "Run analysis first!", vbCritical: Exit Sub

    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant
    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    If NB = 0 Then
        MsgBox "No block geometry found on sheet '" & wsData.Name & "'.", vbExclamation
        Exit Sub
    End If
    NI = NB + 1

    Dim tArr() As Double: tArr = GetInterfaceThickness(wsData, NI)

    Dim Nvec() As Double, Mvec() As Double
    ReDim Nvec(1 To NI): ReDim Mvec(1 To NI)
    For i = 1 To NI
        Nvec(i) = wsForces.Cells(i + 1, 2).Value
        Mvec(i) = wsForces.Cells(i + 1, 4).Value
    Next i
    lam = wsForces.Range("G2").Value

    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    ReDim p1x(1 To NI), p1y(1 To NI), p2x(1 To NI), p2y(1 To NI)

    Dim FiArr() As Double: ReDim FiArr(1 To NI)
    For i = 1 To NI: FiArr(i) = wsData.Cells(i + 1, 8).Value: Next i

    For i = 1 To NI
        Dim curX() As Double, curY() As Double
        Dim prevX() As Double, prevY() As Double

        If i = 1 Then
            curX = bx(1): curY = by(1)
            FindEdgeByNormal curX, curY, FiArr(1), p1x(1), p1y(1), p2x(1), p2y(1)
        ElseIf i = NI Then
            curX = bx(NB): curY = by(NB)
            FindEdgeByNormal curX, curY, FiArr(NI), p1x(NI), p1y(NI), p2x(NI), p2y(NI)
        Else
            prevX = bx(i - 1): prevY = by(i - 1)
            curX = bx(i): curY = by(i)
            FindSharedEdge prevX, prevY, curX, curY, p1x(i), p1y(i), p2x(i), p2y(i)
        End If
        OrientEdgeByFi p1x(i), p1y(i), p2x(i), p2y(i), FiArr(i)
    Next i

    Dim thrX() As Double, thrY() As Double
    ReDim thrX(1 To NI): ReDim thrY(1 To NI)
    Dim eVec() As Double: ReDim eVec(1 To NI)

    Dim Nref As Double: Nref = 0
    For i = 1 To NI
        If Abs(Nvec(i)) > Nref Then Nref = Abs(Nvec(i))
    Next i
    If Nref = 0 Then Nref = 1

    For i = 1 To NI
        eVec(i) = 0
        If Abs(Nvec(i)) > 0.000001 * Nref Then eVec(i) = -1 * Mvec(i) / Nvec(i)

        Dim dx As Double, dy As Double, L As Double
        dx = p2x(i) - p1x(i): dy = p2y(i) - p1y(i)
        L = Sqr(dx * dx + dy * dy)
        If L > 0 Then dx = dx / L: dy = dy / L
        Dim midx As Double, midy As Double, e As Double
        midx = (p1x(i) + p2x(i)) / 2: midy = (p1y(i) + p2y(i)) / 2
        e = eVec(i)
        If e > tArr(i) / 2 Then e = tArr(i) / 2
        If e < -tArr(i) / 2 Then e = -tArr(i) / 2
        thrX(i) = midx + e * dx: thrY(i) = midy + e * dy
    Next i

    Dim rawH() As Long, nRaw As Long, finalH() As Long, nHinges As Long
    Call FindHingeCandidates(eVec, tArr, NI, rawH, nRaw)
    Call ReduceToFourHinges(rawH, nRaw, eVec, finalH, nHinges)

    Application.DisplayAlerts = False
    On Error Resume Next: wbUser.Sheets("ThrustPlot").Delete: On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    Set wsPlot = wbUser.Sheets.Add: wsPlot.Name = "ThrustPlot"

    ' ---------- COLUMN LAYOUT (all data left of column N) ----------
    '   A=BlockX     B=BlockY
    '   C=ThrustX    D=ThrustY
    '   E=HingeX     F=HingeY
    '   G=BackX      H=BackY
    '   I=FStartX    J=FStartY
    '   K=FEndX      L=FEndY
    '   (chart placed at Left:=800 -> roughly column O onward)
    ' ----------------------------------------------------------------

    ' ---- Blocks (cols A,B) ----
    wsPlot.Cells(1, 1).Value = "BlockX": wsPlot.Cells(1, 2).Value = "BlockY"
    Dim row As Long: row = 2
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            Dim ax() As Double, ay() As Double
            ax = bx(i): ay = by(i)
            For j = 1 To UBound(ax)
                wsPlot.Cells(row, 1).Value = ax(j): wsPlot.Cells(row, 2).Value = ay(j): row = row + 1
            Next j
            wsPlot.Cells(row, 1).Value = ax(1): wsPlot.Cells(row, 2).Value = ay(1): row = row + 1
            wsPlot.Cells(row, 1).Value = CVErr(xlErrNA): wsPlot.Cells(row, 2).Value = CVErr(xlErrNA): row = row + 1
        End If
    Next i
    Dim blockRows As Long: blockRows = row - 1

    ' ---- Thrust (cols C,D) and Hinges (cols E,F) ----
    wsPlot.Cells(1, 3).Value = "ThrustX": wsPlot.Cells(1, 4).Value = "ThrustY"
    For i = 1 To NI
        wsPlot.Cells(i + 1, 3).Value = thrX(i): wsPlot.Cells(i + 1, 4).Value = thrY(i)
    Next i
    wsPlot.Cells(1, 5).Value = "HingeX": wsPlot.Cells(1, 6).Value = "HingeY"
    For i = 1 To nHinges
        If finalH(i) > 0 And finalH(i) <= NI Then
            wsPlot.Cells(i + 1, 5).Value = thrX(finalH(i)): wsPlot.Cells(i + 1, 6).Value = thrY(finalH(i))
        End If
    Next i

    ' ---- Backfill (cols G,H) ----
    wsPlot.Cells(1, 7).Value = "BackX": wsPlot.Cells(1, 8).Value = "BackY"
    Dim backRow As Long: backRow = 2
    For i = 1 To NB
        If Not fx(i) Is Nothing Then
            Dim k As Long
            For k = 1 To fx(i).Count
                Dim bfx() As Double, bfy() As Double
                bfx = fx(i).Item(k): bfy = fy(i).Item(k)
                For j = 1 To UBound(bfx)
                    wsPlot.Cells(backRow, 7).Value = bfx(j): wsPlot.Cells(backRow, 8).Value = bfy(j): backRow = backRow + 1
                Next j
                wsPlot.Cells(backRow, 7).Value = bfx(1): wsPlot.Cells(backRow, 8).Value = bfy(1): backRow = backRow + 1
                wsPlot.Cells(backRow, 7).Value = CVErr(xlErrNA): wsPlot.Cells(backRow, 8).Value = CVErr(xlErrNA): backRow = backRow + 1
            Next k
        End If
    Next i
    Dim backRows As Long: backRows = backRow - 1

    ' ---- Force arrows (cols I,J = start; K,L = end) ----
    Dim FxArr() As Double, FyArr() As Double
    Call GetBlockForces(wsData, NB, FxArr, FyArr)

    Dim xMinAll As Double, xMaxAll As Double, yMinAll As Double, yMaxAll As Double
    Dim gotBound As Boolean: gotBound = False
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            ax = bx(i): ay = by(i)
            For j = 1 To UBound(ax)
                If Not gotBound Then
                    xMinAll = ax(j): xMaxAll = ax(j): yMinAll = ay(j): yMaxAll = ay(j): gotBound = True
                Else
                    If ax(j) < xMinAll Then xMinAll = ax(j)
                    If ax(j) > xMaxAll Then xMaxAll = ax(j)
                    If ay(j) < yMinAll Then yMinAll = ay(j)
                    If ay(j) > yMaxAll Then yMaxAll = ay(j)
                End If
            Next j
        End If
    Next i
    Dim arrowLen As Double
    arrowLen = 0.12 * (yMaxAll - yMinAll)
    If arrowLen <= 0 Then arrowLen = 0.12 * (xMaxAll - xMinAll)
    If arrowLen <= 0 Then arrowLen = 10

    ' FIX (border pushes/clips backfill): the axis scale set below only
    ' auto-fits to the "Arch" and "Thrust line" chart series - backfill is
    ' never a series (it's drawn manually by DrawBlocksOnChart using
    ' whatever axis scale happens to already exist). If backfill extends
    ' beyond the blocks' bounding box - which it normally does - its
    ' points get mapped outside the plot area and clipped/distorted at the
    ' chart's edge. So fold backfill extents into xMinAll/xMaxAll/yMinAll/
    ' yMaxAll here too, BEFORE the axis scale is set, so the chart is
    ' guaranteed to be sized to fit everything it will actually draw.
    For i = 1 To NB
        If Not fx(i) Is Nothing Then
            Dim kk As Long
            For kk = 1 To fx(i).Count
                Dim bfxB() As Double, bfyB() As Double
                bfxB = fx(i).Item(kk): bfyB = fy(i).Item(kk)
                For j = 1 To UBound(bfxB)
                    If bfxB(j) < xMinAll Then xMinAll = bfxB(j)
                    If bfxB(j) > xMaxAll Then xMaxAll = bfxB(j)
                    If bfyB(j) < yMinAll Then yMinAll = bfyB(j)
                    If bfyB(j) > yMaxAll Then yMaxAll = bfyB(j)
                Next j
            Next kk
        End If
    Next i

    wsPlot.Cells(1, 9).Value = "FStartX": wsPlot.Cells(1, 10).Value = "FStartY"
    wsPlot.Cells(1, 11).Value = "FEndX": wsPlot.Cells(1, 12).Value = "FEndY"
    Dim frow As Long: frow = 2
    For i = 1 To NB
        If Abs(FxArr(i)) > 0.0000001 Or Abs(FyArr(i)) > 0.0000001 Then
            ax = bx(i): ay = by(i)
            Dim cxA As Double, cyA As Double, npts As Long
            npts = UBound(ax)

            Dim signedArea As Double, cross As Double, x0 As Double, y0 As Double, x1 As Double, y1 As Double
            signedArea = 0: cxA = 0: cyA = 0
            For j = 1 To npts
                x0 = ax(j): y0 = ay(j)
                If j = npts Then
                    x1 = ax(1): y1 = ay(1)
                Else
                    x1 = ax(j + 1): y1 = ay(j + 1)
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
                    cxA = cxA + ax(j): cyA = cyA + ay(j)
                Next j
                cxA = cxA / npts: cyA = cyA / npts
            End If
            Dim fmag As Double: fmag = Sqr(FxArr(i) ^ 2 + FyArr(i) ^ 2)
            Dim ux As Double, uy As Double
            If fmag > 0 Then ux = FxArr(i) / fmag: uy = FyArr(i) / fmag Else ux = 0: uy = 0
            wsPlot.Cells(frow, 9).Value = cxA
            wsPlot.Cells(frow, 10).Value = cyA
            wsPlot.Cells(frow, 11).Value = cxA + ux * arrowLen
            wsPlot.Cells(frow, 12).Value = cyA + uy * arrowLen

            ' Fold the arrow tip into the bounding box too (see FIX above)
            Dim tipX As Double, tipY As Double
            tipX = cxA + ux * arrowLen: tipY = cyA + uy * arrowLen
            If tipX < xMinAll Then xMinAll = tipX
            If tipX > xMaxAll Then xMaxAll = tipX
            If tipY < yMinAll Then yMinAll = tipY
            If tipY > yMaxAll Then yMaxAll = tipY

            frow = frow + 1
        End If
    Next i
    Dim forceRows As Long: forceRows = frow - 1

    ' ---- Create chart, placed at column ~O ----
    Dim ch As ChartObject: Set ch = wsPlot.ChartObjects.Add(Left:=800, Top:=10, Width:=750, Height:=420)
    With ch.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True: .ChartTitle.Text = "Thrust Line - lambda = " & Format$(lam, "0.00") & " [N]"
        .ChartTitle.Font.Size = 16: .ChartTitle.Font.Bold = True

        ' Chart background WHITE (visible), plot area transparent, no
        ' borders. White background is essential for GIF export to render
        ' properly on the userform's image control.
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse

        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop

        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True: .AxisTitle.Text = "x [mm]": .AxisTitle.Font.Size = 12
            .TickLabels.Font.Size = 11
            .HasMajorGridlines = True
            With .MajorGridlines.Format.Line
                .Visible = msoTrue: .ForeColor.RGB = RGB(190, 190, 190): .Weight = 0.75: .DashStyle = msoLineDash
            End With
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True: .AxisTitle.Text = "y [mm]": .AxisTitle.Font.Size = 12
            .TickLabels.Font.Size = 11
            .HasMajorGridlines = True
            With .MajorGridlines.Format.Line
                .Visible = msoTrue: .ForeColor.RGB = RGB(190, 190, 190): .Weight = 0.75: .DashStyle = msoLineDash
            End With
        End With
        .Legend.Font.Size = 11

        With .SeriesCollection.NewSeries
            .Name = "Arch": .XValues = wsPlot.Range("A2:A" & blockRows): .Values = wsPlot.Range("B2:B" & blockRows)
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0): .MarkerStyle = xlMarkerStyleNone
        End With
        With .SeriesCollection.NewSeries
            .Name = "Thrust line": .XValues = wsPlot.Range("C2:C" & NI + 1): .Values = wsPlot.Range("D2:D" & NI + 1)
            .Format.Line.ForeColor.RGB = RGB(220, 20, 60): .Format.Line.Weight = 2: .MarkerStyle = xlMarkerStylePlus: .MarkerSize = 7: .MarkerForegroundColor = RGB(220, 20, 60)
        End With
    End With

    ' ---- Padding: extend by 1 whole major-unit each side, integer ticks ----
    On Error Resume Next
    With ch.Chart
        .Axes(xlCategory).MinimumScaleIsAuto = True
        .Axes(xlCategory).MaximumScaleIsAuto = True
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
        DoEvents
        Dim majX As Double, majY As Double
        majX = .Axes(xlCategory).MajorUnit
        majY = .Axes(xlValue).MajorUnit
        .Axes(xlCategory).MinimumScale = .Axes(xlCategory).MinimumScale - majX
        .Axes(xlCategory).MaximumScale = .Axes(xlCategory).MaximumScale + majX
        .Axes(xlValue).MinimumScale = .Axes(xlValue).MinimumScale - majY
        .Axes(xlValue).MaximumScale = .Axes(xlValue).MaximumScale + majY

        ' FIX (border pushes/clips backfill): Excel's autoscale above only
        ' considered the "Arch" and "Thrust line" series - it knows nothing
        ' about backfill or force-arrow geometry (those are hand-drawn
        ' shapes, not series). Widen the scale further, if needed, so the
        ' FULL bounding box computed earlier (xMinAll/xMaxAll/yMinAll/
        ' yMaxAll, which DOES include backfill + arrows) is guaranteed to
        ' fit with the same one-major-unit margin. This is what
        ' DrawBlocksOnChart's pixel mapping relies on to avoid clipping
        ' backfill against the chart border.
        Dim padX As Double, padY As Double
        padX = 0.1 * (xMaxAll - xMinAll): If padX <= 0 Then padX = majX
        padY = 0.1 * (yMaxAll - yMinAll): If padY <= 0 Then padY = majY
        If xMinAll - padX < .Axes(xlCategory).MinimumScale Then .Axes(xlCategory).MinimumScale = xMinAll - padX
        If xMaxAll + padX > .Axes(xlCategory).MaximumScale Then .Axes(xlCategory).MaximumScale = xMaxAll + padX
        If yMinAll - padY < .Axes(xlValue).MinimumScale Then .Axes(xlValue).MinimumScale = yMinAll - padY
        If yMaxAll + padY > .Axes(xlValue).MaximumScale Then .Axes(xlValue).MaximumScale = yMaxAll + padY

        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlValue).TickLabels.NumberFormat = "0"
    End With
    DoEvents
    On Error GoTo ErrorHandler

    Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 1, 2, 2, blockRows, RGB(242, 136, 89), RGB(0, 0, 0), 1, False)
    If backRows > 1 Then Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 7, 8, 2, backRows, RGB(0, 0, 0), RGB(100, 100, 100), 1, True)
    If forceRows > 1 Then Call DrawForceArrows(ch.Chart.Parent, wsPlot, 9, 10, 11, 12, 2, forceRows, RGB(0, 112, 192))
    Call DrawHingesOnChart(ch.Chart.Parent, wsPlot, 5, 6, 2, nHinges + 1, RGB(220, 20, 60))
    Exit Sub

ErrorHandler:
    MsgBox "Error in PlotThrustLine: " & Err.Description, vbCritical
End Sub
