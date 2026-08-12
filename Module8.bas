Attribute VB_Name = "Module8"
Option Explicit

Sub PlotDeformedShape(Optional wsDataIn As Worksheet)
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
        MsgBox "No block geometry found on sheet '" & wsData.Name & "'." & vbCrLf & _
               "Make sure your CSV data sheet is selected/active before plotting.", vbExclamation
        Exit Sub
    End If
    NI = NB + 1

    Dim tArr() As Double: tArr = GetInterfaceThickness(wsData, NI)
    Dim Nvec() As Double, Mvec() As Double: ReDim Nvec(1 To NI): ReDim Mvec(1 To NI)
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

    Dim thrX() As Double, thrY() As Double: ReDim thrX(1 To NI): ReDim thrY(1 To NI)
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
        dx = p2x(i) - p1x(i): dy = p2y(i) - p1y(i): L = Sqr(dx * dx + dy * dy)
        If L > 0 Then dx = dx / L: dy = dy / L
        Dim e As Double: e = eVec(i)
        If e > tArr(i) / 2 Then e = tArr(i) / 2
        If e < -tArr(i) / 2 Then e = -tArr(i) / 2
        thrX(i) = (p1x(i) + p2x(i)) / 2 + e * dx: thrY(i) = (p1y(i) + p2y(i)) / 2 + e * dy
    Next i

    Dim rawH() As Long, nRaw As Long, hIfaces() As Long, nH As Long
    Call FindHingeCandidates(eVec, tArr, NI, rawH, nRaw)
    Call ReduceToFourHinges(rawH, nRaw, eVec, hIfaces, nH)

    Dim boundaries(1 To 4) As Long
    For i = 1 To nH: boundaries(i) = hIfaces(i): Next i
    If nH <> 4 Then Exit Sub

    Dim PLx As Double, PLy As Double, H1x As Double, H1y As Double, H2x As Double, H2y As Double, PRx As Double, PRy As Double
    PLx = thrX(boundaries(1)): PLy = thrY(boundaries(1))
    H1x = thrX(boundaries(2)): H1y = thrY(boundaries(2))
    H2x = thrX(boundaries(3)): H2y = thrY(boundaries(3))
    PRx = thrX(boundaries(4)): PRy = thrY(boundaries(4))

    Dim segOf() As Long: ReDim segOf(1 To NB)
    Dim segTmp As Long: segTmp = 0
    For i = 1 To NB
        If i = boundaries(1) Then segTmp = 1
        If i = boundaries(2) Then segTmp = 2
        If i = boundaries(3) Then segTmp = 3
        If i = boundaries(4) Then segTmp = 4
        segOf(i) = segTmp
    Next i

    Dim chains(0 To 4) As String, seg As Long: seg = 0
    For i = 1 To NB
        If i = boundaries(1) Then seg = 1
        If i = boundaries(2) Then seg = 2
        If i = boundaries(3) Then seg = 3
        If i = boundaries(4) Then seg = 4
        chains(seg) = chains(seg) & i & ","
    Next i

    Dim theta1 As Double, H1nx As Double, H1ny As Double, H2nx As Double, H2ny As Double
    Dim maxDeg As Double: maxDeg = 20#
    Dim degStep As Double
    If NB > 30 Then
        degStep = 0.1
    Else
        degStep = 0.5
    End If
    Dim nDegSteps As Long: nDegSteps = CLng(maxDeg / degStep)
    Dim Rmid(1 To 2, 1 To 2) As Double, TmX As Double, TmY As Double, theta3 As Double
    Dim testSign As Variant, found As Boolean: found = False

    Dim blkSeg1Last As Long, blkSeg2First As Long, blkSeg2Last As Long, blkSeg3First As Long
    blkSeg1Last = boundaries(2) - 1: blkSeg2First = boundaries(2)
    blkSeg2Last = boundaries(3) - 1: blkSeg3First = boundaries(3)

    Dim signOrder As Variant
    If Sgn(eVec(boundaries(2))) >= 0 Then
        signOrder = Array(1, -1)
    Else
        signOrder = Array(-1, 1)
    End If

    Dim branchOrder As Variant: branchOrder = Array(0, 1, 2)
    Dim testBranch As Variant
    Dim pen As Double

    Dim bestPenetration As Double: bestPenetration = 1E+30
    Dim bestTheta1 As Double, bestTheta3 As Double
    Dim bestH1nx As Double, bestH1ny As Double, bestH2nx As Double, bestH2ny As Double
    Dim bestRmid(1 To 2, 1 To 2) As Double, bestTmX As Double, bestTmY As Double
    Dim haveBest As Boolean: haveBest = False

    For i = 0 To nDegSteps - 1
        For Each testSign In signOrder
            theta1 = testSign * (maxDeg - i * degStep) * 3.14159265 / 180
            For Each testBranch In branchOrder
                If FourBarSolve2(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, H1nx, H1ny, H2nx, H2ny, CLng(testBranch)) Then
                    Call RigidTransform2(H1x, H1y, H2x, H2y, H1nx, H1ny, H2nx, H2ny, Rmid, TmX, TmY)
                    theta3 = Atn2b(H2ny - PRy, H2nx - PRx) - Atn2b(H2y - PRy, H2x - PRx)

                    pen = MaxPenetrationFull(bx, by, fx, fy, NB, segOf, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)

                    If pen <= 0 Then
                        found = True: Exit For
                    ElseIf pen < bestPenetration Then
                        bestPenetration = pen
                        bestTheta1 = theta1: bestTheta3 = theta3
                        bestH1nx = H1nx: bestH1ny = H1ny: bestH2nx = H2nx: bestH2ny = H2ny
                        bestRmid(1, 1) = Rmid(1, 1): bestRmid(1, 2) = Rmid(1, 2)
                        bestRmid(2, 1) = Rmid(2, 1): bestRmid(2, 2) = Rmid(2, 2)
                        bestTmX = TmX: bestTmY = TmY
                        haveBest = True
                    End If
                End If
            Next testBranch
            If found Then Exit For
        Next testSign
        If found Then Exit For
    Next i

    If Not found And haveBest Then
        theta1 = bestTheta1: theta3 = bestTheta3
        H1nx = bestH1nx: H1ny = bestH1ny: H2nx = bestH2nx: H2ny = bestH2ny
        Rmid(1, 1) = bestRmid(1, 1): Rmid(1, 2) = bestRmid(1, 2)
        Rmid(2, 1) = bestRmid(2, 1): Rmid(2, 2) = bestRmid(2, 2)
        TmX = bestTmX: TmY = bestTmY
        MsgBox "Note: this geometry does not allow a perfectly clean rigid-body " & _
               "collapse mechanism within the tested crank-angle range. The closest " & _
               "non-interpenetrating configuration found still has a small residual " & _
               "overlap (max penetration ~" & Format(bestPenetration, "0.00") & " mm) " & _
               "near one of the hinges - shown below.", vbInformation
    ElseIf Not found And Not haveBest Then
        theta1 = 1 * 8# * 0.01 * 3.14159265 / 180
        Call FourBarSolve2(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy, theta1, H1nx, H1ny, H2nx, H2ny)
        Call RigidTransform2(H1x, H1y, H2x, H2y, H1nx, H1ny, H2nx, H2ny, Rmid, TmX, TmY)
        theta3 = Atn2b(H2ny - PRy, H2nx - PRx) - Atn2b(H2y - PRy, H2x - PRx)
    End If

    Debug.Print "found=" & found & " haveBest=" & haveBest & " bestPenetration=" & bestPenetration

    Application.DisplayAlerts = False
    On Error Resume Next: wbUser.Sheets("DeformedPlot").Delete: On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    Set wsPlot = wbUser.Sheets.Add: wsPlot.Name = "DeformedPlot"

    wsPlot.Cells(1, 1).Value = "OrigX": wsPlot.Cells(1, 2).Value = "OrigY"
    Dim origRow As Long: origRow = 2

    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            Dim oax() As Double, oay() As Double
            oax = bx(i): oay = by(i)
            For j = 1 To UBound(oax)
                wsPlot.Cells(origRow, 1).Value = oax(j): wsPlot.Cells(origRow, 2).Value = oay(j): origRow = origRow + 1
            Next j
            wsPlot.Cells(origRow, 1).Value = oax(1): wsPlot.Cells(origRow, 2).Value = oay(1): origRow = origRow + 1
            wsPlot.Cells(origRow, 1).Value = CVErr(xlErrNA): wsPlot.Cells(origRow, 2).Value = CVErr(xlErrNA): origRow = origRow + 1
        End If
    Next i
    Dim origRows As Long: origRows = origRow - 1

    wsPlot.Cells(1, 3).Value = "DefX": wsPlot.Cells(1, 4).Value = "DefY"
    ' Moved DefBackX/DefBackY from cols 11-12 to cols 7-8 (right after
    ' the hinge cols 5-6), so they sit in the left "data area" of the
    ' sheet instead of under the chart. The chart is placed starting
    ' at Left:=300 which is roughly column K, so cols 7-8 are safely
    ' clear of it.
    wsPlot.Cells(1, 7).Value = "DefBackX": wsPlot.Cells(1, 8).Value = "DefBackY"

    Dim defRow As Long: defRow = 2
    Dim defBackRow As Long: defBackRow = 2
    Dim segIdx As Long, bIdx As Long, bnum As Long, parts() As String
    Dim npx As Double, npy As Double, px As Double, py As Double

    Dim blockObstacles As Collection: Set blockObstacles = New Collection
    For i = 1 To NB
        If Not IsEmpty(bx(i)) Then
            blockObstacles.Add TransformPolyBySeg(bx(i), by(i), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        End If
    Next i
    Dim backfillObstacles As Collection: Set backfillObstacles = New Collection

    For segIdx = 0 To 4
        If Len(chains(segIdx)) > 0 Then
            parts = Split(Left(chains(segIdx), Len(chains(segIdx)) - 1), ",")
            For bIdx = 0 To UBound(parts)
                bnum = CLng(parts(bIdx))

                If Not IsEmpty(bx(bnum)) Then
                    Dim ax() As Double, ay() As Double
                    ax = bx(bnum): ay = by(bnum)
                    For j = 1 To UBound(ax)
                        px = ax(j): py = ay(j)
                        TransformPoint px, py, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, npx, npy
                        wsPlot.Cells(defRow, 3).Value = npx: wsPlot.Cells(defRow, 4).Value = npy: defRow = defRow + 1
                    Next j
                    px = ax(1): py = ay(1)
                    TransformPoint px, py, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY, npx, npy
                    wsPlot.Cells(defRow, 3).Value = npx: wsPlot.Cells(defRow, 4).Value = npy: defRow = defRow + 1
                    wsPlot.Cells(defRow, 3).Value = CVErr(xlErrNA): wsPlot.Cells(defRow, 4).Value = CVErr(xlErrNA): defRow = defRow + 1
                End If

                If Not fx(bnum) Is Nothing Then
                    Dim k As Long
                    For k = 1 To fx(bnum).Count
                        Dim bfx() As Double, bfy() As Double
                        bfx = fx(bnum).Item(k): bfy = fy(bnum).Item(k)
                        Dim bfPoly() As Double, bfNudged() As Double
                        bfPoly = TransformPolyBySeg(bfx, bfy, segIdx, PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
                        Dim clearAgainst As Collection: Set clearAgainst = New Collection
                        Dim ob As Variant
                        For Each ob In blockObstacles: clearAgainst.Add ob: Next ob
                        For Each ob In backfillObstacles: clearAgainst.Add ob: Next ob
                        bfNudged = NudgeClearByRotation(bfPoly, clearAgainst)
                        backfillObstacles.Add bfNudged

                        Dim bn As Long: bn = UBound(bfNudged, 1)
                                                For j = 1 To bn
                            wsPlot.Cells(defBackRow, 7).Value = bfNudged(j, 1): wsPlot.Cells(defBackRow, 8).Value = bfNudged(j, 2): defBackRow = defBackRow + 1
                        Next j
                        wsPlot.Cells(defBackRow, 7).Value = bfNudged(1, 1): wsPlot.Cells(defBackRow, 8).Value = bfNudged(1, 2): defBackRow = defBackRow + 1
                        wsPlot.Cells(defBackRow, 7).Value = CVErr(xlErrNA): wsPlot.Cells(defBackRow, 8).Value = CVErr(xlErrNA): defBackRow = defBackRow + 1
                    Next k
                End If
            Next bIdx
        End If
    Next segIdx

    wsPlot.Cells(2, 5).Value = PLx: wsPlot.Cells(2, 6).Value = PLy
    wsPlot.Cells(3, 5).Value = H1nx: wsPlot.Cells(3, 6).Value = H1ny
    wsPlot.Cells(4, 5).Value = H2nx: wsPlot.Cells(4, 6).Value = H2ny
    wsPlot.Cells(5, 5).Value = PRx: wsPlot.Cells(5, 6).Value = PRy

        ' Chart moved further right (Left:=550) so columns A..J are fully
    ' available for the analysis data without being obscured.
        ' Chart is larger AND moved further right so it never overlaps the
    ' data columns (which now include BackX/BackY at H..I and force
    ' arrow data at O..R after Fix 3 below).
        Dim ch As ChartObject: Set ch = wsPlot.ChartObjects.Add(Left:=800, Top:=10, Width:=750, Height:=420)
    With ch.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True: .ChartTitle.Text = "Collapse Mechanism - lambda = " & Format$(lam, "0.00") & " [N]"
        .ChartTitle.Font.Size = 16: .ChartTitle.Font.Bold = True

        ' Chart background WHITE (visible), plot area transparent, no
        ' borders. White background is essential so exported GIF renders
        ' properly on the userform image control (else it's fully
        ' transparent and appears blank).
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
            .Name = "Original": .XValues = wsPlot.Range("A2:A" & origRows): .Values = wsPlot.Range("B2:B" & origRows)
            .Format.Line.ForeColor.RGB = RGB(180, 180, 180): .Format.Line.Weight = 1: .MarkerStyle = xlMarkerStyleNone
        End With

        With .SeriesCollection.NewSeries
            .Name = "Deformed": .XValues = wsPlot.Range("C2:C" & defRow - 1): .Values = wsPlot.Range("D2:D" & defRow - 1)
            .Format.Line.ForeColor.RGB = RGB(220, 20, 60): .Format.Line.Weight = 2: .MarkerStyle = xlMarkerStyleNone
        End With
        
    End With

            ' =====================================================================
    '  PADDING BLOCK v3: keep round tick numbers, add margin by extending
    '  axis range by a whole number of major units, and DO NOT resync
    '  MajorUnit (that's what was giving tick labels 5+ decimals and
    '  causing them to overlap).
    ' =====================================================================
    On Error Resume Next
    With ch.Chart
        Dim majX As Double, majY As Double
        Dim curXMin As Double, curXMax As Double, curYMin As Double, curYMax As Double

        ' Let Excel auto-fit first (do NOT set any axis manually yet)
        .Axes(xlCategory).MinimumScaleIsAuto = True
        .Axes(xlCategory).MaximumScaleIsAuto = True
        .Axes(xlValue).MinimumScaleIsAuto = True
        .Axes(xlValue).MaximumScaleIsAuto = True
        DoEvents

        majX = .Axes(xlCategory).MajorUnit
        majY = .Axes(xlValue).MajorUnit
        curXMin = .Axes(xlCategory).MinimumScale
        curXMax = .Axes(xlCategory).MaximumScale
        curYMin = .Axes(xlValue).MinimumScale
        curYMax = .Axes(xlValue).MaximumScale

        ' Extend by 1 major unit on all sides (round-number aligned)
        .Axes(xlCategory).MinimumScale = curXMin - majX
        .Axes(xlCategory).MaximumScale = curXMax + majX
        .Axes(xlValue).MinimumScale = curYMin - majY
        .Axes(xlValue).MaximumScale = curYMax + majY

        ' FIX (border pushes/clips backfill): DefBackX/DefBackY (cols G,H)
        ' is drawn manually below via DrawBlocksOnChart using the axis
        ' scale set here - but it was never part of either series
        ' ("Original"/"Deformed") that Excel's autoscale above actually
        ' looked at. If the (nudged) deformed backfill extends beyond the
        ' block outlines, its points map outside the plot area and get
        ' clipped/distorted at the chart's edge. Scan the backfill cells
        ' that were just written and widen the axis further if needed.
        Dim gMin As Double, gMax As Double, hMin As Double, hMax As Double
        Dim gotBB As Boolean: gotBB = False
        Dim rr As Long
        For rr = 2 To defBackRow - 1
            If Not IsError(wsPlot.Cells(rr, 7).Value) And Not IsEmpty(wsPlot.Cells(rr, 7).Value) Then
                Dim gv As Double, hv As Double
                gv = wsPlot.Cells(rr, 7).Value: hv = wsPlot.Cells(rr, 8).Value
                If Not gotBB Then
                    gMin = gv: gMax = gv: hMin = hv: hMax = hv: gotBB = True
                Else
                    If gv < gMin Then gMin = gv
                    If gv > gMax Then gMax = gv
                    If hv < hMin Then hMin = hv
                    If hv > hMax Then hMax = hv
                End If
            End If
        Next rr
        If gotBB Then
            Dim padX2 As Double, padY2 As Double
            padX2 = 0.1 * (gMax - gMin): If padX2 <= 0 Then padX2 = majX
            padY2 = 0.1 * (hMax - hMin): If padY2 <= 0 Then padY2 = majY
            If gMin - padX2 < .Axes(xlCategory).MinimumScale Then .Axes(xlCategory).MinimumScale = gMin - padX2
            If gMax + padX2 > .Axes(xlCategory).MaximumScale Then .Axes(xlCategory).MaximumScale = gMax + padX2
            If hMin - padY2 < .Axes(xlValue).MinimumScale Then .Axes(xlValue).MinimumScale = hMin - padY2
            If hMax + padY2 > .Axes(xlValue).MaximumScale Then .Axes(xlValue).MaximumScale = hMax + padY2
        End If

        ' Force integer tick labels only (no decimals) - the whole geometry
        ' is in mm, decimals are meaningless clutter.
        .Axes(xlCategory).TickLabels.NumberFormat = "0"
        .Axes(xlValue).TickLabels.NumberFormat = "0"
    End With
    DoEvents
    On Error GoTo ErrorHandler
    ' =====================================================================
    '  END PADDING BLOCK v3
    ' =====================================================================

    Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 3, 4, 2, defRow - 1, RGB(242, 136, 89), RGB(220, 20, 60), 2, False)
        If defBackRow > 2 Then Call DrawBlocksOnChart(ch.Chart.Parent, wsPlot, 7, 8, 2, defBackRow - 1, RGB(0, 0, 0), RGB(100, 100, 100), 2, True)
    Call DrawHingesOnChart(ch.Chart.Parent, wsPlot, 5, 6, 2, 5, RGB(0, 0, 0))
    Exit Sub

ErrorHandler:
    MsgBox "Error in PlotDeformedShape: " & Err.Description, vbCritical
End Sub

' ============================================================================
'  Helper functions for PlotDeformedShape
' ============================================================================

Private Function NudgeClearByRotation(poly() As Double, obstacles As Collection) As Double()
    Dim n As Long
    Dim pivX As Double
    Dim pivY As Double
    Dim i As Long

    n = UBound(poly, 1)

    pivX = poly(1, 1)
    pivY = poly(1, 2)

    For i = 2 To n
        If poly(i, 2) < pivY Then
            pivY = poly(i, 2)
            pivX = poly(i, 1)
        End If
    Next i

    Dim maxDeg As Double
    Dim stepDeg As Double
    maxDeg = 15#
    stepDeg = 0.25

    Dim s As Double
    Dim testSign As Variant
    Dim testPoly() As Double
    Dim ok As Boolean
    Dim ob As Variant
    Dim rad As Double

    For s = 0 To maxDeg Step stepDeg

        For Each testSign In Array(1, -1)

            If Not (s = 0 And testSign = -1) Then

                rad = CDbl(testSign) * s * 3.14159265358979 / 180#

                ReDim testPoly(1 To n, 1 To 2)

                For i = 1 To n
                    RotateAboutPt poly(i, 1), poly(i, 2), _
                                  pivX, pivY, rad, _
                                  testPoly(i, 1), testPoly(i, 2)
                Next i

                ok = True

                For Each ob In obstacles
                    If N_ConvexPolysOverlap(testPoly, ob) Then
                        ok = False
                        Exit For
                    End If
                Next ob

                If ok Then
                    NudgeClearByRotation = testPoly
                    Exit Function
                End If

            End If

        Next testSign

    Next s

    NudgeClearByRotation = testPoly

End Function

Private Sub TransformPoint(px As Double, py As Double, segIdx As Long, PLx As Double, PLy As Double, theta1 As Double, PRx As Double, PRy As Double, theta3 As Double, Rmid() As Double, TmX As Double, TmY As Double, ByRef npx As Double, ByRef npy As Double)
    Select Case segIdx
        Case 0: npx = px: npy = py
        Case 1: Call RotateAboutPt(px, py, PLx, PLy, theta1, npx, npy)
        Case 2: npx = Rmid(1, 1) * px + Rmid(1, 2) * py + TmX: npy = Rmid(2, 1) * px + Rmid(2, 2) * py + TmY
        Case 3: Call RotateAboutPt(px, py, PRx, PRy, theta3, npx, npy)
        Case 4: npx = px: npy = py
    End Select
End Sub

Private Sub RotateAboutPt(px As Double, py As Double, cx As Double, cy As Double, angle As Double, ByRef nx As Double, ByRef ny As Double)
    Dim c As Double, s As Double: c = Cos(angle): s = Sin(angle)
    nx = c * (px - cx) - s * (py - cy) + cx: ny = s * (px - cx) + c * (py - cy) + cy
End Sub

Private Function Atn2b(Y As Double, X As Double) As Double
    Const PI As Double = 3.14159265358979
    If X > 0 Then Atn2b = Atn(Y / X) Else If X < 0 Then If Y >= 0 Then Atn2b = Atn(Y / X) + PI Else Atn2b = Atn(Y / X) - PI Else If Y > 0 Then Atn2b = PI / 2 Else If Y < 0 Then Atn2b = -PI / 2 Else Atn2b = 0
End Function

Private Sub RigidTransform2(a1x As Double, a1y As Double, a2x As Double, a2y As Double, b1x As Double, b1y As Double, b2x As Double, b2y As Double, ByRef r() As Double, ByRef Tx As Double, ByRef Ty As Double)
    Dim ang As Double: ang = Atn2b(b2y - b1y, b2x - b1x) - Atn2b(a2y - a1y, a2x - a1x)
    Dim c As Double, s As Double: c = Cos(ang): s = Sin(ang)
    r(1, 1) = c: r(1, 2) = -s: r(2, 1) = s: r(2, 2) = c
    Tx = b1x - (c * a1x - s * a1y): Ty = b1y - (s * a1x + c * a1y)
End Sub

Private Function TransformPolyBySeg(ax As Variant, ay As Variant, segIdx As Long, _
                                     PLx As Double, PLy As Double, theta1 As Double, _
                                     PRx As Double, PRy As Double, theta3 As Double, _
                                     Rmid() As Double, TmX As Double, TmY As Double) As Double()
    Dim n As Long: n = UBound(ax)
    Dim result() As Double
    ReDim result(1 To n, 1 To 2)
    Dim c As Long, nx As Double, ny As Double
    For c = 1 To n
        Select Case segIdx
            Case 1: Call RotateAboutPt(CDbl(ax(c)), CDbl(ay(c)), PLx, PLy, theta1, nx, ny)
            Case 2: nx = Rmid(1, 1) * ax(c) + Rmid(1, 2) * ay(c) + TmX: ny = Rmid(2, 1) * ax(c) + Rmid(2, 2) * ay(c) + TmY
            Case 3: Call RotateAboutPt(CDbl(ax(c)), CDbl(ay(c)), PRx, PRy, theta3, nx, ny)
            Case Else: nx = ax(c): ny = ay(c)
        End Select
        result(c, 1) = nx: result(c, 2) = ny
    Next c
    TransformPolyBySeg = result
End Function

Private Function MaxPenetrationFull(bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, _
                                     NB As Long, segOf() As Long, _
                                     PLx As Double, PLy As Double, theta1 As Double, _
                                     PRx As Double, PRy As Double, theta3 As Double, _
                                     Rmid() As Double, TmX As Double, TmY As Double) As Double
    Dim worst As Double: worst = 0
    Dim allPolys As Collection: Set allPolys = New Collection
    Dim blkSeg() As Long: ReDim blkSeg(1 To NB)
    Dim i As Long

    For i = 1 To NB
        Dim polys As Collection: Set polys = New Collection
        blkSeg(i) = segOf(i)
        If Not IsEmpty(bx(i)) Then
            polys.Add TransformPolyBySeg(bx(i), by(i), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        End If
        If Not fx(i) Is Nothing Then
            Dim kf As Long
            For kf = 1 To fx(i).Count
                polys.Add TransformPolyBySeg(fx(i)(kf), fy(i)(kf), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
            Next kf
        End If
        allPolys.Add polys
    Next i

    Dim iBlk As Long, jBlk As Long
    For iBlk = 1 To NB - 1
        For jBlk = iBlk + 1 To NB
            If blkSeg(iBlk) <> blkSeg(jBlk) Then
                Dim pa As Variant, pb As Variant, d As Double
                For Each pa In allPolys(iBlk)
                    For Each pb In allPolys(jBlk)
                        d = PolyPenetrationDepth(pa, pb)
                        If d > worst Then worst = d
                    Next pb
                Next pa
            End If
        Next jBlk
    Next iBlk

    MaxPenetrationFull = worst
End Function

Private Function AnyCollisionFull(bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, _
                                   NB As Long, segOf() As Long, _
                                   PLx As Double, PLy As Double, theta1 As Double, _
                                   PRx As Double, PRy As Double, theta3 As Double, _
                                   Rmid() As Double, TmX As Double, TmY As Double) As Boolean
    AnyCollisionFull = False

    Dim allPolys As Collection: Set allPolys = New Collection
    Dim blkSeg() As Long: ReDim blkSeg(1 To NB)
    Dim i As Long

    For i = 1 To NB
        Dim polys As Collection: Set polys = New Collection
        blkSeg(i) = segOf(i)
        If Not IsEmpty(bx(i)) Then
            polys.Add TransformPolyBySeg(bx(i), by(i), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
        End If
        If Not fx(i) Is Nothing Then
            Dim kf As Long
            For kf = 1 To fx(i).Count
                polys.Add TransformPolyBySeg(fx(i)(kf), fy(i)(kf), segOf(i), PLx, PLy, theta1, PRx, PRy, theta3, Rmid, TmX, TmY)
            Next kf
        End If
        allPolys.Add polys
    Next i

    Dim iBlk As Long, jBlk As Long
    For iBlk = 1 To NB - 1
        For jBlk = iBlk + 1 To NB
            If blkSeg(iBlk) <> blkSeg(jBlk) Then
                Dim pa As Variant, pb As Variant
                For Each pa In allPolys(iBlk)
                    For Each pb In allPolys(jBlk)
                        If N_ConvexPolysOverlap(pa, pb) Then
                            AnyCollisionFull = True
                            Exit Function
                        End If
                    Next pb
                Next pa
            End If
        Next jBlk
    Next iBlk
End Function

Private Function SignedArea4(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                              x3 As Double, y3 As Double, x4 As Double, y4 As Double) As Double
    SignedArea4 = 0.5 * ((x1 * y2 - x2 * y1) + (x2 * y3 - x3 * y2) + (x3 * y4 - x4 * y3) + (x4 * y1 - x1 * y4))
End Function

Private Function FourBarSolve2(PLx As Double, PLy As Double, H1x As Double, H1y As Double, H2x As Double, H2y As Double, PRx As Double, PRy As Double, theta1 As Double, ByRef H1nx As Double, ByRef H1ny As Double, ByRef H2nx As Double, ByRef H2ny As Double, Optional forceBranch As Long = 0) As Boolean
    Call RotateAboutPt(H1x, H1y, PLx, PLy, theta1, H1nx, H1ny)
    Dim Lc As Double, Lr As Double, d As Double
    Lc = Sqr((H2x - H1x) ^ 2 + (H2y - H1y) ^ 2)
    Lr = Sqr((H2x - PRx) ^ 2 + (H2y - PRy) ^ 2)
    d = Sqr((PRx - H1nx) ^ 2 + (PRy - H1ny) ^ 2)

    If d < 0.000000000001 Or d > Lc + Lr + 0.000001 Or d < Abs(Lc - Lr) - 0.000001 Then
        FourBarSolve2 = False
        Exit Function
    End If

    Dim A As Double, hSq As Double, h As Double
    A = (Lc ^ 2 - Lr ^ 2 + d ^ 2) / (2 * d): hSq = Lc ^ 2 - A ^ 2
    If hSq < 0 Then FourBarSolve2 = False: Exit Function
    h = Sqr(hSq)
    Dim exx As Double, eyy As Double: exx = (PRx - H1nx) / d: eyy = (PRy - H1ny) / d
    Dim mx As Double, my As Double: mx = H1nx + A * exx: my = H1ny + A * eyy
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double
    c1x = mx + h * (-eyy): c1y = my + h * exx: c2x = mx - h * (-eyy): c2y = my - h * exx

    Dim origOrient As Double, areaC1 As Double, areaC2 As Double
    origOrient = SignedArea4(PLx, PLy, H1x, H1y, H2x, H2y, PRx, PRy)
    areaC1 = SignedArea4(PLx, PLy, H1nx, H1ny, c1x, c1y, PRx, PRy)
    areaC2 = SignedArea4(PLx, PLy, H1nx, H1ny, c2x, c2y, PRx, PRy)

    If forceBranch = 1 Then
        H2nx = c1x: H2ny = c1y
    ElseIf forceBranch = 2 Then
        H2nx = c2x: H2ny = c2y
    ElseIf Sgn(areaC1) = Sgn(origOrient) And Sgn(areaC1) <> 0 And Sgn(areaC2) <> Sgn(origOrient) Then
        H2nx = c1x: H2ny = c1y
    ElseIf Sgn(areaC2) = Sgn(origOrient) And Sgn(areaC2) <> 0 And Sgn(areaC1) <> Sgn(origOrient) Then
        H2nx = c2x: H2ny = c2y
    Else
        If (c1x - H2x) ^ 2 + (c1y - H2y) ^ 2 < (c2x - H2x) ^ 2 + (c2y - H2y) ^ 2 Then
            H2nx = c1x: H2ny = c1y
        Else
            H2nx = c2x: H2ny = c2y
        End If
    End If
    FourBarSolve2 = True
End Function

Private Function TransformPoly(ax As Variant, ay As Variant, mode As Long, side As Long, _
                                p1 As Double, p2 As Double, ang As Double, _
                                r() As Double, Tx As Double, Ty As Double) As Double()
    Dim n As Long: n = UBound(ax)
    Dim result() As Double
    ReDim result(1 To n, 1 To 2)
    Dim c As Long, nx As Double, ny As Double
    Dim useRotate As Boolean
    If mode = 1 Then useRotate = (side = 1) Else useRotate = (side = 2)
    For c = 1 To n
        If useRotate Then
            Call RotateAboutPt(CDbl(ax(c)), CDbl(ay(c)), p1, p2, ang, nx, ny)
        Else
            nx = r(1, 1) * ax(c) + r(1, 2) * ay(c) + Tx
            ny = r(2, 1) * ax(c) + r(2, 2) * ay(c) + Ty
        End If
        result(c, 1) = nx: result(c, 2) = ny
    Next c
    TransformPoly = result
End Function

Private Function SegmentsCollide(bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, _
                                  NB As Long, blockA As Long, blockB As Long, _
                                  p1 As Double, p2 As Double, ang As Double, _
                                  r() As Double, Tx As Double, Ty As Double, mode As Long) As Boolean
    SegmentsCollide = False
    If blockA < 1 Or blockA > NB Or blockB < 1 Or blockB > NB Then Exit Function

    Dim polysA As Collection, polysB As Collection
    Set polysA = New Collection
    Set polysB = New Collection
    Dim k As Long

    If Not IsEmpty(bx(blockA)) Then polysA.Add TransformPoly(bx(blockA), by(blockA), mode, 1, p1, p2, ang, r, Tx, Ty)
    If Not fx(blockA) Is Nothing Then
        For k = 1 To fx(blockA).Count
            polysA.Add TransformPoly(fx(blockA)(k), fy(blockA)(k), mode, 1, p1, p2, ang, r, Tx, Ty)
        Next k
    End If

    If Not IsEmpty(bx(blockB)) Then polysB.Add TransformPoly(bx(blockB), by(blockB), mode, 2, p1, p2, ang, r, Tx, Ty)
    If Not fx(blockB) Is Nothing Then
        For k = 1 To fx(blockB).Count
            polysB.Add TransformPoly(fx(blockB)(k), fy(blockB)(k), mode, 2, p1, p2, ang, r, Tx, Ty)
        Next k
    End If

    Dim pa As Variant, pb As Variant
    For Each pa In polysA
        For Each pb In polysB
            If N_ConvexPolysOverlap(pa, pb) Then
                SegmentsCollide = True
                Exit Function
            End If
        Next pb
    Next pa
End Function

