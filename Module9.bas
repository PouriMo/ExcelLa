Attribute VB_Name = "Module9"
Option Explicit

Public Function GetInterfaceThickness(wsData As Worksheet, NI As Long) As Double()
    Dim tArr() As Double
    ReDim tArr(1 To NI)
    Dim i As Long, v As Variant, lastGood As Double

    For i = 1 To NI
        v = wsData.Cells(i + 1, 6).Value
        If IsNumeric(v) And Not IsEmpty(v) Then
            tArr(i) = CDbl(v)
            lastGood = tArr(i)
        ElseIf lastGood > 0 Then
            tArr(i) = lastGood
        End If
    Next i
    GetInterfaceThickness = tArr
End Function

Public Function FindHingeCandidates(eVec() As Double, tArr() As Double, NI As Long, ByRef rawH() As Long, ByRef nRaw As Long) As Double
    Dim tolSteps As Variant
    tolSteps = Array(0.005, 0.01, 0.02, 0.05, 0.1)
    Dim s As Long, i As Long, usedTol As Double
    usedTol = -1
    ReDim rawH(1 To NI)
    
    For s = 0 To UBound(tolSteps)
        nRaw = 0
        For i = 1 To NI
            If tArr(i) / 2 > 0 Then
                If Abs(Abs(eVec(i)) - tArr(i) / 2) < CDbl(tolSteps(s)) * (tArr(i) / 2) Then
                    nRaw = nRaw + 1
                    rawH(nRaw) = i
                End If
            End If
        Next i
        If nRaw >= 4 Then
            usedTol = tolSteps(s)
            Exit For
        End If
    Next s

    ' FIX: if even the loosest tolerance band (10% of t/2) never turns up 4
    ' candidates, that does NOT mean the arch has fewer than 4 real hinges -
    ' a genuine hinge's eccentricity can fall short of t/2 due to LP
    ' rounding, especially on irregular geometry. A true hinge is, by
    ' definition, a LOCAL PEAK of |e|/(t/2) along the arch - so fall back to
    ' picking the actual local maxima instead of insisting on the tolerance.
        If nRaw < 4 Then
        Dim ratio() As Double: ReDim ratio(1 To NI)
        For i = 1 To NI
            If tArr(i) / 2 > 0 Then ratio(i) = Abs(eVec(i)) / (tArr(i) / 2) Else ratio(i) = 0
        Next i

        ' FIX: at fine discretizations the abutment interfaces (i=1 and
        ' i=NI) can have artificially high |e|/(t/2) ratios that are NOT
        ' real hinges - they're just the endpoints of the thrust line.
        ' Look for peaks over a WINDOW of neighbouring interfaces so that
        ' a real hinge (which is a broad local peak spanning several
        ' interfaces at fine discretization) is picked instead of the
        ' single-point spikes at the abutments.
        Dim window As Long
        window = WorksheetFunction.mAx(2, CLng(NI / 20))  ' ~5% of interfaces

        Dim isPeak() As Boolean: ReDim isPeak(1 To NI)
        Dim w As Long, jj As Long, isMax As Boolean
        For i = 1 To NI
            isMax = True
            For w = -window To window
                jj = i + w
                If jj >= 1 And jj <= NI And jj <> i Then
                    If ratio(jj) > ratio(i) Then
                        isMax = False
                        Exit For
                    End If
                End If
            Next w
            isPeak(i) = isMax And (ratio(i) > 0.5)  ' require at least 50% of t/2
        Next i

        Dim peakIdx() As Long, peakCount As Long
        ReDim peakIdx(1 To NI): peakCount = 0
        For i = 1 To NI
            If isPeak(i) Then peakCount = peakCount + 1: peakIdx(peakCount) = i
        Next i

        Dim A As Long, b As Long, bestJ As Long, tmp As Long, keepCount As Long
        For A = 1 To peakCount
            bestJ = A
            For b = A + 1 To peakCount
                If ratio(peakIdx(b)) > ratio(peakIdx(bestJ)) Then bestJ = b
            Next b
            If bestJ <> A Then tmp = peakIdx(A): peakIdx(A) = peakIdx(bestJ): peakIdx(bestJ) = tmp
        Next A
        keepCount = peakCount
        If keepCount > 4 Then keepCount = 4

        If keepCount > nRaw Then
            Dim kept() As Long: ReDim kept(1 To keepCount)
            For A = 1 To keepCount: kept(A) = peakIdx(A): Next A
            For A = 1 To keepCount - 1
                For b = A + 1 To keepCount
                    If kept(b) < kept(A) Then tmp = kept(A): kept(A) = kept(b): kept(b) = tmp
                Next b
            Next A
            nRaw = keepCount
            For A = 1 To keepCount: rawH(A) = kept(A): Next A
            If keepCount >= 4 Then usedTol = tolSteps(UBound(tolSteps))
        End If
    End If

    FindHingeCandidates = usedTol
End Function

Public Sub ReduceToFourHinges(rawH() As Long, nRaw As Long, eVec() As Double, ByRef finalH() As Long, ByRef nFinal As Long)
    ' FIX: only NEED to pick-and-reduce when there are MORE than 4 raw
    ' candidates. When exactly 4 were found, they ARE the hinge set already -
    ' the old code still ran its sign-flip search on them, which assumes a
    ' strict alternating - + - + pattern and can pick the same interface
    ' twice (dropping a real hinge) whenever the actual sign pattern doesn't
    ' alternate cleanly, as on this geometry (- + - -).
    If nRaw <= 4 Then
        nFinal = nRaw
        ReDim finalH(1 To IIf(nFinal > 0, nFinal, 1))
        Dim k As Long
        For k = 1 To nFinal
            finalH(k) = rawH(k)
        Next k
        Exit Sub
    End If
    
    nFinal = 4
    ReDim finalH(1 To 4)
    finalH(1) = rawH(1)
    finalH(4) = rawH(nRaw)
    
    Dim s1 As Integer, s4 As Integer, i2 As Long, i3 As Long, i As Long
    s1 = Sgn(eVec(rawH(1)))
    s4 = Sgn(eVec(rawH(nRaw)))
    
    For i = 2 To nRaw - 1
        If Sgn(eVec(rawH(i))) <> s1 Then
            i2 = i
            Exit For
        End If
    Next i
    
    For i = nRaw - 1 To 2 Step -1
        If Sgn(eVec(rawH(i))) <> s4 Then
            i3 = i
            Exit For
        End If
    Next i
    
    If i2 > 0 And i3 > 0 And i2 <= i3 Then
        finalH(2) = rawH(i2)
        finalH(3) = rawH(i3)
    Else
        finalH(2) = rawH(2)
        finalH(3) = rawH(nRaw - 1)
    End If
End Sub

Public Sub ReadCSVGeometry(ws As Worksheet, ByRef bx() As Variant, ByRef by() As Variant, ByRef fx() As Variant, ByRef fy() As Variant, ByRef NB As Long)
    Dim lastRow As Long, i As Long, s As String, bType As String, bIdx As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    NB = 0
    
    For i = 2 To lastRow
        s = CStr(ws.Cells(i, 1).Value)
        If Left(s, 5) = "Block" Then
            bIdx = CLng(Trim(Replace(s, "Block", "")))
            If bIdx > NB Then NB = bIdx
        End If
    Next i
    
    If NB = 0 Then
        ReDim bx(1 To 1): ReDim by(1 To 1): ReDim fx(1 To 1): ReDim fy(1 To 1)
        Exit Sub
    End If
    
    ReDim bx(1 To NB)
    ReDim by(1 To NB)
    ReDim fx(1 To NB)
    ReDim fy(1 To NB)
    
    For i = 1 To NB
        Set fx(i) = New Collection
        Set fy(i) = New Collection
    Next i
    
    Dim curBlock As Long, isBackfill As Boolean
    Dim curX As New Collection, curY As New Collection
    
    For i = 2 To lastRow + 1
        s = CStr(ws.Cells(i, 1).Value)
        
        If s <> "" Or i > lastRow Then
            Dim px As Double, py As Double
            
            If i <= lastRow Then
                If IsNumeric(ws.Cells(i, 2).Value) Then px = CDbl(ws.Cells(i, 2).Value) Else px = 0
                If IsNumeric(ws.Cells(i, 3).Value) Then py = CDbl(ws.Cells(i, 3).Value) Else py = 0
            End If
            
            If s <> "" Then
                If Left(s, 5) = "Block" Then
                    bType = "Block"
                    bIdx = CLng(Trim(Replace(s, "Block", "")))
                End If
                If Left(s, 8) = "Backfill" Then
                    bType = "Backfill"
                    bIdx = CLng(Trim(Replace(s, "Backfill", "")))
                End If
            End If
            
            If curBlock = 0 And bIdx > 0 Then
                curBlock = bIdx
                isBackfill = (bType = "Backfill")
            ElseIf (bIdx <> curBlock Or bType <> IIf(isBackfill, "Backfill", "Block") Or i > lastRow) Then
                FlushPoly curX, curY, isBackfill, curBlock, bx, by, fx, fy
                curBlock = bIdx
                isBackfill = (bType = "Backfill")
            End If
            
            If i <= lastRow And s <> "" Then
                curX.Add px
                curY.Add py
                If curX.Count > 1 Then
                    If Abs(px - curX(1)) < 0.001 And Abs(py - curY(1)) < 0.001 Then
                        FlushPoly curX, curY, isBackfill, curBlock, bx, by, fx, fy
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub FlushPoly(ByRef cx As Collection, ByRef cy As Collection, isBackfill As Boolean, cBlock As Long, ByRef bx() As Variant, ByRef by() As Variant, ByRef fx() As Variant, ByRef fy() As Variant)
    If cx.Count = 0 Or cBlock = 0 Then Exit Sub
    Dim n As Long: n = cx.Count
    If n > 1 Then
        If Abs(cx(n) - cx(1)) < 0.001 And Abs(cy(n) - cy(1)) < 0.001 Then
            n = n - 1
        End If
    End If
    Dim ax() As Double, ay() As Double, j As Long
    ReDim ax(1 To n)
    ReDim ay(1 To n)
    
    For j = 1 To n
        ax(j) = cx(j)
        ay(j) = cy(j)
    Next j
    
    If isBackfill Then
        fx(cBlock).Add ax
        fy(cBlock).Add ay
    Else
        bx(cBlock) = ax
        by(cBlock) = ay
    End If
    
    Set cx = New Collection
    Set cy = New Collection
End Sub

Public Sub DrawBlocksOnChart(chObj As ChartObject, ws As Worksheet, xCol As Long, yCol As Long, startRow As Long, endRow As Long, fillColor As Long, lineRgb As Long, lineWt As Double, isHatch As Boolean)
    Dim ch As Chart
    Set ch = chObj.Chart
    
    Dim axX As Axis, axY As Axis, pArea As PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    Set pArea = ch.PlotArea
    
    Dim scaleX As Double, scaleY As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim builder As FreeformBuilder, shp As Shape
    Dim r As Long, isNew As Boolean
    isNew = True
    
    For r = startRow To endRow
        If IsError(ws.Cells(r, xCol).Value) Or IsEmpty(ws.Cells(r, xCol).Value) Then
            If Not builder Is Nothing Then
                Set shp = builder.ConvertToShape
                shp.Fill.Visible = msoTrue
                                If isHatch Then
                    ' FIX: patterned fills in Excel don't natively support the
                    ' .Transparency property, so instead of a solid diagonal
                    ' hatch we use a much sparser, lighter grey pattern and
                    ' also lighten the line/border. This keeps the "backfill
                    ' looks different from blocks" cue without hiding the
                    ' chart title, legend or axis labels behind it.
                    shp.Fill.Patterned msoPatternWideUpwardDiagonal
                    shp.Fill.ForeColor.RGB = RGB(140, 140, 140)
                    shp.Fill.BackColor.RGB = RGB(245, 245, 245)
                    ' Try to apply transparency (works on newer Excel, silently
                    ' skipped on older versions where patterns are opaque)
                    On Error Resume Next
                    shp.Fill.Transparency = 0.6
                    On Error GoTo 0
                Else
                    shp.Fill.Solid
                    shp.Fill.ForeColor.RGB = fillColor
                    shp.Fill.Transparency = 0.5
                End If
                shp.Line.Visible = msoTrue
                shp.Line.ForeColor.RGB = lineRgb
                shp.Line.Weight = lineWt
                shp.ZOrder msoSendToBack
                Set builder = Nothing
            End If
            isNew = True
        Else
            Dim px As Double, py As Double
            px = pArea.InsideLeft + (ws.Cells(r, xCol).Value - axX.MinimumScale) * scaleX
            py = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yCol).Value) * scaleY
            
            If isNew Then
                Set builder = ch.Shapes.BuildFreeform(msoEditingAuto, px, py)
                isNew = False
            Else
                builder.AddNodes msoSegmentLine, msoEditingAuto, px, py
            End If
        End If
    Next r
End Sub

Public Sub DrawHingesOnChart(chObj As ChartObject, ws As Worksheet, xCol As Long, yCol As Long, startRow As Long, endRow As Long, borderRgb As Long)
    Dim ch As Chart, pArea As PlotArea, axX As Axis, axY As Axis
    Set ch = chObj.Chart
    Set pArea = ch.PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    
    Dim scaleX As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    Dim scaleY As Double
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim r As Long, radius As Double
    radius = 5
    
    For r = startRow To endRow
        If IsNumeric(ws.Cells(r, xCol).Value) And Not IsEmpty(ws.Cells(r, xCol).Value) Then
            Dim px As Double, py As Double
            px = pArea.InsideLeft + (ws.Cells(r, xCol).Value - axX.MinimumScale) * scaleX
            py = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yCol).Value) * scaleY
            
            Dim shp As Shape
            Set shp = ch.Shapes.AddShape(msoShapeOval, px - radius, py - radius, radius * 2, radius * 2)
            shp.Fill.Visible = msoTrue
            shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = borderRgb
            shp.Line.Weight = 1.5
        End If
    Next r
End Sub

Public Sub FindEdgeByNormal(ax() As Double, ay() As Double, targetFiDeg As Double, ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double)
    Dim i As Long, n As Long
    n = UBound(ax)
    Dim bestDiff As Double
    bestDiff = 1000
    
    For i = 1 To n
        Dim px1 As Double, py1 As Double, px2 As Double, py2 As Double
        px1 = ax(i)
        py1 = ay(i)
        
        If i = n Then
            px2 = ax(1)
            py2 = ay(1)
        Else
            px2 = ax(i + 1)
            py2 = ay(i + 1)
        End If
        
        Dim ang As Double, fi As Double
        ang = WorksheetFunction.Atan2(px2 - px1, py2 - py1) * 180 / WorksheetFunction.PI()
        fi = ang - 90
        Do While fi > 180
            fi = fi - 360
        Loop
        Do While fi <= -180
            fi = fi + 360
        Loop
        
        Dim diff1 As Double, diff2 As Double
        diff1 = Abs(fi - targetFiDeg)
        Do While diff1 > 180
            diff1 = Abs(diff1 - 360)
        Loop
        
        diff2 = Abs((fi + 180) - targetFiDeg)
        Do While diff2 > 180
            diff2 = Abs(diff2 - 360)
        Loop
        
        If diff1 < bestDiff Then
            bestDiff = diff1
            p1x = px1
            p1y = py1
            p2x = px2
            p2y = py2
        End If
        
        If diff2 < bestDiff Then
            bestDiff = diff2
            p1x = px2
            p1y = py2
            p2x = px1
            p2y = py1
        End If
    Next i
End Sub

Public Sub FindSharedEdge(ax() As Double, ay() As Double, bx() As Double, by() As Double, ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double)
    Dim i As Long, j As Long
    Dim nA As Long, NB As Long
    nA = UBound(ax)
    NB = UBound(bx)

    ' --- FIX: tol used to be a fixed 0.1 (drawing units, e.g. mm). That's
    ' fine at ~30mm block scale (~0.3% of an edge) but becomes ~2%+ of an
    ' edge at ~5mm scale, which can match the WRONG pair of edges as the
    ' "shared interface" between two blocks (or miss the true match and
    ' fall through to the nearest-edge fallback below). Make it relative
    ' to the average edge length of the two polygons being compared instead,
    ' so behaviour doesn't change with how big/small you draw the arch.
    Dim avgEdge As Double, edgeCount As Long, i2 As Long, j2 As Long
    avgEdge = 0: edgeCount = 0
    For i = 1 To nA
        i2 = (i Mod nA) + 1
        avgEdge = avgEdge + Sqr((ax(i2) - ax(i)) ^ 2 + (ay(i2) - ay(i)) ^ 2)
        edgeCount = edgeCount + 1
    Next i
    For j = 1 To NB
        j2 = (j Mod NB) + 1
        avgEdge = avgEdge + Sqr((bx(j2) - bx(j)) ^ 2 + (by(j2) - by(j)) ^ 2)
        edgeCount = edgeCount + 1
    Next j
    If edgeCount > 0 Then avgEdge = avgEdge / edgeCount

    Dim tol As Double
    tol = 0.01 * avgEdge                      ' 1% of the average edge length
    If tol < 0.0000001 Then tol = 0.0000001   ' guard against degenerate/zero-size input
    
    Dim a1x As Double, a1y As Double, a2x As Double, a2y As Double
    Dim b1x As Double, b1y As Double, b2x As Double, b2y As Double

    For i = 1 To nA
        a1x = ax(i)
        a1y = ay(i)
        If i = nA Then
            a2x = ax(1)
            a2y = ay(1)
        Else
            a2x = ax(i + 1)
            a2y = ay(i + 1)
        End If
        
        For j = 1 To NB
            b1x = bx(j)
            b1y = by(j)
            If j = NB Then
                b2x = bx(1)
                b2y = by(1)
            Else
                b2x = bx(j + 1)
                b2y = by(j + 1)
            End If
            
            If (Abs(a1x - b1x) < tol And Abs(a1y - b1y) < tol And Abs(a2x - b2x) < tol And Abs(a2y - b2y) < tol) Or _
               (Abs(a1x - b2x) < tol And Abs(a1y - b2y) < tol And Abs(a2x - b1x) < tol And Abs(a2y - b1y) < tol) Then
                p1x = a1x
                p1y = a1y
                p2x = a2x
                p2y = a2y
                Exit Sub
            End If
        Next j
    Next i

    Dim bestDist As Double: bestDist = 1E+30
    Dim bp1x As Double, bp1y As Double, bp2x As Double, bp2y As Double
    Dim mAx As Double, mAy As Double, mBx As Double, mBy As Double, d As Double

    For i = 1 To nA
        a1x = ax(i)
        a1y = ay(i)
        If i = nA Then
            a2x = ax(1)
            a2y = ay(1)
        Else
            a2x = ax(i + 1)
            a2y = ay(i + 1)
        End If
        mAx = (a1x + a2x) / 2
        mAy = (a1y + a2y) / 2
        
        For j = 1 To NB
            b1x = bx(j)
            b1y = by(j)
            If j = NB Then
                b2x = bx(1)
                b2y = by(1)
            Else
                b2x = bx(j + 1)
                b2y = by(j + 1)
            End If
            mBx = (b1x + b2x) / 2
            mBy = (b1y + b2y) / 2
            
            d = (mAx - mBx) ^ 2 + (mAy - mBy) ^ 2
            If d < bestDist Then
                bestDist = d
                bp1x = a1x: bp1y = a1y: bp2x = a2x: bp2y = a2y
            End If
        Next j
    Next i

    p1x = bp1x: p1y = bp1y: p2x = bp2x: p2y = bp2y
End Sub

Public Sub OrientEdgeByFi(ByRef p1x As Double, ByRef p1y As Double, ByRef p2x As Double, ByRef p2y As Double, targetFiDeg As Double)
    Dim ang As Double, fi As Double
    ang = WorksheetFunction.Atan2(p2x - p1x, p2y - p1y) * 180 / WorksheetFunction.PI()
    fi = ang - 90
    Do While fi > 180
        fi = fi - 360
    Loop
    Do While fi <= -180
        fi = fi + 360
    Loop
    
    Dim diffSame As Double, diffFlip As Double
    diffSame = Abs(fi - targetFiDeg)
    Do While diffSame > 180
        diffSame = Abs(diffSame - 360)
    Loop
    diffFlip = Abs((fi + 180) - targetFiDeg)
    Do While diffFlip > 180
        diffFlip = Abs(diffFlip - 360)
    Loop
    
    If diffFlip < diffSame Then
        Dim Tx As Double, Ty As Double
        Tx = p1x: Ty = p1y
        p1x = p2x: p1y = p2y
        p2x = Tx: p2y = Ty
    End If
End Sub

Public Function N_ConvexPolysOverlap(polyA As Variant, polyB As Variant) As Boolean
    N_ConvexPolysOverlap = True
    Dim shrink As Double
    shrink = 0.94
    
    Dim cax As Double, cay As Double, cbx As Double, cby As Double
    Dim i As Long, p As Long, poly As Long
    Dim nA As Long, NB As Long
    
    nA = UBound(polyA, 1)
    NB = UBound(polyB, 1)
    
    For i = 1 To nA
        cax = cax + polyA(i, 1)
        cay = cay + polyA(i, 2)
    Next i
    cax = cax / nA
    cay = cay / nA
    
    For i = 1 To NB
        cbx = cbx + polyB(i, 1)
        cby = cby + polyB(i, 2)
    Next i
    cbx = cbx / NB
    cby = cby / NB
    
    Dim A() As Double, b() As Double
    ReDim A(1 To nA, 1 To 2)
    ReDim b(1 To NB, 1 To 2)
    
    For i = 1 To nA
        A(i, 1) = cax + (polyA(i, 1) - cax) * shrink
        A(i, 2) = cay + (polyA(i, 2) - cay) * shrink
    Next i
    
    For i = 1 To NB
        b(i, 1) = cbx + (polyB(i, 1) - cbx) * shrink
        b(i, 2) = cby + (polyB(i, 2) - cby) * shrink
    Next i
    
    Dim edgeX As Double, edgeY As Double, axisX As Double, axisY As Double
    Dim minA As Double, maxA As Double, minB As Double, maxB As Double, proj As Double
    
    For poly = 1 To 2
        Dim nEdges As Long
        If poly = 1 Then
            nEdges = nA
        Else
            nEdges = NB
        End If
        
        For p = 1 To nEdges
            Dim p2 As Long
            p2 = (p Mod nEdges) + 1
            
            If poly = 1 Then
                edgeX = A(p2, 1) - A(p, 1)
                edgeY = A(p2, 2) - A(p, 2)
            Else
                edgeX = b(p2, 1) - b(p, 1)
                edgeY = b(p2, 2) - b(p, 2)
            End If
            
            axisX = -edgeY
            axisY = edgeX
            
            If Not (axisX = 0 And axisY = 0) Then
                minA = 1E+30
                maxA = -1E+30
                
                For i = 1 To nA
                    proj = A(i, 1) * axisX + A(i, 2) * axisY
                    If proj < minA Then minA = proj
                    If proj > maxA Then maxA = proj
                Next i
                
                minB = 1E+30
                maxB = -1E+30
                
                For i = 1 To NB
                    proj = b(i, 1) * axisX + b(i, 2) * axisY
                    If proj < minB Then minB = proj
                    If proj > maxB Then maxB = proj
                Next i
                
                If maxA < minB Or maxB < minA Then
                    N_ConvexPolysOverlap = False
                    Exit Function
                End If
            End If
        Next p
    Next poly
End Function

Public Sub GetBlockForces(wsData As Worksheet, NB As Long, ByRef FxArr() As Double, ByRef FyArr() As Double)
    ReDim FxArr(1 To NB)
    ReDim FyArr(1 To NB)
    Dim lastRow As Long, r As Long, s As String, bIdx As Long, curBlock As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    curBlock = 0
    For r = 2 To lastRow
        s = Trim(CStr(wsData.Cells(r, 1).Value))
        If Left(s, 5) = "Block" Then
            bIdx = CLng(Trim(Replace(s, "Block", "")))
            If bIdx > curBlock Then
                curBlock = bIdx
                If curBlock >= 1 And curBlock <= NB Then
                    If IsNumeric(wsData.Cells(r, 12).Value) Then FxArr(curBlock) = CDbl(wsData.Cells(r, 12).Value)
                    If IsNumeric(wsData.Cells(r, 13).Value) Then FyArr(curBlock) = CDbl(wsData.Cells(r, 13).Value)
                End If
            End If
        End If
    Next r
End Sub

Public Sub DrawForceArrows(chObj As ChartObject, ws As Worksheet, xStartCol As Long, yStartCol As Long, xEndCol As Long, yEndCol As Long, startRow As Long, endRow As Long, lineRgb As Long)
    Dim ch As Chart, pArea As PlotArea, axX As Axis, axY As Axis
    Set ch = chObj.Chart
    Set pArea = ch.PlotArea
    Set axX = ch.Axes(xlCategory)
    Set axY = ch.Axes(xlValue)
    
    Dim scaleX As Double, scaleY As Double
    scaleX = pArea.InsideWidth / (axX.MaximumScale - axX.MinimumScale)
    scaleY = pArea.InsideHeight / (axY.MaximumScale - axY.MinimumScale)
    
    Dim r As Long
    For r = startRow To endRow
        If IsNumeric(ws.Cells(r, xStartCol).Value) And Not IsEmpty(ws.Cells(r, xStartCol).Value) Then
            Dim sx As Double, sy As Double, ex As Double, ey As Double
            sx = pArea.InsideLeft + (ws.Cells(r, xStartCol).Value - axX.MinimumScale) * scaleX
            sy = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yStartCol).Value) * scaleY
            ex = pArea.InsideLeft + (ws.Cells(r, xEndCol).Value - axX.MinimumScale) * scaleX
            ey = pArea.InsideTop + (axY.MaximumScale - ws.Cells(r, yEndCol).Value) * scaleY
            
            Dim shp As Shape
            Set shp = ch.Shapes.AddLine(sx, sy, ex, ey)
            shp.Line.ForeColor.RGB = lineRgb
            shp.Line.Weight = 2.25
            shp.Line.EndArrowheadStyle = msoArrowheadTriangle
            shp.Line.EndArrowheadLength = msoArrowheadLong
            shp.Line.EndArrowheadWidth = msoArrowheadWide

            Dim dotR As Double: dotR = 3.5
            Dim dot As Shape
            Set dot = ch.Shapes.AddShape(msoShapeOval, sx - dotR, sy - dotR, dotR * 2, dotR * 2)
            dot.Fill.Visible = msoTrue
            dot.Fill.ForeColor.RGB = lineRgb
            dot.Line.Visible = msoFalse

            Dim lbl As Shape
            Set lbl = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, sx + 4, sy - 16, 16, 14)
            lbl.TextFrame.Characters.Text = "P"
            lbl.TextFrame.Characters.Font.Bold = True
            lbl.TextFrame.Characters.Font.Size = 10
            lbl.TextFrame.Characters.Font.Color = lineRgb
            lbl.Fill.Visible = msoFalse
            lbl.Line.Visible = msoFalse
            lbl.TextFrame.MarginLeft = 0: lbl.TextFrame.MarginRight = 0
            lbl.TextFrame.MarginTop = 0: lbl.TextFrame.MarginBottom = 0
        End If
    Next r
End Sub
Sub DumpHingeDiagnostics()
    Dim wbUser As Workbook: Set wbUser = ActiveWorkbook
    Dim wsData As Worksheet: Set wsData = wbUser.ActiveSheet
    Dim wsForces As Worksheet: Set wsForces = wbUser.Sheets("ForcesTable")

    Dim bx() As Variant, by() As Variant, fx() As Variant, fy() As Variant, NB As Long
    Call ReadCSVGeometry(wsData, bx, by, fx, fy, NB)
    Dim NI As Long: NI = NB + 1

    Dim tArr() As Double: tArr = GetInterfaceThickness(wsData, NI)
    Dim Nvec() As Double, Mvec() As Double: ReDim Nvec(1 To NI): ReDim Mvec(1 To NI)
    Dim i As Long
    For i = 1 To NI
        Nvec(i) = wsForces.Cells(i + 1, 2).Value
        Mvec(i) = wsForces.Cells(i + 1, 4).Value
    Next i

    Dim FiArr() As Double: ReDim FiArr(1 To NI)
    For i = 1 To NI: FiArr(i) = wsData.Cells(i + 1, 8).Value: Next i

    Dim p1x() As Double, p1y() As Double, p2x() As Double, p2y() As Double
    ReDim p1x(1 To NI), p1y(1 To NI), p2x(1 To NI), p2y(1 To NI)
    Dim curX() As Double, curY() As Double, prevX() As Double, prevY() As Double
    For i = 1 To NI
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

    On Error Resume Next: wbUser.Sheets("Diag").Delete: On Error GoTo 0
    Dim ws As Worksheet: Set ws = wbUser.Sheets.Add: ws.Name = "Diag"
    ws.Range("A1:J1").Value = Array("i", "N", "M", "t", "e", "ratio", "p1x", "p1y", "p2x", "p2y")
    For i = 1 To NI
        Dim eV As Double: eV = 0
        If Abs(Nvec(i)) > 0.000000001 Then eV = -1 * Mvec(i) / Nvec(i)
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = Nvec(i)
        ws.Cells(i + 1, 3).Value = Mvec(i)
        ws.Cells(i + 1, 4).Value = tArr(i)
        ws.Cells(i + 1, 5).Value = eV
        If tArr(i) > 0 Then ws.Cells(i + 1, 6).Value = Abs(eV) / (tArr(i) / 2)
        ws.Cells(i + 1, 7).Value = p1x(i): ws.Cells(i + 1, 8).Value = p1y(i)
        ws.Cells(i + 1, 9).Value = p2x(i): ws.Cells(i + 1, 10).Value = p2y(i)
    Next i
    ws.Columns.AutoFit
    MsgBox "Diagnostics written to 'Diag' sheet."
End Sub

' Returns 0 if the two convex polygons are separated (or exactly touching),
' otherwise the magnitude of the Minimum Translation Vector needed to
' separate them - i.e. HOW BADLY they overlap, not just whether they do.
' Used to rank candidate crank angles by severity of interpenetration,
' instead of only knowing pass/fail (see PlotDeformedShape's search loop).
Public Function PolyPenetrationDepth(polyA As Variant, polyB As Variant) As Double
    Dim shrink As Double: shrink = 0.98
    Dim nA As Long, NB As Long, i As Long
    nA = UBound(polyA, 1): NB = UBound(polyB, 1)

    Dim cax As Double, cay As Double, cbx As Double, cby As Double
    For i = 1 To nA: cax = cax + polyA(i, 1): cay = cay + polyA(i, 2): Next i
    cax = cax / nA: cay = cay / nA
    For i = 1 To NB: cbx = cbx + polyB(i, 1): cby = cby + polyB(i, 2): Next i
    cbx = cbx / NB: cby = cby / NB

    Dim A() As Double, b() As Double
    ReDim A(1 To nA, 1 To 2): ReDim b(1 To NB, 1 To 2)
    For i = 1 To nA
        A(i, 1) = cax + (polyA(i, 1) - cax) * shrink
        A(i, 2) = cay + (polyA(i, 2) - cay) * shrink
    Next i
    For i = 1 To NB
        b(i, 1) = cbx + (polyB(i, 1) - cbx) * shrink
        b(i, 2) = cby + (polyB(i, 2) - cby) * shrink
    Next i

    Dim minOverlap As Double: minOverlap = 1E+30
    Dim poly As Long, p As Long, p2 As Long, nEdges As Long
    Dim edgeX As Double, edgeY As Double, axisX As Double, axisY As Double, axisLen As Double
    Dim minA As Double, maxA As Double, minB As Double, maxB As Double, proj As Double, ov As Double

    For poly = 1 To 2
        If poly = 1 Then nEdges = nA Else nEdges = NB
        For p = 1 To nEdges
            p2 = (p Mod nEdges) + 1
            If poly = 1 Then
                edgeX = A(p2, 1) - A(p, 1): edgeY = A(p2, 2) - A(p, 2)
            Else
                edgeX = b(p2, 1) - b(p, 1): edgeY = b(p2, 2) - b(p, 2)
            End If
            axisX = -edgeY: axisY = edgeX
            axisLen = Sqr(axisX ^ 2 + axisY ^ 2)
            If axisLen > 0.000000001 Then
                axisX = axisX / axisLen: axisY = axisY / axisLen

                minA = 1E+30: maxA = -1E+30
                For i = 1 To nA
                    proj = A(i, 1) * axisX + A(i, 2) * axisY
                    If proj < minA Then minA = proj
                    If proj > maxA Then maxA = proj
                Next i
                minB = 1E+30: maxB = -1E+30
                For i = 1 To NB
                    proj = b(i, 1) * axisX + b(i, 2) * axisY
                    If proj < minB Then minB = proj
                    If proj > maxB Then maxB = proj
                Next i

                If maxA < minB Or maxB < minA Then
                    PolyPenetrationDepth = 0   ' separating axis exists -> no overlap
                    Exit Function
                End If
                ov = WorksheetFunction.Min(maxA, maxB) - WorksheetFunction.mAx(minA, minB)
                If ov < minOverlap Then minOverlap = ov
            End If
        Next p
    Next poly

    PolyPenetrationDepth = minOverlap
End Function
