Attribute VB_Name = "Module6"
Option Explicit

Sub RunSolverAndBackSubstitute()
    Dim wsOpt As Worksheet
    Dim wsGass As Worksheet, wsFass As Worksheet
    Dim wsForces As Worksheet
    Dim wbUser As Workbook
    Dim lastRow As Long
    Dim i As Long, j As Long

    On Error GoTo ErrorHandler
    Set wbUser = ActiveWorkbook
    Set wsOpt = wbUser.Sheets("optimization")
    wsOpt.Activate

    lastRow = wsOpt.Cells(wsOpt.Rows.Count, "A").End(xlUp).row

    ' ============================================================
    ' Build the LP model using the Solver-compatible API
    ' (OpenSolver reads the same model definitions).
    ' ============================================================
    Call SetUpOpenSolverModel(wsOpt, lastRow)

    ' ============================================================
    ' Call OpenSolver. Try several known entry-point names in
    ' order until one works. This makes the code portable across
    ' OpenSolver versions.
    ' ============================================================
    Call InvokeOpenSolver

    Debug.Print "After solve: N=" & wsOpt.Range("B2").Value & _
                "  V=" & wsOpt.Range("C2").Value & _
                "  M=" & wsOpt.Range("D2").Value & _
                "  lambda=" & wsOpt.Range("E2").Value

    Dim N_last As Double, V_last As Double, M_last As Double, lam As Double
    N_last = wsOpt.Range("B2").Value
    V_last = wsOpt.Range("C2").Value
    M_last = wsOpt.Range("D2").Value
    lam = wsOpt.Range("E2").Value

    Set wsGass = wbUser.Sheets("Gass")
    Set wsFass = wbUser.Sheets("Fass")

    Dim gassRows As Long, gassCols As Long
    gassRows = wsGass.Cells(wsGass.Rows.Count, "A").End(xlUp).row
    gassCols = wsGass.Cells(1, wsGass.Columns.Count).End(xlToLeft).Column

    Dim G1 As Variant, G2 As Variant
    G1 = wsGass.Range(wsGass.Cells(1, 1), wsGass.Cells(gassRows, gassCols - 3)).Value
    G2 = wsGass.Range(wsGass.Cells(1, gassCols - 2), wsGass.Cells(gassRows, gassCols)).Value

    Dim fassRows As Long
    fassRows = wsFass.Cells(wsFass.Rows.Count, "A").End(xlUp).row
    Dim Fw As Variant, Flam As Variant
    Fw = wsFass.Range("A2:A" & fassRows).Value
    Flam = wsFass.Range("B2:B" & fassRows).Value

    Dim x_last(1 To 3, 1 To 1) As Double
    x_last(1, 1) = N_last: x_last(2, 1) = V_last: x_last(3, 1) = M_last

    Dim G2x() As Double
    ReDim G2x(1 To UBound(G2, 1), 1 To 1)
    Dim sumG2 As Double
    For i = 1 To UBound(G2, 1)
        sumG2 = 0
        For j = 1 To 3
            sumG2 = sumG2 + G2(i, j) * x_last(j, 1)
        Next j
        G2x(i, 1) = sumG2
    Next i

    Dim rhs() As Double
    ReDim rhs(1 To UBound(Fw, 1), 1 To 1)
    For i = 1 To UBound(Fw, 1)
        rhs(i, 1) = Fw(i, 1) + lam * Flam(i, 1) - G2x(i, 1)
    Next i

    Dim xFullVar As Variant
    xFullVar = SolveLinearSystem(G1, rhs)

    Dim xFull() As Double
    ReDim xFull(1 To UBound(xFullVar, 1), 1 To 1)
    For i = 1 To UBound(xFullVar, 1)
        xFull(i, 1) = xFullVar(i, 1)
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("ForcesTable").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsForces = wbUser.Sheets.Add
    wsForces.Name = "ForcesTable"

    wsForces.Cells(1, 1).Value = "Interface"
    wsForces.Cells(1, 2).Value = "N"
    wsForces.Cells(1, 3).Value = "V"
    wsForces.Cells(1, 4).Value = "M"
    wsForces.Cells(1, 5).Value = "e=M/N"
    wsForces.Cells(1, 6).Value = "Status"
    wsForces.Range("A1:F1").Font.Bold = True
    wsForces.Range("A1:F1").Interior.Color = RGB(200, 200, 200)

    Dim NI As Long
    NI = UBound(xFull, 1) / 3 + 1

    Dim Nref As Double: Nref = 0
    For i = 1 To NI - 1
        If Abs(xFull((i - 1) * 3 + 1, 1)) > Nref Then Nref = Abs(xFull((i - 1) * 3 + 1, 1))
    Next i
    If Abs(N_last) > Nref Then Nref = Abs(N_last)
    If Nref = 0 Then Nref = 1

    Dim row As Long: row = 2
    Dim Nval As Double, Vval As Double, Mval As Double, eVal As Double
    For i = 1 To NI - 1
        Nval = xFull((i - 1) * 3 + 1, 1)
        Vval = xFull((i - 1) * 3 + 2, 1)
        Mval = xFull((i - 1) * 3 + 3, 1)
        eVal = 0
        If Abs(Nval) > 0.000001 * Nref Then eVal = Mval / Nval
        wsForces.Cells(row, 1).Value = i
        wsForces.Cells(row, 2).Value = Nval
        wsForces.Cells(row, 3).Value = Vval
        wsForces.Cells(row, 4).Value = Mval
        wsForces.Cells(row, 5).Value = eVal
        wsForces.Cells(row, 6).Value = "Computed"
        row = row + 1
    Next i

    eVal = 0
    If Abs(N_last) > 0.000001 * Nref Then eVal = M_last / N_last
    wsForces.Cells(row, 1).Value = NI
    wsForces.Cells(row, 2).Value = N_last
    wsForces.Cells(row, 3).Value = V_last
    wsForces.Cells(row, 4).Value = M_last
    wsForces.Cells(row, 5).Value = eVal
    wsForces.Cells(row, 6).Value = "Solver (last)"

    wsForces.Cells(1, 7).Value = "Lambda"
    wsForces.Cells(2, 7).Value = lam
    wsForces.Cells(1, 7).Font.Bold = True
    wsForces.Cells(2, 7).NumberFormat = "0.0000"
    wsForces.Columns("A:G").AutoFit
    Exit Sub

ErrorHandler:
    MsgBox "Error in RunSolverAndBackSubstitute: " & Err.Description, vbCritical
End Sub

' ============================================================
' Calls OpenSolver. Different OpenSolver versions expose the
' solve entry point under different names; we try the most
' common ones in order until one succeeds.
' ============================================================
Private Sub InvokeOpenSolver()
    ' Call OpenSolver's public macro directly via the .xlam workbook name.
    ' The workbook is loaded as an Excel Add-in (visible in Options > Add-ins)
    ' so its macros are addressable via Application.Run using the file name.
    
    On Error Resume Next
    
    ' Attempt 1: RunOpenSolver (main entry in OpenSolver 2.9.x)
    Application.Run "OpenSolver.xlam!RunOpenSolver"
    If Err.Number = 0 Then
        Debug.Print "OpenSolver called via: OpenSolver.xlam!RunOpenSolver"
        Exit Sub
    End If
    Debug.Print "  RunOpenSolver failed: " & Err.Description
    Err.Clear
    
    ' Attempt 2: SolveModel
    Application.Run "OpenSolver.xlam!SolveModel"
    If Err.Number = 0 Then
        Debug.Print "OpenSolver called via: OpenSolver.xlam!SolveModel"
        Exit Sub
    End If
    Debug.Print "  SolveModel failed: " & Err.Description
    Err.Clear
    
    ' Attempt 3: Fallback via ribbon Solve button
    Application.CommandBars.ExecuteMso "SolverSolve"
    If Err.Number = 0 Then
        Debug.Print "OpenSolver called via ribbon ExecuteMso"
        Exit Sub
    End If
    Debug.Print "  ExecuteMso failed: " & Err.Description
    Err.Clear
    
    On Error GoTo 0
    Err.Raise vbObjectError + 513, "InvokeOpenSolver", _
              "Could not invoke OpenSolver by any known method."
End Sub

' ============================================================
' Configures the OpenSolver model on the "optimization" sheet.
' Mirrors what SolverOk/SolverAdd/SolverOptions did before,
' but writes the config to defined names that OpenSolver reads.
' Must be called BEFORE OpenSolver.RunOpenSolver().
' ============================================================
Private Sub SetUpOpenSolverModel(wsOpt As Worksheet, lastRow As Long)
    ' Use the Solver-compatible API - OpenSolver intercepts these too
    SolverReset
    SolverOk SetCell:="$F$3", MaxMinVal:=1, ValueOf:=0, _
             ByChange:="$B$2:$E$2"
    SolverAdd CellRef:="$G$5:$G$" & lastRow, Relation:=1, _
              FormulaText:="$I$5:$I$" & lastRow
    SolverOptions AssumeNonNeg:=False, Precision:=0.0000001
    ' Save the model
    SolverFinish KeepFinal:=1
End Sub

Sub TestOpenSolverIsLive()
    On Error Resume Next
    
    ' Try the version function
    Dim v As String
    v = Application.Run("OpenSolver.OpenSolverVersion")
    If Err.Number = 0 And Len(v) > 0 Then
        MsgBox "SUCCESS: OpenSolver version " & v & " is loaded."
        Exit Sub
    End If
    Err.Clear
    
    ' Fallback: try to see if the add-in workbook exists
    Dim wb As Workbook
    For Each wb In Workbooks
        If InStr(1, wb.Name, "OpenSolver", vbTextCompare) > 0 Then
            MsgBox "OpenSolver workbook is loaded as: " & wb.Name & _
                   vbCrLf & "But the OpenSolverVersion macro isn't callable." & _
                   vbCrLf & "This usually means macros are blocked in the add-in."
            Exit Sub
        End If
    Next wb
    
    MsgBox "OpenSolver is NOT loaded in this Excel session."
End Sub
Sub FindOpenSolverFile()
    Dim wb As Workbook
    Dim ai As AddIn
    Dim msg As String
    
    msg = "=== Open Workbooks ===" & vbCrLf
    For Each wb In Workbooks
        msg = msg & wb.Name & "  |  Path: " & wb.path & vbCrLf
    Next wb
    
    msg = msg & vbCrLf & "=== Registered Add-ins ===" & vbCrLf
    For Each ai In Application.AddIns
        If InStr(1, ai.Name, "solver", vbTextCompare) > 0 Or _
           InStr(1, ai.Name, "opensolver", vbTextCompare) > 0 Then
            msg = msg & ai.Name & "  |  Installed=" & ai.Installed & _
                  "  |  Path: " & ai.FullName & vbCrLf
        End If
    Next ai
    
    msg = msg & vbCrLf & "=== COM Add-ins ===" & vbCrLf
    Dim ca As COMAddIn
    For Each ca In Application.COMAddIns
        If InStr(1, ca.Description, "solver", vbTextCompare) > 0 Or _
           InStr(1, ca.Description, "opensolver", vbTextCompare) > 0 Then
            msg = msg & ca.Description & "  |  Connected=" & ca.Connect & vbCrLf
        End If
    Next ca
    
    Debug.Print msg
    MsgBox msg
End Sub

