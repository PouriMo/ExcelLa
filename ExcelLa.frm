VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelLa 
   Caption         =   "Excel Limit Analysis Toolkit"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "ExcelLa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelLa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==================== Helper: Get the user's workbook ====================
Private Function UserWorkbook() As Workbook
    If ActiveWorkbook Is Nothing Then
        Set UserWorkbook = Workbooks.Add
    ElseIf ActiveWorkbook Is ThisWorkbook Then
        Set UserWorkbook = Workbooks.Add
    Else
        Set UserWorkbook = ActiveWorkbook
    End If
End Function

' ==================== Form Initialize ====================
Private Sub UserForm_Initialize()
    Call RefreshSheetList
    txtLambda.Value = "-"
    txtHinges.Value = "-"
    btnShowThrust.Enabled = False
    btnShowDeformed.Enabled = False

    ' Configure text log to handle its own scrolling properly
    txtLog.Value = ""
    txtLog.MultiLine = True
    txtLog.ScrollBars = 2 ' fmScrollBarsVertical
    txtLog.Enabled = True ' Required to allow scrolling interaction
    txtLog.Locked = True  ' Prevents user from typing in the log

    ' Make image panel backgrounds transparent
    imgThrust.BackStyle = 0 ' fmBackStyleTransparent
    imgDeformed.BackStyle = 0 ' fmBackStyleTransparent

    Call LogMsg("Ready. Load or select a CSV data sheet to begin.")
End Sub

' ==================== Data Section ====================
Private Sub btnRefreshSheets_Click()
    Call RefreshSheetList
End Sub

Private Sub RefreshSheetList()
    Dim ws As Worksheet
    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()

    cmbDataSheet.Clear
    For Each ws In wbUser.Worksheets
        Select Case ws.Name
            Case "Gass_in", "Gass", "Fass", "Ain", "Bin", "optimization", _
                 "ForcesTable", "ThrustPlot", "DeformedPlot"
                ' skip generated sheets
            Case Else
                If Left(CStr(ws.Cells(2, 1).Value), 5) = "Block" Then
                    cmbDataSheet.AddItem ws.Name
                End If
        End Select
    Next ws
    If cmbDataSheet.ListCount > 0 Then cmbDataSheet.ListIndex = 0
End Sub

Private Sub btnImportCSV_Click()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Select the CSV file from the LISP export")
    If filePath = False Then Exit Sub

    Call LogMsg("Importing CSV: " & filePath)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()

    Dim wbCSV As Workbook
    Dim sheetName As String
    sheetName = Left(GetFilenameWithoutExt(CStr(filePath)), 31)

    On Error Resume Next
    wbUser.Sheets(sheetName).Delete
    On Error GoTo 0

    Set wbCSV = Workbooks.Open(CStr(filePath))
    wbCSV.Sheets(1).Copy Before:=wbUser.Sheets(1)
    wbUser.Sheets(1).Name = sheetName
    wbCSV.Close SaveChanges:=False

    wbUser.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Call RefreshSheetList
    cmbDataSheet.Value = sheetName
    Call LogMsg("Imported '" & sheetName & "' into '" & wbUser.Name & "'")
End Sub

Private Function GetFilenameWithoutExt(fullPath As String) As String
    Dim fname As String
    fname = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(fname, ".") > 0 Then
        fname = Left(fname, InStrRev(fname, ".") - 1)
    End If
    GetFilenameWithoutExt = fname
End Function

' ==================== Run Full Analysis ====================
Private Sub btnRunAll_Click()
    Dim dataSheet As String
    dataSheet = CStr(cmbDataSheet.Value)
    If dataSheet = "" Then
        MsgBox "Please select or import a CSV data sheet first.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim wbUser As Workbook
    Set wbUser = UserWorkbook()
    wbUser.Sheets(dataSheet).Activate

    Call LogMsg("=== Starting full analysis on '" & wbUser.Name & "' ===")

    Call LogMsg("Step 1/6: Generating Gass_in matrix...")
    Application.Run "Generate_Gass_in"

    wbUser.Sheets(dataSheet).Activate
    Call LogMsg("Step 2/6: Generating Gass matrix...")
    Application.Run "GenerateGassMatrix"

    wbUser.Sheets(dataSheet).Activate
    Call LogMsg("Step 3/6: Creating Fass sheet...")
    Application.Run "CreateFassSheet"

    Call LogMsg("Step 4a/6: Creating Ain matrix...")
    Application.Run "CreateAinSheet"

    Call LogMsg("Step 4b/6: Creating Bin vector...")
    Application.Run "CreateBinSheet"

    Call LogMsg("Step 5/6: Building optimization sheet...")
    Application.Run "CreateOptimizationSheet"

    Call LogMsg("Step 6/6: Running Solver + back-substitution...")
    Application.Run "RunSolverAndBackSubstitute"

    Dim wsForces As Worksheet
    Set wsForces = wbUser.Sheets("ForcesTable")
    txtLambda.Value = Format(wsForces.Range("G2").Value, "0.0000")

    Dim NI As Long, i As Long
    Dim n As Double, m As Double
    NI = wsForces.Cells(wsForces.Rows.Count, "A").End(xlUp).row - 1

    Dim tArr() As Double
    tArr = GetInterfaceThickness(wbUser.Sheets(dataSheet), NI)

    Dim eArr() As Double
    ReDim eArr(1 To NI)
    For i = 1 To NI
        eArr(i) = 0
        n = wsForces.Cells(i + 1, 2).Value
        m = wsForces.Cells(i + 1, 4).Value
        If Abs(n) > 0.000000001 Then eArr(i) = -1 * m / n
    Next i

    Dim rawH() As Long, nRaw As Long, usedTol As Double
    usedTol = FindHingeCandidates(eArr, tArr, NI, rawH, nRaw)

    Dim finalH() As Long, nFinal As Long
    Call ReduceToFourHinges(rawH, nRaw, eArr, finalH, nFinal)

    txtHinges.Value = CStr(nFinal)
    If usedTol = -1 Then
        Call LogMsg("Warning: fewer than 4 interfaces reached their eccentricity limit even at the loosest tolerance tested - the arch may not be at its true collapse state.")
    ElseIf usedTol > 0.005 Then
        Call LogMsg("Note: hinge tolerance had to be relaxed to " & Format(usedTol * 100, "0.0") & "% of local thickness to find " & nFinal & " hinge(s).")
    End If

    btnShowThrust.Enabled = True
    btnShowDeformed.Enabled = True

    Call LogMsg("=== Analysis complete! Lambda = " & _
                Format(wsForces.Range("G2").Value, "0.0000") & " N ===")

    If chkAutoPlots.Value Then
        Call LogMsg("Generating thrust line plot...")
        Application.Run "PlotThrustLine", wbUser.Sheets(dataSheet)
        Call LogMsg("Generating deformed shape plot...")
        Application.Run "PlotDeformedShape", wbUser.Sheets(dataSheet)
        Call LogMsg("Plots created on 'ThrustPlot' and 'DeformedPlot' sheets.")
        Call UpdateImagePanels(wbUser)
        MultiPage1.Value = 1
    End If

    Application.ScreenUpdating = True
    MsgBox "Analysis complete!" & vbCrLf & _
           "Lambda = " & Format(wsForces.Range("G2").Value, "0.0000") & " N" & vbCrLf & _
           "Hinges: " & nFinal, vbInformation, "Done"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call LogMsg("ERROR: " & Err.Description)
    MsgBox "Analysis failed: " & Err.Description, vbCritical
End Sub

' ==================== Show Plots ====================
Private Sub btnShowThrust_Click()
    On Error GoTo ErrorHandler
    Dim wbUser As Workbook: Set wbUser = UserWorkbook()
    If CStr(cmbDataSheet.Value) = "" Then
        MsgBox "Please select a data sheet first.", vbExclamation
        Exit Sub
    End If
    Application.Run "PlotThrustLine", wbUser.Sheets(CStr(cmbDataSheet.Value))
    Call UpdateImagePanels(wbUser)
    MultiPage1.Value = 1
    Exit Sub
ErrorHandler:
    MsgBox "Could not show thrust line: " & Err.Description, vbCritical
End Sub

Private Sub btnShowDeformed_Click()
    On Error GoTo ErrorHandler
    Dim wbUser As Workbook: Set wbUser = UserWorkbook()
    If CStr(cmbDataSheet.Value) = "" Then
        MsgBox "Please select a data sheet first.", vbExclamation
        Exit Sub
    End If
    Application.Run "PlotDeformedShape", wbUser.Sheets(CStr(cmbDataSheet.Value))
    Call UpdateImagePanels(wbUser)
    MultiPage1.Value = 2
    Exit Sub
ErrorHandler:
    MsgBox "Could not show deformed shape: " & Err.Description, vbCritical
End Sub

' ==================== Close ====================
Private Sub btnClose_Click()
    Unload Me
End Sub

' ==================== Logger ====================
Private Sub LogMsg(msg As String)
    Dim timestamp As String
    timestamp = Format(Now, "hh:mm:ss")
    If Len(txtLog.Value) = 0 Then
        txtLog.Value = "[" & timestamp & "] " & msg
    Else
        txtLog.Value = txtLog.Value & vbCrLf & "[" & timestamp & "] " & msg
    End If
    ' Force scroll to bottom
    txtLog.SelStart = Len(txtLog.Value)
    DoEvents
End Sub

' ==================== Image Panels ====================
Private Sub UpdateImagePanels(wbUser As Workbook)
    Dim wsThrust As Worksheet, wsDef As Worksheet
    Dim chThrust As ChartObject, chDef As ChartObject
    Dim pathThrust As String, pathDef As String
    Dim ok As Boolean

    On Error Resume Next
    Set wsThrust = wbUser.Sheets("ThrustPlot")
    Set wsDef = wbUser.Sheets("DeformedPlot")
    On Error GoTo 0

    If Not wsThrust Is Nothing Then
        Set chThrust = wsThrust.ChartObjects(1)
        pathThrust = Environ("TEMP") & "\thrust_temp.gif"
        ok = ExportChartHiRes(chThrust, pathThrust)
        If ok Then
            On Error Resume Next
            imgThrust.PictureSizeMode = 3 ' fmPictureSizeModeZoom
            Set imgThrust.Picture = LoadPicture(pathThrust)
            If Err.Number <> 0 Then
                Call LogMsg("Warning: could not load the Thrust Line image (" & Err.Description & ").")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If

    If Not wsDef Is Nothing Then
        Set chDef = wsDef.ChartObjects(1)
        pathDef = Environ("TEMP") & "\def_temp.gif"
        ok = ExportChartHiRes(chDef, pathDef)
        If ok Then
            On Error Resume Next
            imgDeformed.PictureSizeMode = 3 ' fmPictureSizeModeZoom
            Set imgDeformed.Picture = LoadPicture(pathDef)
            If Err.Number <> 0 Then
                Call LogMsg("Warning: could not load the Deformed Shape image (" & Err.Description & ").")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Sub

' Now simply exports the natively high-res chart exactly as drawn
Private Function ExportChartHiRes(ch As ChartObject, path As String) As Boolean
    On Error GoTo Fail
    ch.Chart.Export Filename:=path, FilterName:="GIF"
    ExportChartHiRes = True
    Exit Function
Fail:
    ExportChartHiRes = False
End Function

