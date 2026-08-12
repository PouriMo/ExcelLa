Attribute VB_Name = "Module0"
Option Explicit

' Automatically loads the Solver add-in AND adds the VBA reference to it.
' Call this once from Workbook_Open (or at the start of any macro that needs Solver).
Public Sub EnsureSolverReady()
    Dim solverPath As String
    Dim ai As AddIn
    Dim refAdded As Boolean

    ' --- 1. Enable Solver add-in in Excel ---
    On Error Resume Next
    Set ai = Application.AddIns("Solver Add-in")
    If ai Is Nothing Then Set ai = Application.AddIns("Solver")
    If Not ai Is Nothing Then
        If Not ai.Installed Then ai.Installed = True
    End If
    On Error GoTo 0

    ' --- 2. Add the VBA reference to SOLVER.XLAM ---
    ' Check if already referenced
    Dim ref As Object
    For Each ref In ThisWorkbook.VBProject.References
        If InStr(1, LCase(ref.Name), "solver") > 0 Then
            refAdded = True
            Exit For
        End If
    Next ref

    If Not refAdded Then
        ' Try common paths
        Dim paths As Variant, p As Variant
        paths = Array( _
            Environ("ProgramFiles") & "\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM", _
            Environ("ProgramFiles") & "\Microsoft Office\Office16\Library\SOLVER\SOLVER.XLAM", _
            Environ("ProgramFiles(x86)") & "\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM", _
            Environ("ProgramFiles(x86)") & "\Microsoft Office\Office16\Library\SOLVER\SOLVER.XLAM", _
            Environ("ProgramFiles") & "\Microsoft Office\root\Office15\Library\SOLVER\SOLVER.XLAM", _
            Environ("ProgramFiles") & "\Microsoft Office\Office15\Library\SOLVER\SOLVER.XLAM" _
        )
        
        For Each p In paths
            If Dir(CStr(p)) <> "" Then
                solverPath = CStr(p)
                Exit For
            End If
        Next p

        If solverPath <> "" Then
            On Error Resume Next
            ThisWorkbook.VBProject.References.AddFromFile solverPath
            On Error GoTo 0
        End If
    End If
End Sub
