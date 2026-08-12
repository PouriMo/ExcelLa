Attribute VB_Name = "Module3"
Sub CreateFassSheet()
    Dim wsMain As Worksheet, wsFass As Worksheet
    Dim wbUser As Workbook
    Dim i As Long, fassRow As Long

    Set wbUser = ActiveWorkbook
    Set wsMain = wbUser.ActiveSheet

    On Error Resume Next
    Application.DisplayAlerts = False
    wbUser.Sheets("Fass").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsFass = wbUser.Sheets.Add
    wsFass.Name = "Fass"

    wsFass.Range("A1").Value = "fass_w"
    wsFass.Range("B1").Value = "fass_lam"
    
    ' 1. Determine number of blocks (NB) by counting the contiguous weights in Column J
    Dim NB As Long
    NB = 0
    Do While IsNumeric(wsMain.Cells(2 + NB, "J").Value) And wsMain.Cells(2 + NB, "J").Value <> ""
        NB = NB + 1
    Loop
    
    If NB = 0 Then Exit Sub
    
    ' 2. Scan the sheet to capture the Fx and Fy lambda forces for each block
    Dim fx() As Double, fy() As Double
    ReDim fx(1 To NB)
    ReDim fy(1 To NB)
    
    Dim r As Long, lastRow As Long
    Dim currentBlock As Long
    currentBlock = 0
    lastRow = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).row
    
    For r = 2 To lastRow
        Dim s As String
        s = Trim(CStr(wsMain.Cells(r, 1).Value))
        If Left(s, 5) = "Block" Then
            Dim bIdx As Long
            bIdx = CLng(Trim(Replace(s, "Block", "")))
            If bIdx > currentBlock Then currentBlock = bIdx

            ' FIX: the LISP export does NOT always write Fx_lam/Fy_lam on the
            ' first coordinate row of a block - for some blocks it lands on
            ' the CLOSING (duplicate-of-first) vertex row instead. Capturing
            ' only "the very first row we see for this block" silently drops
            ' the load whenever it's written on a later row (this is exactly
            ' what happened to the single applied unit load in the sample
            ' model, which sat on Block 2's closing row) - the load then
            ' reads as 0, which makes the collapse-load LP unbounded and
            ' produces a meaningless lambda. Instead, scan EVERY row that
            ' belongs to this block and take the value from wherever it's
            ' actually non-blank.
            If bIdx >= 1 And bIdx <= NB Then
                If IsNumeric(wsMain.Cells(r, "L").Value) And wsMain.Cells(r, "L").Value <> "" Then
                    fx(bIdx) = CDbl(wsMain.Cells(r, "L").Value)
                End If
                If IsNumeric(wsMain.Cells(r, "M").Value) And wsMain.Cells(r, "M").Value <> "" Then
                    fy(bIdx) = CDbl(wsMain.Cells(r, "M").Value)
                End If
            End If
        End If
    Next r

    ' 3. Populate the Fass matrix
    fassRow = 2
    For i = 1 To NB
        Dim w As Double
        w = CDbl(wsMain.Cells(1 + i, "J").Value)
        
        ' F_ass_W (Constant Dead Loads)
        wsFass.Cells(fassRow, "A").Value = 0
        wsFass.Cells(fassRow + 1, "A").Value = w
        wsFass.Cells(fassRow + 2, "A").Value = 0

        ' F_ass_lam (Live Load Vectors drawn by user)
        wsFass.Cells(fassRow, "B").Value = fx(i)
        wsFass.Cells(fassRow + 1, "B").Value = -fy(i)
        wsFass.Cells(fassRow + 2, "B").Value = 0

        fassRow = fassRow + 3
    Next i

    wsFass.Columns("A:B").AutoFit
    wsFass.Range("A1:B1").Font.Bold = True
    wsFass.Range("A1:B1").HorizontalAlignment = xlCenter
    
    If fassRow > 2 Then
        With wsFass.Range("A2:B" & (fassRow - 1))
            .Borders.Weight = xlThin
        End With
    End If
End Sub

