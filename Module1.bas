Attribute VB_Name = "Module1"
Sub Generate_Gass_in()
    Dim srcSheet As Worksheet, tgtSheet As Worksheet
    Dim wbUser As Workbook
    Dim lastRow As Long, i As Long, m As Long, matrixCount As Long
    Dim t As Double

    Set wbUser = ActiveWorkbook
    Set srcSheet = wbUser.ActiveSheet

    On Error Resume Next
    Set tgtSheet = wbUser.Worksheets("Gass_in")
    On Error GoTo 0
    If tgtSheet Is Nothing Then
        Set tgtSheet = wbUser.Worksheets.Add
        tgtSheet.Name = "Gass_in"
    Else
        tgtSheet.Cells.Clear
    End If

    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "F").End(xlUp).row
    Dim TValues() As Double
    ReDim TValues(1 To 1)
    matrixCount = 0

    For i = 2 To lastRow
        If IsEmpty(srcSheet.Cells(i, "F").Value) Then Exit For
        If Not IsNumeric(srcSheet.Cells(i, "F").Value) Then Exit For
        matrixCount = matrixCount + 1
        ReDim Preserve TValues(1 To matrixCount)
        TValues(matrixCount) = CDbl(srcSheet.Cells(i, "F").Value)
    Next i

    If matrixCount = 0 Then Exit Sub

    Dim totalRows As Long, totalCols As Long
    totalRows = matrixCount * 2
    totalCols = matrixCount * 3
    tgtSheet.Range(tgtSheet.Cells(1, 1), tgtSheet.Cells(totalRows, totalCols)).Value = 0

    For m = 1 To matrixCount
        t = TValues(m)
        Dim r0 As Long, c0 As Long
        r0 = (m - 1) * 2 + 1
        c0 = (m - 1) * 3 + 1

        tgtSheet.Cells(r0, c0).Value = t / 2
        tgtSheet.Cells(r0, c0 + 1).Value = 0
        tgtSheet.Cells(r0, c0 + 2).Value = 1
        tgtSheet.Cells(r0 + 1, c0).Value = t / 2
        tgtSheet.Cells(r0 + 1, c0 + 1).Value = 0
        tgtSheet.Cells(r0 + 1, c0 + 2).Value = -1
    Next m
End Sub
