Option Explicit

' Tasks list layout: input columns B:J, task rows 7 to 1006
Private Const FIRST_ROW As Long = 7
Private Const LAST_ROW As Long = 1006
Private Const N_COLS As Long = 9

' RESET FILTERS button (Board): team and project filters back to All
Public Sub ResetBoardFilters()
    With ThisWorkbook.Worksheets("Board")
        .Range("B7").Value = "All"
        .Range("D7").Value = "All"
    End With
End Sub

' ARCHIVE DONE TASKS button (Tasks): move Done tasks to the hidden Archive sheet and close up the list
Public Sub ArchiveDoneTasks()
    Dim tasks As Worksheet, arc As Worksheet
    Dim data As Variant, ids As Variant, keep() As Variant, moved() As Variant
    Dim doneName As String
    Dim r As Long, c As Long, nDone As Long, nKeep As Long, nextRow As Long
    Dim hasContent As Boolean

    Set tasks = ThisWorkbook.Worksheets("Tasks")
    Set arc = ThisWorkbook.Worksheets("Archive")
    doneName = ThisWorkbook.Worksheets("Lists").Range("A6").Value
    data = tasks.Range(tasks.Cells(FIRST_ROW, 2), tasks.Cells(LAST_ROW, 1 + N_COLS)).Value
    ids = tasks.Range(tasks.Cells(FIRST_ROW, 1), tasks.Cells(LAST_ROW, 1)).Value

    For r = 1 To UBound(data, 1)
        If Len(CStr(data(r, 1))) > 0 And CStr(data(r, 3)) = doneName Then nDone = nDone + 1
    Next r
    If nDone = 0 Then
        MsgBox "There are no " & doneName & " tasks to archive.", vbInformation, "Archive"
        Exit Sub
    End If
    If MsgBox("Archive " & nDone & " " & doneName & " task(s)?" & vbCrLf & vbCrLf & _
              "They move to the hidden Archive sheet and the task list closes up.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Are you sure?") <> vbYes Then Exit Sub

    ' split rows: Done tasks to the archive, everything else kept in order with gaps removed
    ReDim keep(1 To UBound(data, 1), 1 To N_COLS)
    ReDim moved(1 To nDone, 1 To N_COLS + 2)
    nDone = 0
    For r = 1 To UBound(data, 1)
        If Len(CStr(data(r, 1))) > 0 And CStr(data(r, 3)) = doneName Then
            nDone = nDone + 1
            moved(nDone, 1) = Date
            moved(nDone, 2) = ids(r, 1)
            For c = 1 To N_COLS
                moved(nDone, c + 2) = data(r, c)
            Next c
        Else
            hasContent = False
            For c = 1 To N_COLS
                If Len(CStr(data(r, c))) > 0 Then hasContent = True
            Next c
            If hasContent Then
                nKeep = nKeep + 1
                For c = 1 To N_COLS
                    keep(nKeep, c) = data(r, c)
                Next c
            End If
        End If
    Next r

    Application.ScreenUpdating = False
    nextRow = arc.Cells(arc.Rows.Count, "C").End(xlUp).Row + 1
    With arc.Cells(nextRow, 1).Resize(nDone, N_COLS + 2)
        .Value = moved
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Arial"
        .Font.Size = 10
    End With
    arc.Cells(nextRow, 1).Resize(nDone, 1).NumberFormat = "d mmm yy"
    arc.Cells(nextRow, 8).Resize(nDone, 3).NumberFormat = "d mmm yy"
    tasks.Range(tasks.Cells(FIRST_ROW, 2), tasks.Cells(LAST_ROW, 1 + N_COLS)).Value = keep
    Application.ScreenUpdating = True

    MsgBox nDone & " task(s) archived.", vbInformation, "Archive"
End Sub
