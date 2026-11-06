Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim targetZoom As Long
    Dim summaryExists As Boolean
    Dim lastRow As Long
    Dim rng As Range

    ' Optimize performance
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    On Error GoTo ErrorHandler

    ' TM1 calculation off (skip if TM1 not loaded)
    On Error Resume Next
    Application.Run "TM1RECALCOFF"
    On Error GoTo ErrorHandler

    targetZoom = 110
    summaryExists = False

    For Each ws In ThisWorkbook.Worksheets
        ' Check for Summary sheet
        If ws.Name = "Summary" Then summaryExists = True

        With ws
            ' Unprotect first (in case already protected)
            .Unprotect Password:=""

            ' Clear any active filters but keep filters available
            If .AutoFilterMode Then
                If .FilterMode Then .ShowAllData
            End If

            ' Collapse all groups to level 1 (if any)
            On Error Resume Next
            .Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
            On Error GoTo ErrorHandler

            ' Select A1, reset scroll and zoom (visible sheets only: a hidden sheet cannot be activated)
            If .Visible = xlSheetVisible Then
                .Activate
                .Range("A1").Select
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
                ActiveWindow.View = xlNormalView
                ActiveWindow.Zoom = targetZoom
            End If

            ' Protect sheet with grouping and filtering allowed (Resume Next: Mac Excel rejects some options)
            On Error Resume Next
            .Protect Password:="", _
                DrawingObjects:=False, _
                Contents:=True, _
                Scenarios:=True, _
                UserInterfaceOnly:=True, _
                AllowFiltering:=True, _
                AllowUsingPivotTables:=True, _
                AllowFormattingColumns:=True, _
                AllowFormattingRows:=True

            ' Enable outlining (grouping) at workbook level
            ThisWorkbook.Worksheets(ws.Name).EnableOutlining = True
            On Error GoTo ErrorHandler
        End With
    Next ws

    ' Activate Summary sheet if it exists
    If summaryExists Then
        With ThisWorkbook.Sheets("Summary")
            .Activate
            .Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End With
    End If

Cleanup:
    ' Restore application settings
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

    ' Re-enable TM1 calculation
    On Error Resume Next
    Application.Run "TM1RECALCON"
    On Error GoTo 0

    ' Show welcome message
    ' MsgBox "Welcome Stuart D", vbInformation
    Exit Sub

ErrorHandler:
    Resume Cleanup
End Sub
