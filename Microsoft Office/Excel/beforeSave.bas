'全てのシートをA1選択・拡大率100%にし、最初のシートを選択する(保存して閉じる前に使う)
Option Explicit

Sub beforeSave()
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Cells(1, 1).Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        ActiveWindow.Zoom = 100
    Next ws

    Sheets(1).Select
    Application.ScreenUpdating = True
End Sub
