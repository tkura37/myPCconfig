Option Explicit

Sub ApplyCustomCellStyle4()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(226, 239, 218)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
