Option Explicit

Sub ApplyCustomCellStyle2()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(221, 235, 247)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
