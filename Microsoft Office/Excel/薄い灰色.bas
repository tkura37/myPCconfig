Option Explicit

Sub ApplyCustomCellStyle1()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(242, 242, 242)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
