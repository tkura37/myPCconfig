Option Explicit

Sub ApplyCustomCellStyle3()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(255, 242, 204)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
