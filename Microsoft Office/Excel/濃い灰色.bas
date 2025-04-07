Option Explicit

Sub ApplyCustomCellStyle5()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(165, 165, 165)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
