'選択中のセル範囲の色を薄い灰色にする
Option Explicit

Sub interiorGray1()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(242, 242, 242)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
