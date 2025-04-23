'選択中のセル範囲の色を薄い緑色にする
Option Explicit

Sub interiorLightgreen()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(226, 239, 218)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
