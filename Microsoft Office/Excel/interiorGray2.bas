'選択中のセル範囲の色を濃い灰色にする
Option Explicit

Sub interiorGray2()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(165, 165, 165)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
