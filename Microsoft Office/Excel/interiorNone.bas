'選択中のセル範囲の塗りつぶしを解除する
Option Explicit

Sub interiorNone()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = xlNone
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
