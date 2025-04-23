'選択中のセル範囲の色を薄い黄色にする
Option Explicit

Sub interiorLightyellow()
    On Error GoTo ErrorHandler
    Selection.Interior.Color = RGB(255, 242, 204)
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
