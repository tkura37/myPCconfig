Attribute VB_Name = "薄い緑色"
Option Explicit

Sub ApplyCustomCellStyle4()
    On Error GoTo ErrorHandler
    Selection.Style = "薄い緑色"
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
