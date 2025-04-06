Attribute VB_Name = "薄い水色"
Option Explicit

Sub ApplyCustomCellStyle2()
    On Error GoTo ErrorHandler
    Selection.Style = "薄い水色"
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
