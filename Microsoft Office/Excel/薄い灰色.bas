Attribute VB_Name = "薄い灰色"
Option Explicit

Sub ApplyCustomCellStyle1()
    On Error GoTo ErrorHandler
    Selection.Style = "薄い灰色"
    Exit Sub

ErrorHandler:
    MsgBox "セルスタイルの適用に失敗しました。スタイル名を確認してください。"
End Sub
