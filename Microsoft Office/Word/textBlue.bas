'選択中の文字の色を青にする
Option Explicit

Sub textBlue()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorBlue
    Else
        MsgBox "文字列を選択してください。", vbExclamation
    End If
End Sub
