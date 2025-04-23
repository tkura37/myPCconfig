'選択中の文字の色を黒にする
Option Explicit

Sub textBlack()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorBlack
    Else
        MsgBox "文字列を選択してください。", vbExclamation
    End If
End Sub
