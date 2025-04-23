'選択中の文字の色を赤にする
Option Explicit

Sub textRed()
    If Selection.Type <> wdNoSelection Then
        Selection.Font.Color = wdColorRed
    Else
        MsgBox "文字列を選択してください。", vbExclamation
    End If
End Sub
