Attribute VB_Name = "テーブルスタイル1"
Option Explicit

Sub ApplyCustomTableStyle1()
    Dim rng As Range
    On Error GoTo ErrorHandler
    Set rng = Selection

    Dim tbl As ListObject
    Set tbl = rng.Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    
    tbl.TableStyle = "MyTableStyle1"
    
    Exit Sub

ErrorHandler:
    MsgBox "セル範囲を選択してください。"
End Sub
