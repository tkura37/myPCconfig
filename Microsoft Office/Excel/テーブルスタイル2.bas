Attribute VB_Name = "テーブルスタイル2"
Option Explicit

Sub ApplyCustomTableStyle2()
    Dim rng As Range
    On Error GoTo ErrorHandler
    Set rng = Selection

    Dim tbl As ListObject
    Set tbl = rng.Worksheet.ListObjects.Add(xlSrcRange, rng, , xlNo)
    
    tbl.TableStyle = "MyTableStyle2"
    
    tbl.ShowHeaders = False
    
    Exit Sub

ErrorHandler:
    MsgBox "セル範囲を選択してください。"
End Sub
