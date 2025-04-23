'選択中のセル範囲に罫線を設定する(見出し行の下が点線)
Option Explicit

Sub tableStyle2()
    Dim rng As Range
    Dim cell As Range
    Dim topRow As Long
    Dim bottomRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim borderColor As Long
    
    ' 線の色を設定
    borderColor = RGB(165, 165, 165)
    
    ' ユーザーが選択している範囲を取得
    On Error Resume Next
    Set rng = Application.Selection
    If rng Is Nothing Then
        MsgBox "セル範囲を選択してください。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 範囲の情報を取得
    topRow = rng.Row
    bottomRow = rng.Rows(rng.Rows.Count).Row
    leftCol = rng.Column
    rightCol = rng.Columns(rng.Columns.Count).Column

    ' すべてのセルの枠線をクリアしてから設定
    rng.Borders.LineStyle = xlNone
    
    ' 外枠を実線で設定
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = borderColor
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = borderColor
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = borderColor
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = borderColor
    End With
    
    ' 縦線（内側の縦）を実線で設定
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = borderColor
    End With
    
    ' 横線（内側の横）をすべて点線に設定
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Color = borderColor
    End With
    
End Sub
