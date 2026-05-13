Attribute VB_Name = "各種クリア"
Option Explicit
Sub sh1クリア() '工程表
    
    Call SheetNames
    Dim i, j As Long
    
    With sh1
        For i = 0 To 5 '縦ループ
            For j = 0 To 1 '横ループ
                .Cells(2, 4).Offset(i * 35, j * 8).ClearContents
                .Cells(2, 5).Offset(i * 35, j * 8).MergeArea.ClearContents
                .Cells(3, 3).Offset(i * 35, j * 8).ClearContents
                .Cells(3, 4).Offset(i * 35, j * 8).ClearContents
                .Cells(3, 5).Offset(i * 35, j * 8).ClearContents
                .Cells(6, 4).Offset(i * 35, j * 8).Resize(30, 3).ClearContents
            Next j
        Next i
    End With
    
End Sub
Sub sh3クリア() '基本設定

    If MsgBox("入力内容をすべてクリアします" & vbCrLf & "続行しますか？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub

    Call ApplicationFalse
    Application.Calculation = xlCalculationManual

    Dim ClearRange As Range
    Dim BaseR As Long, num As Long
    Call SheetNames
    
    sh3.Range(Cells(3, 3), Cells(8, 4)).ClearContents
    sh3.Range(Cells(12, 3), Cells(23, 8)).ClearContents
    
    Application.Calculation = xlCalculationAutomatic
    Call ApplicationTrue

End Sub
Sub sh6クリア() '見積書

    Dim i As Long, j As Long, Row As Long
    Call SheetNames

    With sh6
    
        .Range("A5:A6").ClearContents
        .Range("F10:S11").ClearContents
        .Range("AA10:AL11").ClearContents
        .Range("A13:AL18").ClearContents
        
        For i = 0 To 5
        
            'データクリア
            .Range("A22:AL46").Offset(i * 48, 0).ClearContents
            .Range("Y47").Offset(i * 48, 0) = "=SUM(Y" & 22 + (48 * i) & ":AC" & 46 + (48 * i) & ")"
            .Range("Y48").Offset(i * 48, 0).MergeArea.ClearContents
            
            .Range("A23:AL23").Offset(48 * i, 0).Copy
            .Range("A22:AL22").Offset(48 * i, 0).PasteSpecial Paste:=xlPasteFormats
            
            Row = 48 * i + 22
            
            For j = 0 To 24
                
                '材料名列 左寄せ
                .Cells(22 + j, 1).Offset(48 * i, 0).HorizontalAlignment = xlHAlignLeft
           
                '金額列 数式記述
                .Cells(22 + j, 25).Offset(48 * i, 0) = "=IF(OR(M" & Row & "="""",U" _
                        & Row & "=""""),"""",ROUND(M" & Row & "*U" & Row & ",0))"
                Row = Row + 1
            Next
            
        Next
    
    End With

End Sub
Sub sh10クリア() '特価申請書

    Dim i As Long, n As Long
    Call SheetNames

    With sh10
        .Range("W2") = Date
        .Range("G12:AJ21").ClearContents
        .Range("L23:R24").ClearContents
        .Range("AD23:AJ23").ClearContents
        .Range("G29:AJ33").ClearContents
        .Range("G38:AJ38").ClearContents
        .Range("A41:AJ65").ClearContents
        .Range("A69:AJ73").ClearContents
        
        For i = 1 To 25
            n = 40 + i
            .Cells(n, 23) = "=IF(OR($K" & n & "="""",$S" & n & "=""""),"""",$K" & n & "*$S" & n & ")"
            .Cells(n, 32) = "=IF(OR($K" & n & "="""",$AB" & n & "=""""),"""",$K" & n & "*$AB" & n & ")"
        Next
    
    End With

End Sub
