Attribute VB_Name = "Sub_Program"
Option Explicit
'反映ボタンが押されたときはこちらを実行
Sub OptionUpdate()

    If MsgBox("設定をデータベースに反映します。よろしいですか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Call UserOption
    MsgBox "設定を反映しました", vbInformation + vbOKOnly
    
End Sub
'ユーザー設定の反映
Sub UserOption() '2019/08/14 Update by Yamashita

    Call ApplicationFalse
    Application.Calculation = xlCalculationManual

    Call SheetNames
    Dim Col1 As Long, Col2 As Long, MaxOption As Long
    Dim Col3 As Long, Col4 As Long, Col5 As Long, MaxDB As Long, myNum As Long
    Dim Item As String, Quantity As Double
    
    Dim Target As Long, i As Long, j As Long
    
    '設定元
    Col1 = FindColumn("材料名", sh3, 3)
    Col2 = FindColumn("使用量", sh3, 3)
    MaxOption = MaxRow(sh3, FindColumn("項目", sh3, 3)) - 3
    
    'データ元(反映先)
    Col3 = FindColumn("ユーザー設定番号", sh2, 1)
    Col4 = FindColumn("材料名", sh2, 1)
    Col5 = FindColumn("使用量", sh2, 1)
    MaxDB = MaxRow(sh2, 1) - 1

    For i = 1 To MaxOption 'オプション数ループ
        '材料名と使用量をセット
        Item = sh3.Cells(3 + i, Col1)
        Quantity = sh3.Cells(3 + i, Col2)
    
        For j = 1 To MaxDB '工程DBデータ数ループ
            Target = sh2.Cells(1 + j, Col3).Value
            
            If Target = i Then
                If Item <> "" Then sh2.Cells(1 + j, Col4) = Item
                If Quantity <> 0 Then sh2.Cells(1 + j, Col5) = Quantity
            End If
            
        Next j
    Next i

    Application.Calculation = xlCalculationAutomatic
    Call ApplicationTrue

End Sub
'荷姿表示
Sub PackingUnit(ByVal Page As Long) '2019/08/14 Update by Yamashita

    Call SheetNames
    Dim TargetRange, UnionRange As Range
    Dim i As Long
    Dim obj As Object
    Dim ItemName As String
    
    '処理範囲を指定
    For i = 1 To Page
    
        '追加対象Range
        Set TargetRange = sh6.Range("A23:A46").Offset(48 * (i - 1), 0)
            
        'ターゲットを処理範囲に追加
        If UnionRange Is Nothing Then
            Set UnionRange = TargetRange
        Else
            Set UnionRange = Union(UnionRange, TargetRange)
        End If
        
    Next i

    '処理範囲に対して荷姿を表示
    On Error Resume Next
    For Each obj In UnionRange
    
        If obj <> "" Then
            ItemName = obj.Value
            obj.Offset(0, 18) = WorksheetFunction.XLookup(ItemName, sh5.ListObjects("材料DB").ListColumns("材料名").DataBodyRange, _
                                                                sh5.ListObjects("材料DB").ListColumns("荷姿 / 単位").DataBodyRange, "")
        End If

    Next obj

End Sub
'マスタ更新
Sub UpdateMaster()

    'ネットワーク接続を確認
    If CheckNetwork = False Then
        MsgBox "ネットワーク接続がないため更新できません。", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    Call ApplicationFalse
    Application.Calculation = xlCalculationManual
    
    Call SheetNames

    Dim MstWB As Workbook
    Dim CopySheet As Worksheet, PasteSheet As Worksheet
    Dim Path As String, FileName As String, VerSheetName As String, TargetSheet As String, UpdateM As String
    Dim PathRow As Long, FileNameRow As Long, VerRow As Long, Version As Long
    Dim NowVer As Long, NowVerRow As Long, SheetNum As Long, i As Long, MessageRows As Long
    Dim VisibleFlag As Boolean
   
    '更新用ファイルパス
    PathRow = FindRow("更新用データ 保存先パス", sh8, 1)
    FileNameRow = FindRow("更新用データ ファイル名", sh8, 1)
    VerRow = FindRow("バージョン情報シート", sh8, 1)
    NowVerRow = FindRow("現バージョン", sh8, 1)
    
    Path = sh8.Cells(PathRow, 2)
    FileName = sh8.Cells(FileNameRow, 2)
    VerSheetName = sh8.Cells(VerRow, 2)
    NowVer = sh8.Cells(NowVerRow, 2)
    
    '更新用ファイルを開いてセット
    Set MstWB = Workbooks.Open(Path & FileName)

    'マスタバージョンのチェック
    Version = MstWB.Worksheets(VerSheetName).Cells(1, 1) 'バージョン
    If Version > NowVer Then '最新バージョンがある場合
    
        If MsgBox("最新版のマスタがあります" & vbCrLf & "更新しますか？", vbYesNo + vbQuestion) = vbNo Then
            MsgBox "キャンセルしました"
            MstWB.Close
            sh3.Select
            Call ApplicationTrue
            Exit Sub
        End If

        '更新処理
        SheetNum = MaxRow(sh8, 4)
    
        For i = 1 To SheetNum
            VisibleFlag = False
            TargetSheet = sh8.Cells(i, 4)
            'コピー元のシートを指定
            Set CopySheet = MstWB.Worksheets(TargetSheet)
            '貼り付け先のシートを指定
            Set PasteSheet = WB.Worksheets(TargetSheet)
            '非表示なら再表示
            If PasteSheet.Visible = False Then
                PasteSheet.Visible = True
                VisibleFlag = True
            End If
            
            On Error Resume Next
            'テーブルが存在しない場合
            If CopySheet.ListObjects.Count = 0 Then
                PasteSheet.Cells.Clear '全削除
                CopySheet.Cells.Copy
                PasteSheet.Cells.PasteSpecial Paste:=xlPasteAll
                
            'テーブルが存在する場合
            Else
                'コピーするテーブルのデータ部をコピー
                CopySheet.ListObjects(1).DataBodyRange.Copy
                '貼り付け
                PasteSheet.Cells(2, 1).PasteSpecial Paste:=xlPasteValues '値
                PasteSheet.Cells(2, 1).PasteSpecial Paste:=xlPasteFormats '書式
                
                'コピーするテーブルの見出し部をコピー
                CopySheet.ListObjects(1).HeaderRowRange.Copy
                '貼り付け
                PasteSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues '値
                PasteSheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats '書式

            End If
            On Error GoTo 0
            
            '非表示シートだった場合再度非表示にする
            If VisibleFlag = True Then
                PasteSheet.Visible = False
            End If
            
        Next i
            
        'バージョン情報の更新
        sh8.Cells(NowVerRow, 2) = Version
        
        'アップデート内容取得
        With MstWB.Worksheets(VerSheetName)
            MessageRows = MaxRow(MstWB.Worksheets(VerSheetName), 2)
            For i = 1 To MessageRows
                If UpdateM = "" Then
                    UpdateM = .Cells(i, 2).Value & vbCrLf
                Else
                    UpdateM = UpdateM & vbCrLf & .Cells(i, 2).Value
                End If
            Next i
        End With
        
        MstWB.Close
        sh3.Select
        
        '工程DBのユーザー設定反映
        Call UserOption
        
        MsgBox "更新が完了しました", vbInformation + vbOKOnly
        
        '更新内容の表示
        MsgBox UpdateM, vbOKOnly
        
    '最新バージョンがない場合
    Else
        MstWB.Close
        sh3.Select
        MsgBox "ご使用のマスタは最新版です", vbInformation + vbOKOnly

    End If
    
    Application.Calculation = xlCalculationAutomatic
    Call ApplicationTrue
  
End Sub
'ネットワーク接続のチェック(FHSドライブがあるか確認)して True or False を返す
Function CheckNetwork() As Boolean

    Call SheetNames
    Dim Path As String
    Dim PathRow As Long
    
    PathRow = FindRow("ネットワーク確認用パス", sh8, 1)
    Path = sh8.Cells(PathRow, 2)

    If CreateObject("Scripting.FileSystemObject").DriveExists(Path) Then
        CheckNetwork = True
    Else
        CheckNetwork = False
    End If
    
End Function
'見積書のインポート
Sub Import()

    If MsgBox("過去に作成した見積書データを読み込みます。" & vbCrLf & "※現在のデータは破棄されます。", vbYesNo + vbExclamation) = vbNo Then Exit Sub

    Call SheetNames
    Call ApplicationFalse
    Application.Calculation = xlCalculationManual
    
    Dim Imbook As Workbook
    Dim OpenFileName As String
    Dim LastRow As Long, top As Long, bot As Long, n As Long, i As Long, Pattern As Long
    Dim Page As Double
    Dim UnionRange As Range, TargetRange As Range, BaseRange As Range
    Dim obj As Object

    ChDrive WB.Path
    ChDir WB.Path

    OpenFileName = Application.GetOpenFilename("Excel ファイル (*.xls; *.xlsx; *.xlsm),*.xls; *.xlsx; *.xlsm")

    '見積書を開いてセット
    If OpenFileName <> "False" Then
        Set Imbook = Workbooks.Open(OpenFileName)
    Else
        Exit Sub
    End If

    With Imbook.ActiveSheet

        'パスワード解除
        On Error Resume Next
        .Unprotect Password:="ssk"
        On Error GoTo 0
            
        'ページ数計算＆行数チェック
        LastRow = .Cells(1, 1).SpecialCells(xlLastCell).Row
        Page = LastRow / 48
        
        If Page <> Int(Page) Then
            MsgBox "行数が変更された見積書のため、読み込む事ができません。", vbOKOnly + vbCritical
            Application.DisplayAlerts = False
            Imbook.Close
            Application.DisplayAlerts = True
            Exit Sub
        End If

        '書式がどちらか調べる
        '検査対象範囲(UnionRange)を設定
        Set BaseRange = .Range("A22:A45")
        For i = 1 To Page
            Set TargetRange = BaseRange.Offset(48 * (i - 1), 0)
            If UnionRange Is Nothing Then
                Set UnionRange = TargetRange
            Else
                Set UnionRange = Union(UnionRange, TargetRange)
            End If
        Next i

        Pattern = 2 '(ページ合計)
        For Each obj In UnionRange
            If obj.Value = "小計" Then
                Pattern = 1 '(部位別集計)
                Exit For
            End If
        Next obj

        Call sh6クリア

        '基本項目転写
        .Rows(1).Insert '1行追加(行を揃えるため)
        sh6.Range("AO11") = Page
        For i = 1 To 2
            sh6.Cells(4 + i, 1) = .Cells(4 + i, 1)
            sh6.Cells(9 + i, 6) = .Cells(9 + i, 6)
            sh6.Cells(9 + i, 27) = .Cells(9 + i, 27)
        Next i
        
        For i = 1 To 6
            sh6.Cells(12 + i, 1) = .Cells(12 + i, 1)
            sh6.Cells(12 + i, 10) = .Cells(12 + i, 10)
            sh6.Cells(12 + i, 15) = .Cells(12 + i, 15)
            sh6.Cells(12 + i, 20) = .Cells(12 + i, 20)
            sh6.Cells(12 + i, 29) = .Cells(12 + i, 29)
            sh6.Cells(12 + i, 34) = .Cells(12 + i, 34)
        Next i

        '見積部転写
        For i = 1 To Page
            .Range("A22:AL46").Offset((i - 1) * 48, 0).Copy
            sh6.Range("A22:AL46").Offset((i - 1) * 48, 0).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                                                                                                       SkipBlanks:=False, Transpose:=False
        Next i

    End With ' Imbook.ActiveSheet

    Application.DisplayAlerts = False
    Imbook.Close
    Application.DisplayAlerts = True
    
    'ページ合計数式を反映
    For i = 0 To 5
        If Pattern = 2 Then
            sh6.Range("Y47").Offset(i * 48, 0) = "=SUM(Y" & 22 + (48 * i) & ":AC" & 46 + (48 * i) & ")"
        Else
            sh6.Range("Y47").Offset(i * 48, 0) = "=SUMIF(A" & 22 + (i * 48) & ":L" & 46 + (i * 48) & _
                                                            "," & """小計""" & ",Y" & 22 + (i * 48) & ":AC" & 46 + (i * 48) & " )"
        End If
    Next i
    
    '見積総額の数式を反映
    sh6.Range("G8") = "=OFFSET(Y49,48*(AO11-1),0)"
    
    Application.Calculation = xlCalculationAutomatic
    Call ApplicationTrue
    
    MsgBox "読み込みが完了しました。", vbOKOnly + vbInformation
    
End Sub
