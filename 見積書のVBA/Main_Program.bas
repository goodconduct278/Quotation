Attribute VB_Name = "Main_Program"
Option Explicit
Sub 工程読込() '2019/08/09 Update by Yamashita

    If MsgBox("基本仕様を読み込み、編集画面へ移行します。よろしいですか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Call SheetNames
    Call ApplicationFalse
    Application.Calculation = xlCalculationManual
    Call sh1クリア

    Dim i As Long, myCount As Long, PageNum As Long, PartNum As Long, CopyCol As Long, LastRow As Long
    Dim RowNum() As Long, SpecRow As Long, SpecNumCol As Long, OffsetCol As Long, OffsetRow As Long
    Dim SpecNum() As String, errorMsg As String
    Dim CopyArea As Range, PasteArea As Range

'-----------------------------------------------------------------------------
'基本設定読み込み

    With sh3
    
        '------------------------エラーチェック-----------------------
        ' 指定範囲（B12:H23）で入力されているセルの最終行を取得
        LastRow = .Range("B12:H23").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        ' 範囲内にデータがない場合は終了
        If LastRow < 12 Then
            MsgBox "指定範囲内にデータがありません。", vbExclamation, "エラー"
            Exit Sub
        End If
        
        ' B列をクリアする処理（C列～H列が空欄の場合）
        For i = 12 To LastRow
            If Not IsEmpty(.Cells(i, "B").Value) And _
               IsEmpty(.Cells(i, "C").Value) And _
               IsEmpty(.Cells(i, "D").Value) And _
               IsEmpty(.Cells(i, "E").Value) And _
               IsEmpty(.Cells(i, "F").Value) And _
               IsEmpty(.Cells(i, "G").Value) And _
               IsEmpty(.Cells(i, "H").Value) Then
                ' B列の値をクリア
                .Cells(i, "B").ClearContents
            End If
        Next i

        ' 指定範囲（B12:H23）で入力されているセルの最終行を再取得
        LastRow = .Range("B12:H23").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        ' エラーメッセージの初期化
        errorMsg = ""
        
        ' 入力必須列の入力チェック（11行目から最終行までループ）
        For i = 11 To LastRow
            ' B列、F列、G列、H列のいずれかが未入力の場合
            If IsEmpty(.Cells(i, "B").Value) Or _
               IsEmpty(.Cells(i, "C").Value) Or _
               IsEmpty(.Cells(i, "F").Value) Or _
               IsEmpty(.Cells(i, "G").Value) Or _
               IsEmpty(.Cells(i, "H").Value) Then
                
                ' エラーメッセージに未入力行を追加
                errorMsg = errorMsg & "行 " & i & " の入力必須項目が未入力です。" & vbCrLf
            End If
        Next i
        
        ' エラーメッセージがある場合、表示
        If errorMsg <> "" Then
            MsgBox "以下の行にエラーがあります：" & vbCrLf & errorMsg, vbExclamation, "入力エラー"
            Exit Sub
        End If
    
        '------------------------エラーチェックここまで-----------------------
    
        PageNum = WorksheetFunction.Max(.Range("B12:B23")) '使用ページ数
        PartNum = WorksheetFunction.CountA(.Range("F12:F23")) '部位数
        ReDim SpecNum(1 To PartNum)
        ReDim RowNum(1 To PartNum)
        For i = 1 To 12
            If .Range("F11").Offset(i, 0) <> "" Then
            myCount = myCount + 1
            SpecNum(myCount) = .Range("F11").Offset(i, 0) '仕様番号
            RowNum(myCount) = i
            End If
        Next
    End With

'-----------------------------------------------------------------------------
'シート「工程表」に各種データを読み込み

    With sh1
        For i = 1 To PartNum
            If i Mod 2 = 1 Then '奇数の場合
                OffsetCol = 0
                OffsetRow = 35 * Fix(i / 2)
            ElseIf i Mod 2 = 0 Then '偶数の場合
                OffsetCol = 8
                OffsetRow = 35 * (Fix(i / 2) - 1)
            End If
            
            .Cells(2, 4).Offset(OffsetRow, OffsetCol) = SpecNum(i)  '仕様番号
            .Cells(3, 3).Offset(OffsetRow, OffsetCol) = sh3.Cells(11, 3).Offset(RowNum(i), 0) '部位
            .Cells(2, 5).Offset(OffsetRow, OffsetCol) = sh3.Cells(11, 4).Offset(RowNum(i), 0) '部位名
            .Cells(3, 4).Offset(OffsetRow, OffsetCol) = sh3.Cells(11, 7).Offset(RowNum(i), 0) '面積
            .Cells(3, 5).Offset(OffsetRow, OffsetCol) = sh3.Cells(11, 8).Offset(RowNum(i), 0) '外周
            
            SpecNumCol = FindColumn("計算用", sh2, 1)
            CopyCol = FindColumn("材料種類", sh2, 1)
            SpecRow = FindRow(SpecNum(i) & sh3.Cells(11, 3).Offset(RowNum(i), 0), sh2, SpecNumCol)
            Set CopyArea = sh2.Cells(SpecRow, CopyCol).Resize(30, 3)
            Set PasteArea = .Cells(6, 4).Offset(OffsetRow, OffsetCol).Resize(30, 3)
            PasteArea.Value = CopyArea.Value

        Next
    End With

    sh1.Select

    Call ApplicationTrue
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub 見積書反映() '2019/08/13 Update by Yamashita

    If MsgBox("入力内容で見積書を作成します。よろしいですか？", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Call SheetNames
    Call ApplicationFalse
    Application.Calculation = xlCalculationManual
    
    Dim Dic As Object
    Dim i As Long, n As Long, MaxPage As Long, PartNum As Long, myCount As Long
    Dim ReferenceRow As Long, SubtotalRow As Long, Target As Long
    Dim PageNum() As Long, SpecCount() As Long
    Dim PartName(), PartText As String
    Dim ResultArray

    Call sh6クリア

'-----------------------------------------------------------------------------
'基本設定読み込み&反映
    
    With sh3
    
        MaxPage = WorksheetFunction.Max(.Range("B12:B23")) '使用ページ数
        PartNum = WorksheetFunction.CountA(.Range("F12:F23")) '部位数
        
        sh6.Cells(5, 1) = .Cells(3, 3) '宛名1
        sh6.Cells(6, 1) = .Cells(4, 3) '宛名1
        sh6.Cells(10, 6) = .Cells(5, 3) '工事店
        sh6.Cells(11, 6) = .Cells(6, 3) '物件名
        sh6.Cells(10, 27) = .Cells(7, 3) '依頼年月日
        sh6.Cells(11, 27) = .Cells(8, 3) '元請
        
        ReDim PageNum(1 To PartNum)
        ReDim PartName(1 To PartNum)
        ReDim SpecCount(1 To PartNum)

        myCount = 0
        For i = 1 To 12
            If .Range("F11").Offset(i, 0) <> "" Then
                myCount = myCount + 1
                PageNum(myCount) = .Range("F11").Offset(i, -4) 'ページ番号
                PartName(myCount) = .Range("F11").Offset(i, -2) '部位
                If myCount < 7 Then
                    sh6.Cells(12 + myCount, 1) = PartName(myCount)  '部位
                    sh6.Cells(12 + myCount, 10) = .Range("F11").Offset(i, 0) '仕様
                    sh6.Cells(12 + myCount, 15) = .Range("F11").Offset(i, 1) '面積
                Else
                    sh6.Cells(6 + myCount, 20) = PartName(myCount) '部位
                    sh6.Cells(6 + myCount, 29) = .Range("F11").Offset(i, 0) '仕様
                    sh6.Cells(6 + myCount, 34) = .Range("F11").Offset(i, 1) '面積
                End If
            End If
        Next i
        
    End With

'-----------------------------------------------------------------------------
'使用量を計算し、見積書に反映(部位別集計の場合)

    If sh3.OLEObjects("OptionButton2").Object.Value = True Then

        For n = 1 To PartNum '部位数ループ
        
            '関数呼び出し
            ResultArray = 使用量計算(n)
            
            With sh6

                '各ページの出力済み仕様数をSpecCount()に格納
                ReferenceRow = 22 + 48 * (PageNum(n) - 1) '出力行(-1)
                If SpecCount(PageNum(n)) = 0 Then
                    '出力エラーメッセージ
                    If UBound(ResultArray, 1) > 23 Then
                        MsgBox "製品数が多いため見積書に収まりません。処理を中止します。", vbCritical + vbOKOnly
                        Exit Sub
                    End If
                Else
                    Target = WorksheetFunction.Match("小計", .Range(.Cells(ReferenceRow + 1, 1), .Cells(ReferenceRow + 24, 1)), 0)
                    ReferenceRow = ReferenceRow + Target + 2
                    '出力エラーメッセージ
                    If ReferenceRow + UBound(ResultArray, 1) > 45 + 48 * (PageNum(n) - 1) Then
                        MsgBox "製品数が多いため見積書に収まりません。処理を中止します。", vbCritical + vbOKOnly
                        Exit Sub
                    End If
                End If

                For i = 1 To UBound(ResultArray, 1)
                    .Cells(ReferenceRow + i, 1) = ResultArray(i, 1) '材料
                    .Cells(ReferenceRow + i, 13) = ResultArray(i, 2) '数量
                    .Cells(ReferenceRow + i, 17) = WorksheetFunction.XLookup(ResultArray(i, 1), sh5.ListObjects("材料DB").ListColumns("材料名").DataBodyRange, _
                                                                    sh5.ListObjects("材料DB").ListColumns("荷姿単位").DataBodyRange, "") '単位
                Next i
                .Cells(ReferenceRow, 1) = "≪" & PartName(n) & "≫" '部位名
                SubtotalRow = ReferenceRow + UBound(ResultArray, 1) + 1
                .Cells(SubtotalRow, 1) = "小計"
                .Cells(SubtotalRow, 1).HorizontalAlignment = xlRight '右寄せ
                .Cells(SubtotalRow, 25) = "=SUM(Y" & ReferenceRow + 1 & ":AC" & SubtotalRow - 1 & ")" '小計数式
                
                'ページ小計の数式を全ページ書き換え
                For i = 0 To 5
                    .Range("Y47").Offset(i * 48, 0) = "=SUMIF(A" & 22 + (i * 48) & ":L" & 46 + (i * 48) & _
                                                                          "," & """小計""" & ",Y" & 22 + (i * 48) & ":AC" & 46 + (i * 48) & " )"
                Next i
                
            End With
            
            '出力済み仕様数を+1
            SpecCount(PageNum(n)) = SpecCount(PageNum(n)) + 1
            
        Next n
        
    '部位別集計の場合END
    
'-----------------------------------------------------------------------------
'使用量を計算し、見積書に反映(ページ内合算の場合)
    
    ElseIf sh3.OLEObjects("OptionButton1").Object.Value = True Then
        For n = 1 To MaxPage 'ページ数ループ
    
            '関数呼び出し
            ResultArray = 使用量計算2(PageNum(), n)

            With sh6

                ReferenceRow = 22 + 48 * (n - 1) '出力行(-1)
                '出力エラーメッセージ
                If UBound(ResultArray, 1) > 24 Then
                    MsgBox "製品数が多いため見積書に収まりません。処理を中止します。", vbCritical + vbOKOnly
                    Exit Sub
                End If
                
                For i = 1 To UBound(ResultArray, 1)
                    .Cells(ReferenceRow + i, 1) = ResultArray(i, 1) '材料
                    .Cells(ReferenceRow + i, 13) = ResultArray(i, 2) '数量
                    .Cells(ReferenceRow + i, 17) = WorksheetFunction.XLookup(ResultArray(i, 1), sh5.ListObjects("材料DB").ListColumns("材料名").DataBodyRange, _
                                                                    sh5.ListObjects("材料DB").ListColumns("荷姿単位").DataBodyRange, "") '単位
                Next i

                PartText = ""
                myCount = 0
                For i = 1 To UBound(PartName) '部位数繰り返し
                    If PageNum(i) = n Then
                        myCount = myCount + 1
                        If myCount > 1 Then PartText = PartText + "/"
                        PartText = PartText + PartName(i)
                    End If
                Next i
                .Cells(ReferenceRow, 1) = "≪" & PartText & "≫" '部位名
                .Cells(ReferenceRow, 1).Resize(1, 29).Merge
            End With

        Next n

    'ページ内合算の場合END

    End If
    
'-----------------------------------------------------------------------------
'後処理

    sh6.Range("AO11") = MaxPage
    sh6.Select
    
    '荷姿表示処理
    If sh3.OLEObjects("CheckBox2").Object.Value = True Then Call PackingUnit(MaxPage)
    
    Call ApplicationTrue
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "作成が完了しました", vbInformation + vbOKOnly

End Sub
'2019/08/14 Update by Yamashita
'指定の部位番号の材料・使用量を計算して配列を返す(部位別集計の場合)
Function 使用量計算(ByVal n As Long) As Variant()
    
    Dim Dic As Object
    Dim i As Long, OffsetCol As Long, OffsetRow As Long
    Dim Coefficient As Double
    Dim buf As String
    Dim TakeArray() As Variant
    Dim Keys, Items

'-----------------------------------------------------------------------------
'重複を削除した商品と使用量のリストを作成

    '読み込む工程表(n)の位置をオフセットに格納
    If n Mod 2 = 1 Then '奇数の場合
        OffsetCol = 0
        OffsetRow = 35 * Fix(n / 2)
    ElseIf n Mod 2 = 0 Then '偶数の場合
        OffsetCol = 8
        OffsetRow = 35 * (Fix(n / 2) - 1)
    End If

    Set Dic = CreateObject("Scripting.Dictionary")
    For i = 1 To 30
        buf = sh1.Cells(5 + i, 5).Offset(OffsetRow, OffsetCol).Value
        If buf = "" Then GoTo Continue
        If Not Dic.Exists(buf) Then '新たなKeyの場合、KeyとItemを登録
            Dic.Add buf, sh1.Cells(5 + i, 8).Offset(OffsetRow, OffsetCol)
        Else '既に存在するKeyの場合、Itemを加算
            Dic(buf) = Dic(buf) + sh1.Cells(5 + i, 8).Offset(OffsetRow, OffsetCol)
        End If
Continue:
    Next i
        
    Keys = Dic.Keys
    Items = Dic.Items
    
'-----------------------------------------------------------------------------
'必要数量を計算   Keys:商品、Items:必要数量

    ReDim TakeArray(1 To UBound(Keys) + 1, 1 To 2)
    On Error Resume Next
    For i = 0 To UBound(Keys)
        Coefficient = WorksheetFunction.XLookup(Keys(i), _
            sh5.ListObjects("材料DB").ListColumns("材料名").DataBodyRange, _
            sh5.ListObjects("材料DB").ListColumns("係数").DataBodyRange, "")
        Items(i) = WorksheetFunction.RoundUp(Items(i) / Coefficient, 0)
        '戻り値の配列に格納(Keys, Items)
        TakeArray(i + 1, 1) = Keys(i)
        TakeArray(i + 1, 2) = Items(i)
    Next

    使用量計算 = TakeArray
    Set Dic = Nothing
    
End Function
'2019/08/14 Update by Yamashita
'各ページ番号を格納した配列と指定のページ番号を受け取り、材料・使用量を計算して配列を返す(ページ内合算の場合)
Function 使用量計算2(ByRef Page() As Long, ByVal n As Long) As Variant()
    
    Dim Dic As Object
    Dim i As Long, OffsetCol As Long, OffsetRow As Long, DataNum As Long
    Dim Coefficient As Double
    Dim buf As String
    Dim BaseRange As Range, TargetRange As Range, UnionRange As Range
    Dim obj As Object
    Dim TakeArray() As Variant
    Dim Keys, Items

    DataNum = UBound(Page()) '合計データ数(部位数)
    Set BaseRange = sh1.Cells(6, 5).Resize(30, 1)

'-----------------------------------------------------------------------------
'処理範囲を設定(Union)

    '部位数分精査
    For i = 1 To DataNum
            '一致の場合
        If Page(i) = n Then
            '読み込む工程表(i)の位置をオフセットに格納
            If i Mod 2 = 1 Then '奇数の場合
                OffsetCol = 0
                OffsetRow = 35 * Fix(i / 2)
            ElseIf i Mod 2 = 0 Then '偶数の場合
                OffsetCol = 8
                OffsetRow = 35 * (Fix(i / 2) - 1)
            End If

            '追加対象Range
            Set TargetRange = BaseRange.Offset(OffsetRow, OffsetCol)
            
            'ターゲットを処理範囲に追加
            If UnionRange Is Nothing Then
                Set UnionRange = TargetRange
            Else
                Set UnionRange = Union(UnionRange, TargetRange)
            End If
        End If
    Next i

'-----------------------------------------------------------------------------
'重複を削除した商品と使用量のリストを作成

    Set Dic = CreateObject("Scripting.Dictionary")

    For Each obj In UnionRange
    
        If obj <> "" Then
            buf = obj.Value
            
            If Not Dic.Exists(buf) Then '新たなKeyの場合、KeyとItemを登録
                Dic.Add buf, obj.Offset(0, 3).Value
            Else '既に存在するKeyの場合、Itemを加算
                Dic(buf) = Dic(buf) + obj.Offset(0, 3).Value
            End If
        
        End If
        
    Next obj
    
    Keys = Dic.Keys
    Items = Dic.Items
    
'-----------------------------------------------------------------------------
'必要数量を計算   Keys:商品、Items:必要数量

    ReDim TakeArray(1 To UBound(Keys) + 1, 1 To 2)
    On Error Resume Next
    For i = 0 To UBound(Keys)
        Coefficient = WorksheetFunction.XLookup(Keys(i), _
            sh5.ListObjects("材料DB").ListColumns("材料名").DataBodyRange, _
            sh5.ListObjects("材料DB").ListColumns("係数").DataBodyRange, "")
        Items(i) = WorksheetFunction.RoundUp(Items(i) / Coefficient, 0)
        '戻り値の配列に格納(Keys, Items)
        TakeArray(i + 1, 1) = Keys(i)
        TakeArray(i + 1, 2) = Items(i)
    Next

    使用量計算2 = TakeArray
    
    Set Dic = Nothing

End Function
Sub 保存() '2019/08/19 Update by Yamashita

    If MsgBox("見積書に保護をかけて保存します。" & vbCrLf & "※この作業は数十秒程度かかる場合があります。" _
                       & vbCrLf & vbCrLf & "編集パスワード : ssk", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Call ApplicationFalse
    Application.Calculation = xlCalculationManual
    Call SheetNames

    Dim SaveWB As Workbook
    Dim SaveFile As String, SaveFileName As String, key As String, LC As String
    Dim i As Integer, n As Integer, M As Integer
    Dim j As Long
    Dim TypeCell As Range, rng As Range, R1 As Range
    Dim shp As Shape

    Call PrintArea

    ChDrive WB.Path
    ChDir WB.Path

    n = sh6.Range("AO11").Value 'ページ数
    sh6.Cells(8, 7) = "=Y" & 1 + 48 * n '金額を値で貼り付け
    
    sh6.Copy 'シートを新しいブックにコピー
    Set SaveWB = ActiveWorkbook
    
    With SaveWB.ActiveSheet
        .Name = "見積書"
    
        key = .Range("F11") '物件名
        
        'ページ数を値で貼り付け
        For i = 0 To n
            Set R1 = Range("AL49").Offset(i * 48, 0)
            R1.Copy
            R1.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Next
        
        .Range("G8") = "=Y" & 49 + 48 * (n - 1)
        
        ActiveWindow.FreezePanes = False
        
        On Error Resume Next
        'ボタン削除
        For j = .Shapes.Count To 1 Step -1
            If .Shapes(j).Type = msoFormControl Then .Shapes(j).Delete
            If .Shapes(j).Type = msoOLEControlObject Then .Shapes(j).Delete
        Next j
    
        '印刷範囲外を削除
        .Rows(1).Delete
        .Range("AM:AU").Delete
        
        '見積書不要ページ削除
        If n <> 6 Then
            M = n * 48 + 1
            .Rows(M & ":288").EntireRow.Select
            
            For Each shp In .Shapes
            
                '図形の配置されているセル範囲をオブジェクト変数にセット
                Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
                
                '図形の配置されているセル範囲と
                '選択されているセル範囲が重なっているときに図形を削除
                If Not (Intersect(rng, Selection) Is Nothing) Then
                    shp.Delete
                End If
            
            Next
    
            .Rows(M & ":288").EntireRow.Delete
        End If
        
        Application.Calculation = xlCalculationAutomatic
        
        .Cells(1, 1).Select
        .Protect Password:="ssk"
    
    End With 'SaveWB.ActiveSheet
    
    SaveFileName = key & ".xlsx"
    SaveFile = Application.Dialogs(xlDialogSaveAs).Show(SaveFileName)
    
    If SaveFile = False Then
        MsgBox "キャンセルされました"
        SaveWB.Close
        Call ApplicationTrue
        Exit Sub
    End If
    
    Application.DisplayAlerts = True
    SaveWB.Close
    
    Call 見積分析出力
    Call ApplicationTrue
    
    MsgBox "保存が完了しました", vbOKOnly + vbInformation
        
End Sub
Sub 見積分析出力() '2019/08/19 Update by Yamashita

    On Error Resume Next
    Dim OpenWB As Workbook
    Dim Outsh As Worksheet
    Dim UserName As String, FilePath As String, FileName, myText As String
    Dim n As Long, FilePathRow As Long, FileNameRow As Long

    Call SheetNames
    Call ApplicationFalse
    
    '保存先を開く
    FilePathRow = FindRow("分析用データ 出力先パス", sh8, 1)
    FilePath = sh8.Cells(FilePathRow, 2).Value
    FileNameRow = FindRow("分析用データ ファイル名", sh8, 1)
    FileName = sh8.Cells(FileNameRow, 2).Value
    Set OpenWB = Workbooks.Open(FilePath & FileName)
    Set Outsh = OpenWB.Worksheets("output")
    
    'ログインユーザー名を取得する
    UserName = CreateObject("WScript.Network").UserName
    
    n = MaxRow(Outsh, 1) '現在の行数

    With Outsh
        .Cells(n + 1, FindColumn("ユーザー名", Outsh, 1)) = UserName
        myText = sh6.Range("A5")
        If sh6.Range("A6") <> "" Then myText = myText & " " & sh6.Range("A6")
        .Cells(n + 1, FindColumn("提出先", Outsh, 1)) = myText
        .Cells(n + 1, FindColumn("物件名", Outsh, 1)) = sh6.Range("F11")
        .Cells(n + 1, FindColumn("面積", Outsh, 1)) = WorksheetFunction.Sum(sh6.Range("O13:O18"), sh6.Range("AH13:AH18"))
        .Cells(n + 1, FindColumn("見積総額", Outsh, 1)) = sh6.Range("G8")
        .Cells(n + 1, FindColumn("見積金額", Outsh, 1)) = sh6.Range("AP7")
        .Cells(n + 1, FindColumn("標準金額", Outsh, 1)) = sh6.Range("AR7")
        .Cells(n + 1, FindColumn("特価率", Outsh, 1)) = sh6.Range("AN8")
        .Cells(n + 1, FindColumn("見積日時", Outsh, 1)) = Now
    End With
    
    OpenWB.Save
    OpenWB.Close
    
    Call ApplicationTrue
    
End Sub
Sub Printout() '2019/08/19 Update by Yamashita

    Call SheetNames
    Call ApplicationFalse

    Call PrintArea
    sh6.PrintPreview
    
    Call ApplicationTrue
    
End Sub
Sub PrintArea() '2019/08/19 Update by Yamashita

    Call SheetNames
    
    Dim i, n, EndRow As Long
    n = sh6.Range("AO11").Value '見積りページ数

    With sh6

        .ResetAllPageBreaks '改ページを全削除
        
        EndRow = n * 48 + 1 '見積り最終行
        .PageSetup.PrintArea = Range(Cells(2, 1), Cells(EndRow, 38)).Address '印刷範囲設定
                
        If EndRow > 49 Then '2ページ以上あるなら
        
            For i = 50 To EndRow Step 48
                DoEvents
                .Rows(i).PageBreak = xlPageBreakManual '改ページ設定
            Next
    
        End If
    
    End With

End Sub

Sub 特価申請書作成() '2019/08/19 Update by Yamashita

Dim RC As VbMsgBoxResult
Dim PageNum, buf, i, j, k As Long
Dim Dic As Object
Dim Keys, Items

    RC = MsgBox("特価申請書を作成します。" & vbCrLf & "※現在、特価申請書に入力されている内容は削除されます。", vbOKCancel + vbExclamation)
    
    If RC = vbCancel Then Exit Sub   'キャンセルが押されたら終了
    
    Call ApplicationFalse
    Call SheetNames
    Call sh10クリア
    
    PageNum = sh6.Range("AO11")
    
    sh10.Cells(12, 7) = sh6.Cells(11, 6) '物件名
    sh10.Cells(40, 7) = sh6.Cells(11, 6) '物件名
'        sh10.Cells(13, 7) = sh6.Cells(11, 6) '物件内容
    sh10.Cells(14, 7) = sh6.Cells(11, 27) '元請
'        sh10.Cells(15, 7) = sh6.Cells(11, 27) '取引先
'        sh10.Cells(16, 7) = sh6.Cells(11, 27) '納入先
    sh10.Cells(23, 12) = sh6.Cells(7, 42) '合計金額(特別単価)
    sh10.Cells(24, 12) = sh6.Cells(7, 45) '合計金額(標準単価)
    sh10.Cells(23, 30) = sh6.Cells(48, 25).Offset(48 * (PageNum - 1), 0) '合計金額(標準単価)
    

    Set Dic = CreateObject("Scripting.Dictionary")

    For i = 0 To PageNum - 1
    
        For j = 1 To 24
            If sh6.Cells(22, 42).Offset(j + 48 * i, 0).Value <> "" Then '材料名が空白でない場合
                buf = sh6.Cells(22, 1).Offset(j + 48 * i, 0).Value
                
                If Not Dic.Exists(buf) Then Dic.Add buf, buf 'Dictionaryに未登録の場合、連想配列に登録
            End If

        Next j
        
    Next i
        
    Keys = Dic.Keys
    
    '特価申請書へ書込み
    For k = 0 To Dic.Count - 1
    
        sh10.Cells(43, 1).Offset(k, 0) = Keys(k)
        sh10.Cells(43, 11).Offset(k, 0) = "=SUMIF(見積書!A:A, A" & 43 + k & ",見積書!M:M)"
        sh10.Cells(43, 15).Offset(k, 0) = "=VLOOKUP(A" & 43 + k & ",材料DB,COLUMN(材料DB[荷姿単位]),FALSE)"
        sh10.Cells(43, 19).Offset(k, 0) = "=VLOOKUP(A" & 43 + k & ",見積書!A:AT,21,FALSE)"
        sh10.Cells(43, 28).Offset(k, 0) = "=VLOOKUP(A" & 43 + k & ",見積書!A:AT,45,FALSE)"

    Next k
    
    Set Dic = Nothing


    '値で貼り付け
    sh10.Range("K43:V62").Copy
    sh10.Range("K43:V62").PasteSpecial Paste:=xlPasteValues
        
    sh10.Range("AB43:AE62").Copy
    sh10.Range("AB43:AE62").PasteSpecial Paste:=xlPasteValues
        
    Application.CutCopyMode = False
        

        
    Call ApplicationTrue
    
    sh10.Select
    MsgBox "作成が完了しました｡", vbInformation

    

End Sub



