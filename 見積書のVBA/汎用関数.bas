Attribute VB_Name = "汎用関数"
Public WB As Workbook
Public sh1 As Worksheet, sh2 As Worksheet, sh3 As Worksheet, sh5 As Worksheet
Public sh6 As Worksheet, sh7 As Worksheet, sh8 As Worksheet, sh10 As Worksheet
Option Explicit
Sub ApplicationFalse()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'    Application.EnableEvents = False
End Sub
Sub ApplicationTrue()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True
End Sub
Function MaxRow(ByVal WSH As Worksheet, ByVal C As Long) As Long '最終行取得関数
    MaxRow = WSH.Cells(Rows.Count, C).End(xlUp).Row
End Function
Function MaxColumn(ByVal WSH As Worksheet, ByVal r As Long) As Long '最終列取得関数
    MaxColumn = WSH.Cells(r, Columns.Count).End(xlToLeft).Column
End Function
Function FindColumn(ByVal Name As String, ByVal WSH As Worksheet, ByVal r As Long) As Long '列数検索関数
    FindColumn = WorksheetFunction.Match(Name, WSH.Rows(r), 0)
End Function
Function FindRow(ByVal Name As String, ByVal WSH As Worksheet, ByVal C As Long) As Long '行数検索関数
    FindRow = WorksheetFunction.Match(Name, WSH.Columns(C), 0)
End Function
Function CountString(ByVal Str As String, ByVal SrchRange As Range) As Long '文字列カウント関数
    CountString = WorksheetFunction.CountIf(SrchRange, Str)
End Function
Function MaxValue(ByVal ValueRange As Range)
    MaxValue = WorksheetFunction.Max(ValueRange)
End Function
Function SheetLock(ByVal WSH As Worksheet)
    WSH.Protect Password:=20140101, UserInterfaceOnly:=True
End Function
Function SheetUnlock(ByVal WSH As Worksheet)
    WSH.Unprotect Password:=20140101
End Function
Sub SheetNames()
    Set WB = ThisWorkbook
    Set sh1 = WB.Worksheets("工程表")
    Set sh2 = WB.Worksheets("工程DB")
    Set sh3 = WB.Worksheets("基本設定")
'    Set sh4 = WB.Worksheets("計算用")
    Set sh5 = WB.Worksheets("材料DB")
    Set sh6 = WB.Worksheets("見積書")
    Set sh7 = WB.Worksheets("設計価格")
    Set sh8 = WB.Worksheets("リスト")
'    Set sh9 = WB.Worksheets("取引先マスタ")
    Set sh10 = WB.Worksheets("特価申請書")
End Sub
Sub Select工程表()
    Call ApplicationFalse
    Call SheetNames
    sh1.Select
    sh1.Range("A2").Select
    Call ApplicationTrue
End Sub
Sub Select入力()
    Call ApplicationFalse
    Call SheetNames
    sh3.Select
    sh3.Range("A1").Select
    Call ApplicationTrue
End Sub
