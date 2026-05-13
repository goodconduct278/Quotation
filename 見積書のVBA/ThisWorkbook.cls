VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'ブックを開く際に実行
Private Sub Workbook_Open()

    Call ApplicationFalse
    Call SheetNames
    
    '各シートの保護を開始(VBAによる編集は許可)
    Call SheetLock(sh1)
    Call SheetLock(sh3)
    Call SheetLock(sh6)
    sh3.Select

    Call ApplicationTrue
    
    Application.Wait Now() + TimeValue("00:00:01")
    '自動チェックが有効な場合、アップデートをチェック
    If sh3.OLEObjects("CheckBox3").Object.Value = True Then Call UpdateMaster
    
End Sub

