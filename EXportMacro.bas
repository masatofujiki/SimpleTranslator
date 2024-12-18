Attribute VB_Name = "EXportMacro"
Option Explicit

Sub ExportAll()
    Dim module As VBComponent      '// モジュール
    Dim moduleList As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath                                       '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook                                  '// 処理対象ブックオブジェクト
    
    '// ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// ブックが開かれている場合は表示しているブックを対象とする
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.path
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        '// クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        '// 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// その他
        Else
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        '// エクスポート実施
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub
