VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'MIT License
'
'Copyright (c) [2021] [Masato Fujiki]
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Option Explicit

Private Const URL_GOOGLE_PATENTS As String = "https://patents.google.com/"
Private Const URL_GOOGLE_PATENTS_PATENT As String = "https://patents.google.com/patent/"

Private Sub DeleteButton_Click()

    Call DeleteColumns

End Sub

Public Sub DeleteColumns()

    Dim lastRow As Long: lastRow = Cells(Rows.Count, TranslatorSheet.COL_SRC_ORIGINAL).End(xlUp).row
    If lastRow <= 1 Then
        Exit Sub
    End If
    Range(Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_SRC_ORIGINAL), Cells(lastRow, TranslatorSheet.COL_SRC_ORIGINAL)).ClearContents
    Range(Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_SRC_ORIGINAL), Cells(lastRow, TranslatorSheet.COL_SRC_ORIGINAL)).Borders.LineStyle = xlNone
    Range(Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_SRC_ORIGINAL), Cells(lastRow, TranslatorSheet.COL_SRC_ORIGINAL)).Borders(xlEdgeTop).LineStyle = xlNone
    Range(Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_SRC_ORIGINAL), Cells(lastRow, TranslatorSheet.COL_SRC_ORIGINAL)).Interior.ColorIndex = xlNone
    Range(Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_SRC_ORIGINAL), Cells(lastRow, TranslatorSheet.COL_SRC_ORIGINAL)).EntireRow.AutoFit
    Range("A2").Select

End Sub

Private Sub ExtractButton_Click()

    If Len(PublicationNumberTextBox.text) = 0 Then
        MsgBox "公開番号が空です。"
        Exit Sub
    End If

    Dim extractor As TextExtracctor
    Set extractor = New TextExtracctor
    Call extractor.Initialize(PublicationNumberTextBox.text)
    If Not extractor.IsReady Then
        MsgBox "願番が存在しません。"
        Exit Sub
    End If
    Call DeleteColumns
    Call extractor.ExtractPatentText

End Sub

Private Sub FormatButton_Click()

    Application.ScreenUpdating = False

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(TranslatorSheet.COL_SRC_ORIGINAL)
    Call shrinker.FormatText

    Application.ScreenUpdating = True

End Sub

Private Sub MargeButton_Click()

'    If (Selection.Count <> 2) Then
'        MsgBox "２つの行が選択されていません。"
'        Exit Sub
'    End If
'
'    If (Selection.Areas.Count = 1) Then
'        If (Selection.Columns.Count <> 1) Then
'            MsgBox "同じ行の列が選択されています。"
'            Exit Sub
'        End If
'        If (Cells(Selection(1).row, 1).Value = "" Or Cells(Selection(2).row, 1) = "") Then
'            Exit Sub
'        End If
'        Cells(Selection(1).row, 1) = Trim(Cells(Selection(1).row, 1) & " " & Cells(Selection(2).row, 1))
'        Cells(Selection(2).row, 1) = ""
'        Exit Sub
'    End If
'
'    If (Selection.Areas.Count = 2) Then
'        If (Selection.Areas(1).Column <> 1 Or Selection.Areas(2).Column <> 1) Then
'            MsgBox "１列目以外の列が選択されています。"
'            Exit Sub
'        End If
'        Cells(Selection.Areas(1).row, 1) = Trim(Cells(Selection.Areas(1).row, 1) & " " & Cells(Selection.Areas(2).row, 1))
'        Cells(Selection.Areas(2).row, 1) = ""
'    End If

    If (Selection.Count < 2) Then
        MsgBox "2つ以上の行が選択されていません。"
        Exit Sub
    End If

    'とびとびの領域がなければ「Selection.Areas.Count」は1になる。とびとびの領域があれば「Selection.Areas.Count」は選択した領域の数になる
    If (Selection.Areas.Count <> 1) Then
        MsgBox "複数の間の空いた領域は選択できません。"
        Exit Sub
    End If

    If (Selection.Columns.Count <> 1) Then
        MsgBox "同じ行の列が選択されています。"
        Exit Sub
    End If

    Dim result As String
    result = Trim(Cells(Selection.Areas(1).row, 1))
    Dim i As Long
    For i = 1 To Selection.Rows.Count - 1
        If Len(Cells(Selection.Areas(1).row + i, 1)) <> 0 Then
            result = result & " " & Trim(Cells(Selection.Areas(1).row + i, 1))
            Cells(Selection.Areas(1).row + i, 1) = ""
        End If
    Next
    Cells(Selection.Areas(1).row, 1) = result
    
    Application.ScreenUpdating = False
    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(TranslatorSheet.COL_SRC_ORIGINAL)
    Call shrinker.FormatText
    Cells(Selection.Areas(1).row, 1).Select
    Application.ScreenUpdating = True

End Sub

Private Sub OutputComboBox_Change()

    FormatButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_SHAPING)
    PatentFormatButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_PATENT_SHAPING)
    TranslateToHtmlButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_TRANSLATION_HTML)
    TranslateToSheetButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_TRANSLATION_SHEET)
    MargeButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_MERGE)
    DeleteButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_DELETE)
    PDFSelectionButton.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_PDF_SELECTION)
    LeftTranslatedTextRadio.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_TRANSLATED_TEXT) + "→" + OutputComboBox.List(OutputComboBox.ListIndex, COL_ORIGINAL_TEXT)
    LeftOriginalTextRadio.Caption = OutputComboBox.List(OutputComboBox.ListIndex, COL_ORIGINAL_TEXT) + "→" + OutputComboBox.List(OutputComboBox.ListIndex, COL_TRANSLATED_TEXT)

End Sub

Private Sub PatentFormatButton_Click()

    Application.ScreenUpdating = False

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(TranslatorSheet.COL_SRC_ORIGINAL)
    Call shrinker.FormatPatentText

    Application.ScreenUpdating = True

End Sub

Private Sub PDFSelectionButton_Click()

    Dim myFileDialog As FileDialog
    Set myFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    myFileDialog.AllowMultiSelect = False
    myFileDialog.Title = "翻訳するPDFを選択してください。"
    myFileDialog.Filters.Clear
    myFileDialog.Filters.Add "すべてのPDFファイル", "*.pdf"

    If myFileDialog.Show <> -1 Then
        myFileDialog.Filters.Clear
        Exit Sub
    End If

    Call DeleteColumns

    Dim myFilePath As String
    myFilePath = myFileDialog.SelectedItems(1)
    myFileDialog.Filters.Clear

    Dim translator As WordTranslator
    Set translator = New WordTranslator
    Call translator.Initialize(myFilePath, OutputComboBox.List(OutputComboBox.ListIndex, COL_OUTPUT_LANGUAGE))
    Call translator.CreateWordParagraphs

End Sub
'
Private Sub PublicationNumberTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'    If KeyCode = vbKeyReturn Then
'        If Len(PublicationNumberTextBox.text) = 0 Then
'            MsgBox "公開番号が空です。"
'            Exit Sub
'        End If
'        Dim extractor As TextExtracctor
'        Set extractor = New TextExtracctor
'        Call extractor.Initialize(PublicationNumberTextBox.text)
'        If Not extractor.IsReady Then
'            MsgBox "願番が存在しません。"
'            Exit Sub
'        End If
'        Call DeleteColumns
'        Call extractor.ExtractPatentText
'    End If

End Sub

Private Sub TranslateToHtmlButton_Click()

    Dim fileFullPath As String
    fileFullPath = GetFilePath
    If Len(fileFullPath) = 0 Then
        Exit Sub
    End If
    
    If Len(Dir(fileFullPath)) > 0 Then
        If MsgBox("同名のファイルが存在します。上書きしますか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileName As String
    fileName = fso.GetBaseName(fileFullPath)
    If FileNameHasNGCharacters(fileName) Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = GetLastRow(ThisWorkbook.Worksheets(1))
    If (lastRow <= 1) Then
        MsgBox "原文が入力されていません"
        Exit Sub
    End If
    
    Dim translator As HtmlTranslator
    Set translator = New HtmlTranslator
    Call translator.Initialize(fileFullPath)
    If Not translator.HasFileName Then
        MsgBox "ファイル名が入力されていません。"
        Exit Sub
    End If
        
    Call translator.TransLateToHtml
    Call translator.CreateHtml
    If Len(PublicationNumberTextBox.text) > 0 Then
        PublicationNumberTextBox.text = ""
    End If

    Application.ScreenUpdating = True

End Sub

Private Sub TranslateToSheetButton_Click()

    Dim lastRow As Long
    lastRow = GetLastRow(ThisWorkbook.Worksheets(1))
    If (lastRow <= 1) Then
        MsgBox "原文が入力されていません"
        Exit Sub
    End If

    Dim translator As SheetTranslator
    Set translator = New SheetTranslator
    Call translator.Initialize(Worksheets.Add(After:=Worksheets(Worksheets.Count)))
    Call translator.TransLateToSheet
    Call translator.DrawLine

End Sub

Public Sub InitializeOutputComboBox()

    Dim settingSheet As Worksheet
    Set settingSheet = ThisWorkbook.Worksheets("deepl_setting")
    Dim outputSetting As Variant
    Dim lastRow As Long
    lastRow = GetLastRow(settingSheet)
    outputSetting = settingSheet.Range(settingSheet.Cells(Setting.ROW_START, Setting.COL_START), settingSheet.Cells(lastRow, Setting.COL_END))
    With OutputComboBox
        .ColumnCount = 1
        .List = outputSetting
    End With

End Sub

