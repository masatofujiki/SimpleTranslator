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

Private Const ROW_TITLE_OF_INVENTION As Long = 2
Private Const COL_EXTRACTION As Long = 1
Private Const TITLE_H1_TAG As String = "h1"

Private Sub DeleteButton_Click()

    Call DeleteColumns

End Sub

Public Sub DeleteColumns()

    Dim lastRow As Long: lastRow = Cells(Rows.Count, COL_SRC_ORIGINAL).End(xlUp).row
    If lastRow <= 1 Then
        Exit Sub
    End If
    Range(Cells(ROW_START, COL_SRC_ORIGINAL), Cells(lastRow, COL_SRC_ORIGINAL)).ClearContents
    Range(Cells(ROW_START, COL_SRC_ORIGINAL), Cells(lastRow, COL_SRC_ORIGINAL)).Borders.LineStyle = xlNone
    Range(Cells(ROW_START, COL_SRC_ORIGINAL), Cells(lastRow, COL_SRC_ORIGINAL)).Borders(xlEdgeTop).LineStyle = xlNone
    Range(Cells(ROW_START, COL_SRC_ORIGINAL), Cells(lastRow, COL_SRC_ORIGINAL)).Interior.ColorIndex = xlNone
    Range(Cells(ROW_START, COL_SRC_ORIGINAL), Cells(lastRow, COL_SRC_ORIGINAL)).EntireRow.AutoFit

End Sub

Private Sub ExtractButton_Click()

    If Len(PublicationNumberTextBox.Text) = 0 Then
        MsgBox "公開番号が空です。"
        Exit Sub
    End If

    Dim extractor As TextExtracctor
    Set extractor = New TextExtracctor
    Call extractor.Initialize(PublicationNumberTextBox.Text)
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
    Call shrinker.Initialize(COL_SRC_ORIGINAL)
    Call shrinker.FormatText

    Application.ScreenUpdating = True

End Sub

Private Sub MargeButton_Click()

    If (Selection.Count <> 2) Then
        MsgBox "２つの行が選択されていません。"
        Exit Sub
    End If

    If (Selection.Areas.Count = 1) Then
        If (Selection.Columns.Count <> 1) Then
            MsgBox "同じ行の列が選択されています。"
            Exit Sub
        End If
        If (Cells(Selection(1).row, 1).Value = "" Or Cells(Selection(2).row, 1) = "") Then
            Exit Sub
        End If
        Cells(Selection(1).row, 1) = Trim(Cells(Selection(1).row, 1) & " " & Cells(Selection(2).row, 1))
        Cells(Selection(2).row, 1) = ""
        Exit Sub
    End If

    If (Selection.Areas.Count = 2) Then
        If (Selection.Areas(1).Column <> 1 Or Selection.Areas(2).Column <> 1) Then
            MsgBox "１列目以外の列が選択されています。"
            Exit Sub
        End If
        Cells(Selection.Areas(1).row, 1) = Trim(Cells(Selection.Areas(1).row, 1) & " " & Cells(Selection.Areas(2).row, 1))
        Cells(Selection.Areas(2).row, 1) = ""
    End If

End Sub

Private Sub PatentFormatButton_Click()

    Application.ScreenUpdating = False

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(COL_SRC_ORIGINAL)
    Call shrinker.FormatPatentText

    Application.ScreenUpdating = True

End Sub

'Private Sub ThesisSelectionButton_Click()
'
'    Dim myFileDialog As FileDialog
'    Set myFileDialog = Application.FileDialog(msoFileDialogFilePicker)
'    myFileDialog.AllowMultiSelect = False
'    myFileDialog.Title = "翻訳する論文PDFを選択してください。"
'    myFileDialog.Filters.Clear
'    myFileDialog.Filters.Add "すべてのPDFファイル", "*.pdf"
'
'    If myFileDialog.Show <> -1 Then
'        myFileDialog.Filters.Clear
'        Exit Sub
'    End If
'
'    Call DeleteColumns
'
'    Dim myFilePath As String
'    myFilePath = myFileDialog.SelectedItems(1)
'    myFileDialog.Filters.Clear
'
'    Dim translator As WordTranslator
'    Set translator = New WordTranslator
'    Call translator.Initialize(myFilePath)
'    Call translator.CreateWordParagraphs
''    Call translator.CreateTranslatedPDF
''    Call translator.CreatePDF
'
''    Dim wordApp As Word.Application
''    Set wordApp = New Word.Application
''    wordApp.Visible = False
''    wordApp.Visible = True
''
''    Dim wordDoc As Word.Document
''    Set wordDoc = wordApp.Documents.Open(myFilePath)
''
''    Debug.Print wordDoc.Paragraphs.Count
''
''    Dim para As Paragraph
''    Dim row As Long: row = 2
''    For Each para In wordDoc.Paragraphs
''
''        Cells(row, 1).Value = para.Range.Text
''        Cells(row, 2).Value = para.Range.Style
''        row = row + 1
''    Next
'
''    Dim mySelection As Word.Selection
''    Set mySelection = wordApp.Selection
''    With mySelection.Find
''        .Text = targetText
''        .Replacement.Text = ""
''        .Wrap = wdFindContinue
''        .MatchByte = False
''        .MatchCase = False
''        Do While .Execute
''            mySelection.Range.Shading.BackgroundPatternColor = targetBackColor
''            mySelection.Font.Color = targetForeColor
''        Loop
''    End With
'
''    Dim ws As Worksheet
''    Set ws = ThisWorkbook.Worksheets(1)
''    Dim i As Long
''    For i = 1 To 11
''        Dim row As Long
''        For row = 0 To ws.OLEObjects("ListBox" & i).Object.ListCount - 1
''            Call ReplaceColor(mySelection, ws.OLEObjects("ListBox" & i).Object.List(row), ws.OLEObjects("ListBox" & i).Object.BackColor, ws.OLEObjects("ListBox" & i).Object.ForeColor)
''        Next
''    Next
''
''    Dim fso As FileSystemObject
''    Set fso = New FileSystemObject
''    Dim dirPath As String: dirPath = fso.GetParentFolderName(myFilePath)
''    Dim baseFileName As String: baseFileName = fso.GetBaseName(myFilePath)
''    Dim fullPath As String: fullPath = fso.BuildPath(dirPath, baseFileName & "_color.pdf")
''    wordDoc.SaveAs2 fileName:=fullPath, FileFormat:=wdFormatPDF
''
''    wordDoc.Close SaveChanges:=False
'''    wordApp.Visible = False
''    wordApp.Quit
''    Set wordApp = Nothing
''
''    MsgBox "色塗りが終了しました。"
'
'End Sub

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
    Call translator.Initialize(myFilePath)
    Call translator.CreateWordParagraphs

End Sub
'
Private Sub PublicationNumberTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Len(PublicationNumberTextBox.Text) = 0 Then
            MsgBox "公開番号が空です。"
            Exit Sub
        End If
        Dim extractor As TextExtracctor
        Set extractor = New TextExtracctor
        Call extractor.Initialize(PublicationNumberTextBox.Text)
        If Not extractor.IsReady Then
            MsgBox "願番が存在しません。"
            Exit Sub
        End If
        Call DeleteColumns
        Call extractor.ExtractPatentText
    End If

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

    Application.ScreenUpdating = True

End Sub

Private Sub TranslateToSheetButton_Click()

    Dim lastRow As Long
    lastRow = GetLastRow(ThisWorkbook.Worksheets(1))
    If (lastRow <= 1) Then
        MsgBox "原文が入力されていません"
        Exit Sub
    End If

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(COL_SRC_ORIGINAL)
    Call shrinker.FormatText

    Dim translator As SheetTranslator
    Set translator = New SheetTranslator
    Call translator.Initialize(Worksheets.Add(After:=Worksheets(Worksheets.Count)))
    Call translator.TransLateToSheet
    Call translator.DrawLine

End Sub

