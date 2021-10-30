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

Private Sub DeleteButton_Click()

    Call DeleteColumns

End Sub

Private Sub DeleteColumns()

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


Private Sub FormatButton_Click()

    Application.ScreenUpdating = False

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(COL_SRC_ORIGINAL)
    Call shrinker.FormatText

    Application.ScreenUpdating = True

End Sub

Private Sub PatentFormatButton_Click()

    Application.ScreenUpdating = False

    Dim shrinker As TextShrinker
    Set shrinker = New TextShrinker
    Call shrinker.Initialize(COL_SRC_ORIGINAL)
    Call shrinker.FormatPatentText

    Application.ScreenUpdating = True

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
    Call translator.Initialize(fileName)
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

