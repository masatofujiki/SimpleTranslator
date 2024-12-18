VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const TARGET_ELEMENT_ID As String = "target-dummydiv"
Private Const OUTPUT_ATTRIBUTE As String = "textContent"
'Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/ja/"
Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/"
Private Const WAIT_MILLISECONDS As Long = 500
Private Const SECOND_TO_RESTART As Long = 300
Private Const MAX_RETRY_COUNT As Long = 100

Private startTime_ As Double
Private stopTime_ As Double
Private driver_ As WebDriver
Private originals_ As Collection
Private fileName_ As String
Private wordApp_ As Object
Private wordDoc_ As Object
Private currentParagraph_ As Long
Private lastParagraph_ As Long
Private outputLanguage_ As String

Public Sub Initialize(ByVal fileName As String, ByVal outputLanguage As String)

    Set wordApp_ = CreateObject("Word.Application")
    Set originals_ = New Collection
    fileName_ = fileName
    wordApp_.Visible = True
    currentParagraph_ = 1
    outputLanguage_ = outputLanguage

End Sub

Public Sub CreateWordParagraphs()

    Set wordDoc_ = wordApp_.Documents.Open(fileName_)
    lastParagraph_ = wordDoc_.Paragraphs.Count
    Dim para As Object
    For Each para In wordDoc_.Paragraphs
        Dim original As WordOriginal
        Set original = New WordOriginal
        Dim fn As String
        Dim originalText As String
        originalText = CleanChar(para.Range.text)
'        fn = "UD デジタル 教科書体 N-R"
        fn = "BIZ UDP明朝 Medium"
'        Debug.Print originalText
        Call original.Initialize(originalText, "", para.LeftIndent, para.RightIndent, para.Alignment, fn, para.Range.Font.Size, para.Range.Style)
        Call originals_.Add(original)
    Next

    wordDoc_.Close SaveChanges:=False
    Set wordDoc_ = Nothing
    wordApp_.Quit
    Set wordApp_ = Nothing

    Dim created As WordOriginal
    Dim row As Long
    row = 2
    For Each created In originals_
        Cells(row, 1).value = created.originalText_
'        Cells(row, 2).Value = created.originalStyle_
        row = row + 1
    Next

End Sub

Public Sub CreateTranslatedPDF()

    Call ProgressBarForm.InitializeProgress(currentParagraph_, lastParagraph_)
    Call ProgressBarForm.Show(vbModeless)
    DoEvents
    Do
        Set driver_ = New WebDriver
        'SafeOpenHeadless driver_, Chrome
        'SafeOpen driver_, Chrome
        SafeOpen driver_, Chrome, CapabilityArgs:="--headless --disable-gpu"
        startTime_ = Timer
        Do While (currentParagraph_ <= lastParagraph_)
            Dim original As WordOriginal
            Set original = originals_.Item(currentParagraph_)
            Dim targetText As String
            targetText = TrimLF(Trim(original.originalText_))
            If Len(targetText) <> 0 Then
                targetText = EncodeBasedOnExcelVersion(targetText)
                original.translatedText_ = TranslateDeepL(targetText)
                Call ProgressBarForm.UpdateProgress(currentParagraph_, lastParagraph_)
            End If
            currentParagraph_ = currentParagraph_ + 1
            stopTime_ = Timer
            If BeRestarted Then
                GoTo CONTINUE
            End If
            If ProgressBarForm.ProgressIsCancel Then
                MsgBox "翻訳処理が中断されました"
                GoTo CONTINUE
            End If
        Loop
CONTINUE:
        driver_.Shutdown
        Debug.Print ("経過秒：" & stopTime_ - startTime_)
        Debug.Print ("完了行：" & currentParagraph_)
    Loop While (currentParagraph_ <= lastParagraph_ And ProgressBarForm.ProgressIsCancel = False)
    Unload ProgressBarForm

End Sub

Private Function BeRestarted() As Boolean

    BeRestarted = (stopTime_ - startTime_ > SECOND_TO_RESTART)

End Function

Private Function TranslateDeepL(ByVal targetText As String) As String

    Dim parameter As String
    parameter = targetText
    parameter = Replace(parameter, "%2F", "%5C%2F")
'    parameter = Replace(parameter, "%26", "%5C%26")
    parameter = Replace(parameter, "%81F", "%5C%81F")
    parameter = Replace(parameter, "%81G", "%5C%81G")
    parameter = Replace(parameter, "%81H", "%5C%81H")
    parameter = Replace(parameter, "%7C", "%5C%7C")
    Call driver_.Navigate(URL_TRANSLATOR & outputLanguage_ & "/" & parameter)
    Dim retryCount As Long
    retryCount = 0
    Dim s As String
    Do
        s = CStr(Trim(driver_.ExecuteScript("return document.getElementById('target-dummydiv').textContent")))
        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, " ", "")
        s = Replace(s, "　", "")
        Call Sleep(WAIT_MILLISECONDS)
        retryCount = retryCount + 1
    Loop Until (Len(s) <> 0 Or retryCount >= MAX_RETRY_COUNT)
    TranslateDeepL = CStr(driver_.ExecuteScript("return document.getElementById('target-dummydiv').textContent"))
    TranslateDeepL = CStr(driver_.ExecuteScript("return document.getElementById('target-dummydiv').textContent"))
    

End Function

Public Sub CreatePDF()

    Set wordDoc_ = wordApp_.Documents.Open(fileName_)
    Dim index As Long
    For index = 1 To wordDoc_.Paragraphs.Count
        Dim original As WordOriginal
        Set original = originals_.Item(index)
        Dim translated As WordOriginal
        Set translated = originals_.Item(index)
        If wordDoc_.Paragraphs(index).Range.Style = "TableGrid" Then
            GoTo CONTINUE
        End If
        If InStr(wordDoc_.Paragraphs(index).Range.text, "/") = 1 Then
            Debug.Print translated.translatedText_
            GoTo CONTINUE
        End If
        If Len(translated.translatedText_) <> 0 And (translated.translatedText_ <> "/") Then
            wordDoc_.Paragraphs(index).Range.text = original.translatedText_
            wordDoc_.Paragraphs(index).LeftIndent = original.originalLeftIndent_
            wordDoc_.Paragraphs(index).RightIndent = original.originalRightIndent_
            wordDoc_.Paragraphs(index).Alignment = original.originalAlignment_
            wordDoc_.Paragraphs(index).Range.Font.Name = original.translatedFontName_
            wordDoc_.Paragraphs(index).Range.Font.Size = original.originalFontSize_
        End If
CONTINUE:
    Next

End Sub

