VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetTranslator"
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

'Private Const TARGET_CLASS As String = "trans-sentence ng-star-inserted inmt-on"
Private Const TARGET_ELEMENT_ID As String = "target-dummydiv"
Private Const OUTPUT_ATTRIBUTE As String = "textContent"
'Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/ja/"
Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/"
'Private Const URL_TRANSLATOR As String = "https://miraitranslate.com/trial/#ko/ja/"
Private Const WINDOW_ZOOM As Long = 60
Private Const PAGE_ZOOM As Long = 45
Private Const WAIT_MILLISECONDS As Long = 500
Private Const SECOND_TO_RESTART As Long = 300
Private Const MAX_RETRY_COUNT As Long = 60

Private srcSheet_ As Worksheet
Private dstSheet_ As Worksheet
Private lastRow_ As Long
Private currentRow_ As Long
Private startTime_ As Double
Private stopTime_ As Double
Private browser_ As Selenium.WebDriver
Private targets_ As Collection
Private originalTextColumn_ As Long
Private translatedTextColumn_ As Long
Private outputCombobox_ As MSForms.ComboBox
Private originalTextTitle_ As String
Private translatedTextTitle_ As String
Private outputLanguage_ As String

Public Sub Initialize(ByRef dstSheet As Worksheet)

    Set targets_ = New Collection
    Set srcSheet_ = ThisWorkbook.Worksheets(1)
    Set dstSheet_ = dstSheet
    lastRow_ = GetLastRow(srcSheet_)
    currentRow_ = 2
    
    Set outputCombobox_ = ThisWorkbook.Worksheets(1).OutputComboBox
    outputLanguage_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_OUTPUT_LANGUAGE)
    originalTextTitle_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_ORIGINAL_TEXT)
    translatedTextTitle_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_TRANSLATED_TEXT)
    If ThisWorkbook.Worksheets(1).LeftOriginalTextRadio.Value Then
        originalTextColumn_ = 1
        translatedTextColumn_ = 2
    Else
        originalTextColumn_ = 2
        translatedTextColumn_ = 1
    End If
    Call FormatTargetSheet
    Call CopyOriginalText

End Sub

Private Sub CopyOriginalText()

    Dim virtualCell As Variant
    virtualCell = srcSheet_.Range(srcSheet_.Cells(TranslatorSheet.ROW_START, TranslatorSheet.COL_START), srcSheet_.Cells(lastRow_, TranslatorSheet.COL_START))
    dstSheet_.Range(dstSheet_.Cells(TranslatorSheet.ROW_START, originalTextColumn_), dstSheet_.Cells(lastRow_, originalTextColumn_)) = virtualCell

End Sub

Private Sub FormatTargetSheet()

    With dstSheet_.Columns(originalTextColumn_)
        .ColumnWidth = 90
        .Font.Name = "UD デジタル 教科書体 N-R"
        .NumberFormatLocal = "@"
        .Font.Size = 14
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With

    With dstSheet_.Columns(translatedTextColumn_)
        .ColumnWidth = 90
        .Font.Name = "UD デジタル 教科書体 N-R"
        .NumberFormatLocal = "@"
        .Font.Size = 14
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    
    With dstSheet_.Cells(TranslatorSheet.ROW_TITLE, originalTextColumn_)
        .Value = originalTextTitle_
        .HorizontalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    With dstSheet_.Cells(TranslatorSheet.ROW_TITLE, translatedTextColumn_)
        .Value = translatedTextTitle_
        .HorizontalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    ActiveWindow.Zoom = WINDOW_ZOOM
    With ActiveSheet.PageSetup
        .Zoom = PAGE_ZOOM
        .TopMargin = Application.CentimetersToPoints(1.4)
        .BottomMargin = Application.CentimetersToPoints(1.4)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .HeaderMargin = Application.CentimetersToPoints(0.6)
        .FooterMargin = Application.CentimetersToPoints(0.6)
        .CenterHeader = "&A"
        .CenterFooter = "&P/&N"
        .CenterHorizontally = True
    End With

End Sub

Public Sub TransLateToSheet()

    Do
        Set browser_ = New Selenium.WebDriver
        browser_.AddArgument "--headless --disable-gpu"
        browser_.Start "chrome"
        startTime_ = Timer
        Do While (currentRow_ <= lastRow_)
            Dim targetText As String
            targetText = Trim(WorksheetFunction.EncodeURL(dstSheet_.Cells(currentRow_, originalTextColumn_).Value))
            If Len(Trim(targetText)) <> 0 Then
                Dim myTarget As Target: Set myTarget = New Target
                Call myTarget.Initialize(dstSheet_.Cells(currentRow_, originalTextColumn_).Value, CStr(TranslateDeepL(targetText)))
                targets_.Add myTarget
                dstSheet_.Cells(currentRow_, translatedTextColumn_).Value = myTarget.translatedText_
            End If
            currentRow_ = currentRow_ + 1
            stopTime_ = Timer
            If BeRestarted Then
                GoTo Continue
            End If
        Loop
Continue:
        browser_.Quit
        Debug.Print ("経過秒：" & stopTime_ - startTime_)
        Debug.Print ("完了行：" & currentRow_ - TranslatorSheet.ROW_START)
    Loop While (currentRow_ <= lastRow_)

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
    Dim connected As Boolean
    connected = browser_.Get(URL_TRANSLATOR & outputLanguage_ & "/" & parameter)
    If Not connected Then
        TranslateDeepL = "タイムアウトしました。"
        Exit Function
    End If
    Dim retryCount As Long
    retryCount = 0
    Dim s As String
    Do
        DoEvents
        s = CStr(browser_.FindElementById(TARGET_ELEMENT_ID).Attribute(OUTPUT_ATTRIBUTE))
'        s = browser_.FindElementByClass(TARGET_CLASS)
        s = Replace(s, vbCrLf, "")
        browser_.Wait WAIT_MILLISECONDS
        retryCount = retryCount + 1
    Loop Until (Len(s) <> 0 Or retryCount >= MAX_RETRY_COUNT)
    TranslateDeepL = CStr(browser_.FindElementById(TARGET_ELEMENT_ID).Attribute(OUTPUT_ATTRIBUTE))
'    TranslateDeepL = browser_.FindElementByClass(TARGET_CLASS)

End Function

Public Sub DrawLine()

    Dim row As Long
    For row = TranslatorSheet.ROW_START To TranslatorSheet.ROW_START + lastRow_ - 1
        If dstSheet_.Rows(row).RowHeight <= TranslatorSheet.MAX_ROW_HEIGHT - TranslatorSheet.MARGIN_HEIGHT Then
            dstSheet_.Rows(row).RowHeight = dstSheet_.Rows(row).RowHeight + TranslatorSheet.MARGIN_HEIGHT
        End If
    Next
    If (originalTextColumn_ < translatedTextColumn_) Then
        dstSheet_.Range(dstSheet_.Cells(TranslatorSheet.ROW_TITLE, originalTextColumn_), dstSheet_.Cells(lastRow_, translatedTextColumn_)).Borders.LineStyle = xlContinuous
        Exit Sub
    End If
    dstSheet_.Range(dstSheet_.Cells(TranslatorSheet.ROW_TITLE, translatedTextColumn_), dstSheet_.Cells(lastRow_, originalTextColumn_)).Borders.LineStyle = xlContinuous

End Sub

