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

Private Const TARGET_ELEMENT_ID As String = "target-dummydiv"
Private Const OUTPUT_ATTRIBUTE As String = "textContent"
Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/ja/"
Private Const WINDOW_ZOOM As Long = 60
Private Const PAGE_ZOOM As Long = 45
Private Const WAIT_TIME As Long = 500

Private srcSheet_ As Worksheet
Private dstSheet_ As Worksheet
Private lastRow_ As Long
Private currentRow_ As Long
Private startTime_ As Double
Private stopTime_ As Double
Private browser_ As Selenium.WebDriver
Private destinations_ As Collection
Private originalTextColumn As Long
Private translatedTextColumn As Long

Public Sub Initialize(ByRef dstSheet As Worksheet)

    Set destinations_ = New Collection
    Set srcSheet_ = ThisWorkbook.Worksheets(1)
    Set dstSheet_ = dstSheet
    lastRow_ = GetLastRow(srcSheet_)
    currentRow_ = 2
    If ThisWorkbook.Worksheets(1).LeftOriginalTextRadio.Value Then
        originalTextColumn = 1
        translatedTextColumn = 2
    Else
        originalTextColumn = 2
        translatedTextColumn = 1
    End If
    Call FormatDstSheet
    Call CopyOriginalText

End Sub

Private Sub CopyOriginalText()

    Dim virtualCell As Variant
    virtualCell = srcSheet_.Range(srcSheet_.Cells(ROW_START, COL_START), srcSheet_.Cells(lastRow_, COL_START))
    dstSheet_.Range(dstSheet_.Cells(ROW_START, originalTextColumn), dstSheet_.Cells(lastRow_, originalTextColumn)) = virtualCell

End Sub

Private Sub FormatDstSheet()

    With dstSheet_.Columns(originalTextColumn)
        .ColumnWidth = 90
        .Font.Name = "UD デジタル 教科書体 N-R"
        .NumberFormatLocal = "@"
        .Font.Size = 14
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With

    With dstSheet_.Columns(translatedTextColumn)
        .ColumnWidth = 90
        .Font.Name = "UD デジタル 教科書体 N-R"
        .NumberFormatLocal = "@"
        .Font.Size = 14
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    
    With dstSheet_.Cells(ROW_TITLE, originalTextColumn)
        .Value = "原文"
        .HorizontalAlignment = xlCenter
        .RowHeight = 50
    End With
    
    With dstSheet_.Cells(ROW_TITLE, translatedTextColumn)
        .Value = "翻訳文"
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
            targetText = Trim(WorksheetFunction.EncodeURL(dstSheet_.Cells(currentRow_, originalTextColumn).Value))
            If Len(Trim(targetText)) <> 0 Then
                Dim myDestination As Destination: Set myDestination = New Destination
                Call myDestination.Initialize(dstSheet_.Cells(currentRow_, originalTextColumn).Value, CStr(TranslateDeepL(targetText)))
                destinations_.Add myDestination
                dstSheet_.Cells(currentRow_, translatedTextColumn).Value = myDestination.translatedText_
            End If
            currentRow_ = currentRow_ + 1
            stopTime_ = Timer
            If (stopTime_ - startTime_ > 300) Then
                GoTo Continue
            End If
        Loop
Continue:
        browser_.Quit
        Debug.Print (stopTime_ - startTime_)
        Debug.Print (currentRow_)
    Loop While (currentRow_ <= lastRow_)

End Sub

Private Function TranslateDeepL(ByVal targetText As String) As String

    Dim parameter As String
    parameter = targetText
    parameter = Replace(parameter, "%2F", "%5C%2F")
    parameter = Replace(parameter, "%26", "%5C%26")
    parameter = Replace(parameter, "%81F", "%5C%81F")
    parameter = Replace(parameter, "%81G", "%5C%81G")
    parameter = Replace(parameter, "%81H", "%5C%81H")
    parameter = Replace(parameter, "%7C", "%5C%7C")
    browser_.Get URL_TRANSLATOR & parameter
    Dim s As String
    Do
        DoEvents
        s = CStr(browser_.FindElementById(TARGET_ELEMENT_ID).Attribute(OUTPUT_ATTRIBUTE))
        s = Replace(s, vbCrLf, "")
        browser_.Wait WAIT_TIME
    Loop Until (Len(s) <> 0)
    TranslateDeepL = CStr(browser_.FindElementById(TARGET_ELEMENT_ID).Attribute(OUTPUT_ATTRIBUTE))

End Function

Public Sub DrawLine()

    Dim row As Long
    For row = ROW_START To ROW_START + lastRow_ - 1
        If dstSheet_.Rows(row).RowHeight <= MAX_ROW_HEIGHT - MARGIN_HEIGHT Then
            dstSheet_.Rows(row).RowHeight = dstSheet_.Rows(row).RowHeight + MARGIN_HEIGHT
        End If
    Next
    If (originalTextColumn < translatedTextColumn) Then
        dstSheet_.Range(dstSheet_.Cells(ROW_TITLE, originalTextColumn), dstSheet_.Cells(lastRow_, translatedTextColumn)).Borders.LineStyle = xlContinuous
        Exit Sub
    End If
    dstSheet_.Range(dstSheet_.Cells(ROW_TITLE, translatedTextColumn), dstSheet_.Cells(lastRow_, originalTextColumn)).Borders.LineStyle = xlContinuous

End Sub

