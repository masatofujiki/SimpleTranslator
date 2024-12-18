VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HtmlTranslator"
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
Private Const URL_TRANSLATOR As String = "https://www.deepl.com/ja/translator#en/"
Private Const WAIT_BROWSER_START_UP_TIME = 3000
Private Const WAIT_MILLISECONDS As Long = 500
Private Const SECOND_TO_RESTART As Long = 300
Private Const MAX_RETRY_COUNT As Long = 100

Private srcSheet_ As Worksheet
Private dstSheet_ As Worksheet
Private lastRow_ As Long
Private currentRow_ As Long
Private startTime_ As Double
Private stopTime_ As Double
Private driver_ As WebDriver
Private targets_ As Collection
Private fileName_ As String
Private originalTextColumn As Long
Private translatedTextColumn As Long
Private outputCombobox_ As MSForms.ComboBox
Private originalTextTitle_ As String
Private translatedTextTitle_ As String
Private outputLanguage_ As String

Public Sub Initialize(ByVal fileName As String)

    Set targets_ = New Collection
    Set srcSheet_ = ThisWorkbook.Worksheets(1)
    lastRow_ = GetLastRow(srcSheet_)
    currentRow_ = 2
    Set outputCombobox_ = ThisWorkbook.Worksheets(1).OutputComboBox
    outputLanguage_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_OUTPUT_LANGUAGE)
    originalTextTitle_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_ORIGINAL_TEXT)
    translatedTextTitle_ = outputCombobox_.List(outputCombobox_.ListIndex, COL_TRANSLATED_TEXT)
    If ThisWorkbook.Worksheets(1).LeftOriginalTextRadio.value Then
        originalTextColumn = 1
        translatedTextColumn = 2
    Else
        originalTextColumn = 2
        translatedTextColumn = 1
    End If
    fileName_ = fileName

End Sub

Public Function HasFileName() As Boolean

    HasFileName = Len(fileName_) > 0

End Function

Public Sub TransLateToHtml()

    Call ProgressBarForm.InitializeProgress(currentRow_, lastRow_)
    Call ProgressBarForm.Show(vbModeless)
    DoEvents
    Do
        Set driver_ = New WebDriver
        'SafeOpenHeadless driver_, Chrome
        SafeOpen driver_, Chrome, CapabilityArgs:="--headless --disable-gpu"
        startTime_ = Timer
        Do While (currentRow_ <= lastRow_)
            Dim targetText As String
            targetText = Trim(EncodeBasedOnExcelVersion(srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).value))
            If Len(Trim(targetText)) <> 0 Then
                Dim myTarget As Target: Set myTarget = New Target
                Dim src As String
                src = srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).value
                Dim dst As String
                dst = TranslateDeepL(targetText)
'                dst = TranslateGoogle(targetText)
                Call myTarget.Initialize(src, dst)
                targets_.Add myTarget
                Call ProgressBarForm.UpdateProgress(currentRow_, lastRow_)
            End If
            currentRow_ = currentRow_ + 1
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
        Debug.Print ("完了行：" & currentRow_ - TranslatorSheet.ROW_START)
    Loop While (currentRow_ <= lastRow_ And ProgressBarForm.ProgressIsCancel = False)
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
'    parameter = Replace(parameter, "%E2%80%A6", "%E3%83%BB%E3%83%BB%E3%83%BB")
    Call driver_.Navigate(URL_TRANSLATOR & outputLanguage_ & "/" & parameter)
    Dim retryCount As Long
    retryCount = 0
    Call Sleep(WAIT_BROWSER_START_UP_TIME)
    Dim s As String
    Do
'        s = CStr(Trim(driver_.ExecuteScript("return document.getElementById('target-dummydiv').textContent")))
        s = CStr(Trim(driver_.ExecuteScript("return document.getElementsByClassName('lmt__translations_as_text__text_btn')[0].textContent")))
        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, " ", "")
        s = Replace(s, "　", "")
        retryCount = retryCount + 1
CONTINUE:
        Call Sleep(WAIT_MILLISECONDS)
    Loop Until (Len(s) <> 0 Or retryCount >= MAX_RETRY_COUNT)
'    TranslateDeepL = CStr(driver_.ExecuteScript("return document.getElementById('target-dummydiv').textContent"))
    TranslateDeepL = CStr(Trim(driver_.ExecuteScript("return document.getElementsByClassName('lmt__translations_as_text__text_btn')[0].textContent")))
End Function

Private Function TranslateGoogle(ByVal targetText As String) As String

    Dim parameter As String
    parameter = targetText
    parameter = Replace(parameter, "%2F", "%5C%2F")
'    parameter = Replace(parameter, "%26", "%5C%26")
    parameter = Replace(parameter, "%81F", "%5C%81F")
    parameter = Replace(parameter, "%81G", "%5C%81G")
    parameter = Replace(parameter, "%81H", "%5C%81H")
    parameter = Replace(parameter, "%7C", "%5C%7C")
    Call driver_.Navigate("https://translate.google.co.jp/?hl=ja&sl=en&tl=ja&text=" & outputLanguage_ & "/" & parameter & "&op=translate")
    Dim retryCount As Long
    retryCount = 0
    Dim s As String
    Do
'        driver_.FindElement(By.CssSelector, span[jsname='W297wb'])
        s = CStr(Trim(driver_.ExecuteScript("return document.querySelector(""span[jsname='W297wb']"")").GetText))
        s = Replace(s, vbCrLf, "")
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, " ", "")
        s = Replace(s, "　", "")
        Call Sleep(WAIT_MILLISECONDS)
        retryCount = retryCount + 1
    Loop Until (Len(s) <> 0 Or retryCount >= MAX_RETRY_COUNT)
    TranslateGoogle = CStr(Trim(driver_.ExecuteScript("return document.querySelector(""span[jsname='W297wb']"")").GetText))

End Function


Public Sub CreateHtml()

    Dim sb As StringBuilder
    Set sb = New StringBuilder
    Call CreateHeader(sb)
    Call CreateContents(sb)
    Call CreateFooter(sb)

'    Dim filePath As String
'    filePath = ThisWorkbook.Path & "\" & fileName_ & ".html"

    On Error GoTo myError
    
    Dim adoStream As ADODB.Stream
    Set adoStream = New ADODB.Stream
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = 10
    adoStream.Open
    adoStream.WriteText sb.ToText
    adoStream.SaveToFile fileName_, 2
    adoStream.Close
    
    Dim wsh As WshShell
    Set wsh = New WshShell
    wsh.Run ("chrome.exe -url " & fileName_)
    Set wsh = Nothing
    
    Exit Sub
    
myError:
    MsgBox "ファイルを開けません", vbExclamation
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub CreateHeader(ByRef sb As StringBuilder)

    sb.Add "<!DOCTYPE html>"
    sb.Add "<html lang=""ja"">"
    sb.Add "<head>"
    sb.Add "<meta charset=""UTF-8"">"
    sb.Add "<link rel=""stylesheet"" href=""translator.css"">"
    sb.Add "<title>"
    sb.Add fileName_
    sb.Add "</title>"
    sb.Add "</head>"
    sb.Add "<body>"
    sb.Add "<table>"
    sb.Add "<tr>"
    If (originalTextColumn = 1) Then
        sb.Add "<th class=""first"">" & originalTextTitle_ & "</th>"
        sb.Add "<th class=""second"">" & translatedTextTitle_ & "</th>"
    Else
        sb.Add "<th class=""first"">" & translatedTextTitle_ & "</th>"
        sb.Add "<th class=""second"">" & originalTextTitle_ & "</th>"
    End If
    sb.Add "</tr>"
    
End Sub

Private Sub CreateContents(ByRef sb As StringBuilder)

    Dim myTarget As Target
    For Each myTarget In targets_
        sb.Add "<tr>"
        sb.Add "<td>"
        Dim originalText As String
        originalText = Replace(myTarget.originalText_, "<", "&lt;")
        originalText = Replace(originalText, ">", "&gt;")
        originalText = Replace(originalText, vbCrLf, "<br>")
        originalText = Replace(originalText, vbLf, "<br>")
        originalText = Replace(originalText, "&", "&amp")
        originalText = Replace(originalText, """", "&quot;")
        Dim translatedText As String
        translatedText = Replace(myTarget.translatedText_, "<", "&lt;")
        translatedText = Replace(translatedText, ">", "&gt;")
        translatedText = Replace(translatedText, vbCrLf, "<br>")
        translatedText = Replace(translatedText, vbLf, "<br>")
        translatedText = Replace(translatedText, "&", "&amp")
        translatedText = Replace(translatedText, """", "&quot;")
        If (originalTextColumn = 1) Then
            sb.Add originalText
            sb.Add "</td>"
            sb.Add "<td>"
            sb.Add translatedText
        Else
            sb.Add translatedText
            sb.Add "</td>"
            sb.Add "<td>"
            sb.Add originalText
        End If
        sb.Add "</td>"
        sb.Add "</tr>"
    Next

End Sub

Private Sub CreateFooter(ByRef sb As StringBuilder)

    sb.Add "</table>"
    sb.Add "</body>"
    sb.Add "</html>"

End Sub

