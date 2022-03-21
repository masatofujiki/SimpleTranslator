VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextExtracctor"
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

Private Const URL_GOOGLE_PATENTS As String = "https://patents.google.com/"
Private Const URL_GOOGLE_PATENTS_PATENT As String = "https://patents.google.com/patent/"
Private Const WAIT_TIME As Long = 3000

Private publicationNumber_ As String
Private currentRow_ As Long
Private descriptionStartRow_ As Long
Private driver_ As WebDriver


Public Sub Initialize(ByVal publicationNumber As String)

    Set driver_ = New WebDriver
    publicationNumber_ = publicationNumber
    currentRow_ = 2

End Sub

Public Function IsReady() As Boolean

'    SafeOpen driver_, Chrome
    SafeOpenHeadless driver_, Chrome
    driver_.Navigate URL_GOOGLE_PATENTS

    If Not PublicationNumberIsExist Then
        driver_.Shutdown
        IsReady = False
        Exit Function
    End If
    IsReady = True

End Function

Public Sub ExtractPatentText()

    Call driver_.Navigate(URL_GOOGLE_PATENTS_PATENT & StrConv(publicationNumber_, vbUpperCase + vbNarrow))
    Call ExtractTitleOfInvention
    Call ExtractAbstract
    Call ExtractDescription
    driver_.Shutdown

End Sub

Private Sub ExtractDescription()

    descriptionStartRow_ = currentRow_
    Call ExtractDescriptionParagraph
    currentRow_ = descriptionStartRow_
    Call ExtractDescriptionLine

End Sub

Private Sub ExtractDescriptionLine()

    If driver_.ExecuteScript("return !!document.getElementsByClassName('description-line')[0]") Then
        Cells(currentRow_, 1).value = "Description"
        currentRow_ = currentRow_ + 1
        Dim elements() As WebElement
        elements = driver_.FindElements(By.ClassName, "description-line")
        Dim i As Long
        For i = LBound(elements) To UBound(elements)
            Dim paragraphNumber As String
            paragraphNumber = CStr(driver_.ExecuteScript("return (document.getElementsByClassName('description-line')[" & i & "]).getAttribute('num')"))
            If paragraphNumber <> "0000" Then
                paragraphNumber = "[" + Replace(CStr(paragraphNumber), "p", "") + "]"
            End If
            Cells(currentRow_, 1).value = paragraphNumber + vbLf + elements(i).GetText
            currentRow_ = currentRow_ + 1
        Next
    End If

End Sub

Private Sub ExtractDescriptionParagraph()

    If driver_.ExecuteScript("return !!document.getElementsByClassName('description-paragraph')[0]") Then
        Cells(currentRow_, 1).value = "Description"
        currentRow_ = currentRow_ + 1
        Dim elements() As WebElement
        elements = driver_.FindElements(By.ClassName, "description-paragraph")
        Dim i As Long
        For i = LBound(elements) To UBound(elements)
            Cells(currentRow_, 1).value = elements(i).GetText
            currentRow_ = currentRow_ + 1
        Next
    End If

End Sub

Private Sub ExtractAbstract()

    If driver_.ExecuteScript("return !!document.getElementsByClassName('abstract')[0]") Then
        Cells(currentRow_, 1).value = "Abstract"
        currentRow_ = currentRow_ + 1
        Dim abstract As String
        abstract = driver_.ExecuteScript("return document.getElementsByClassName('abstract')[0]").GetText
        Cells(currentRow_, 1).value = abstract
        currentRow_ = currentRow_ + 1
    End If
    
End Sub

Private Sub ExtractTitleOfInvention()

    If driver_.ExecuteScript("return !!document.getElementsByTagName('h1')[1]") Then
        Cells(currentRow_, 1).value = driver_.ExecuteScript("return document.getElementsByTagName('h1')[1]").GetText
        currentRow_ = currentRow_ + 1
    End If
    
End Sub

Private Function PublicationNumberIsExist() As Boolean

    Dim element As WebElement
    Set element = driver_.FindElement(By.ID, "searchInput")
    element.SetValue publicationNumber_
    driver_.FindElement(By.ID, "searchButton").Click
    Call Sleep(WAIT_TIME)
    Dim currentUrl As String
    currentUrl = driver_.ExecuteScript("return location.href")
    If currentUrl = URL_GOOGLE_PATENTS Then
        PublicationNumberIsExist = False
        Exit Function
    End If
    PublicationNumberIsExist = True

End Function
