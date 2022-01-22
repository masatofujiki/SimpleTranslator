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
Private browser_ As Selenium.WebDriver

Public Sub Initialize(ByVal publicationNumber As String)

    Set browser_ = New Selenium.WebDriver
    publicationNumber_ = publicationNumber
    currentRow_ = 2

End Sub

Public Function IsReady() As Boolean

    Call browser_.AddArgument("--headless --disable-gpu")
    Call browser_.Start("chrome")
    Call browser_.Get(URL_GOOGLE_PATENTS)
    
    If Not PublicationNumberIsExist Then
        browser_.Quit
        IsReady = False
        Exit Function
    End If
    IsReady = True
    
End Function

Public Sub ExtractPatentText()

    Call browser_.Get(URL_GOOGLE_PATENTS_PATENT & StrConv(publicationNumber_, vbUpperCase + vbNarrow))
    Call ExtractTitleOfInvention
    Call ExtractAbstract
    Call ExtractDescription
    browser_.Quit
    
End Sub

Private Sub ExtractDescription()

    descriptionStartRow_ = currentRow_
    Call ExtractDescriptionParagraph
    currentRow_ = descriptionStartRow_
    Call ExtractDescriptionLine
    
End Sub

Private Sub ExtractDescriptionLine()

    Dim myBy As By
    Set myBy = New By
    If browser_.IsElementPresent(myBy.Class("description-line")) Then
        Cells(currentRow_, 1).Value = "Description"
        currentRow_ = currentRow_ + 1
        Dim elements As Selenium.WebElements
        Set elements = browser_.FindElementsByClass("description-line")
        Dim element As Selenium.WebElement
        For Each element In elements
            Dim paragraphNumber As String: paragraphNumber = ""
            If element.Attribute("num") <> "0000" Then
                paragraphNumber = "[" + Replace(CStr(element.Attribute("num")), "p", "") + "]"
            End If
            Cells(currentRow_, 1).Value = paragraphNumber + vbLf + element.Text
            currentRow_ = currentRow_ + 1
        Next
    End If

End Sub

Private Sub ExtractDescriptionParagraph()

    Dim myBy As By
    Set myBy = New By
    If browser_.IsElementPresent(myBy.Class("description-paragraph")) Then
        Cells(currentRow_, 1).Value = "Description"
        currentRow_ = currentRow_ + 1
        Dim elements As Selenium.WebElements
        Set elements = browser_.FindElementsByClass("description-paragraph")
        Dim element As Selenium.WebElement
        For Each element In elements
            Cells(currentRow_, 1).Value = element.Text
            currentRow_ = currentRow_ + 1
        Next
    End If

End Sub

Private Sub ExtractAbstract()

    Dim myBy As By
    Set myBy = New By
    If browser_.IsElementPresent(myBy.Class("abstract")) Then
        Cells(currentRow_, 1).Value = "Abstract"
        currentRow_ = currentRow_ + 1
        Dim element As WebElement
        Set element = browser_.FindElementByClass("abstract")
        Cells(currentRow_, 1).Value = element.Text
        currentRow_ = currentRow_ + 1
    End If
    
End Sub

Private Sub ExtractTitleOfInvention()

    Dim elements As WebElements
    Set elements = browser_.FindElementsByTag("h1")
    Cells(currentRow_, 1).Value = elements(2).Text
    currentRow_ = currentRow_ + 1

End Sub

Private Function PublicationNumberIsExist() As Boolean

    Dim element As Selenium.WebElement
    Set element = browser_.FindElementById("searchInput")
    Call element.SendKeys(publicationNumber_)
    Dim sendKey As Selenium.Keys
    Set sendKey = New Selenium.Keys
    Call element.SendKeys(sendKey.Enter)
    Call browser_.Wait(WAIT_TIME)
    If browser_.Url = URL_GOOGLE_PATENTS Then
        PublicationNumberIsExist = False
        Exit Function
    End If
    PublicationNumberIsExist = True
    
End Function
