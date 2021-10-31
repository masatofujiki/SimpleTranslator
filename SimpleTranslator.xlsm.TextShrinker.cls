VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextShrinker"
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

Private targetCol_ As Long
Private lastRow As Long
Private allText As String
Private beforeDescriptions As Variant
Private afterDescriptions As Variant
Private srcSheet_ As Worksheet

Public Sub Initialize(ByVal targetCol As Long)

    targetCol_ = targetCol
    Set srcSheet_ = ThisWorkbook.Worksheets(1)
    lastRow = GetLastRow(srcSheet_)

End Sub

Public Sub ShrinkNumberedText()

    Call ShrinkText("(?:\[|Åy)(\d{1,8})(?:\]|Åz)", vbTab + "Åy" + "$1" + "Åz")

End Sub

Public Sub ShrinkNormalText()

    Call ShrinkText("\n+", vbTab)

End Sub

Private Sub ShrinkText(ByVal splitPattern As String, ByVal replaceToken As String)

    Dim reg As RegExp: Set reg = New RegExp
    reg.IgnoreCase = True
    reg.Global = True
    reg.Pattern = splitPattern
    Dim allDescriptions As Variant
    afterDescriptions = Split(reg.Replace(allText, replaceToken), vbTab)

    Dim cellData As Variant
    cellData = srcSheet_.Range(srcSheet_.Cells(LBound(afterDescriptions) + ROW_START - 1, targetCol_), srcSheet_.Cells(UBound(afterDescriptions) + ROW_START - 1, targetCol_))
    Dim i As Long
    For i = LBound(afterDescriptions) To UBound(afterDescriptions)
        If Len(Trim(TrimLF(afterDescriptions(i)))) > 0 Then
            cellData(i + 1, 1) = Trim(TrimLF(afterDescriptions(i)))
        End If
    Next

    srcSheet_.Range(srcSheet_.Cells(ROW_START, targetCol_), srcSheet_.Cells(lastRow, targetCol_)).ClearContents
    srcSheet_.Range(srcSheet_.Cells(LBound(afterDescriptions) + ROW_START, targetCol_), srcSheet_.Cells(UBound(afterDescriptions) + ROW_START, targetCol_)) = cellData
    srcSheet_.Rows(CStr(ROW_START) + ":" + CStr(UBound(afterDescriptions) + ROW_START)).AutoFit
    
    If srcSheet_.Cells(ROW_START, COL_START).Value = "" Then
        srcSheet_.Rows(ROW_START).Delete
    End If

End Sub

Public Sub FormatPatentText()
    
    If HasNoText() Then
        MsgBox "å¥ï∂Ç™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    
    If HasOneText() Then
        allText = TrimLF(srcSheet_.Cells(ROW_START, 1))
        srcSheet_.Cells(ROW_START, 1).Value = allText
        Rows(CStr(ROW_START) + ":" + CStr(ROW_START)).AutoFit
        Exit Sub
    End If

    beforeDescriptions = srcSheet_.Range(srcSheet_.Cells(ROW_START, targetCol_), srcSheet_.Cells(lastRow, targetCol_))
    Dim sb As StringBuilder: Set sb = New StringBuilder
    sb.Initialize
    Dim i As Long
    For i = LBound(beforeDescriptions) To UBound(beforeDescriptions)
        If Len(Replace(beforeDescriptions(i, 1), ChrW(160), "")) > 0 Then
            sb.Add CStr(beforeDescriptions(i, 1)) + vbLf
        End If
    Next
    allText = TrimLF(sb.ToText)

    If IsNumbered Then
        Call ShrinkNumberedText
        Exit Sub
    End If
    Call ShrinkNormalText

End Sub

Public Sub FormatText()
    
    If HasNoText() Then
        MsgBox "å¥ï∂Ç™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    
    If HasOneText() Then
        allText = TrimLF(srcSheet_.Cells(ROW_START, 1))
        srcSheet_.Cells(ROW_START, 1).Value = allText
        Rows(CStr(ROW_START) + ":" + CStr(ROW_START)).AutoFit
        Exit Sub
    End If

    beforeDescriptions = srcSheet_.Range(srcSheet_.Cells(ROW_START, targetCol_), srcSheet_.Cells(lastRow, targetCol_))
    Dim sb As StringBuilder: Set sb = New StringBuilder
    sb.Initialize
    Dim i As Long
    For i = LBound(beforeDescriptions) To UBound(beforeDescriptions)
        If Len(Replace(beforeDescriptions(i, 1), ChrW(160), "")) > 0 Then
            sb.Add CStr(beforeDescriptions(i, 1)) + vbLf
        End If
    Next
    allText = TrimLF(sb.ToText)
    Call ShrinkNormalText

End Sub

Public Function HasNoText() As Boolean

    HasNoText = (lastRow <= 1)

End Function

Public Function HasOneText() As Boolean

    HasOneText = (lastRow = ROW_START)

End Function

Public Function HasSomeText() As Boolean

    HasSomeText = (lastRow > ROW_START)

End Function

Public Function IsNumbered() As Boolean

    Dim reg As RegExp: Set reg = New RegExp
    reg.IgnoreCase = True
    reg.Global = True
    reg.Pattern = "(?:\[|Åy)(\d{1,8})(?:\]|Åz)"
    IsNumbered = reg.Test(allText)

End Function
