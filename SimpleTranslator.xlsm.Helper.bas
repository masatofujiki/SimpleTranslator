Attribute VB_Name = "Helper"
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

Public Function GetFilePath() As String

    Dim fileName As Variant
    fileName = Application.GetSaveAsFilename(InitialFileName:=Application.ThisWorkbook.Path & "\" & _
                                             ThisWorkbook.Worksheets(1).PublicationNumberTextBox.Text, FileFilter:="HTMLファイル, *.html")
    If VarType(fileName) = vbBoolean Then
        GetFilePath = ""
        Exit Function
    End If
    GetFilePath = fileName

End Function

Public Function FileNameHasNGCharacters(ByVal fileName As String) As Boolean

    Dim reg As RegExp
    Set reg = New RegExp
    reg.Pattern = "[""\*\/:<=>?\[\]\\\n]"
    If reg.Test(fileName) Then
        FileNameHasNGCharacters = True
        Exit Function
    End If
    FileNameHasNGCharacters = False

End Function

Public Function TrimLF(ByVal TARGET As String) As String
  
  Dim temp As String: temp = TARGET
  Do Until Left(temp, 1) <> vbLf
    temp = Mid(temp, 2)
  Loop
  Do Until Right(temp, 1) <> vbLf
    temp = Left(temp, Len(temp) - 1)
  Loop
  Do Until Left(temp, 1) <> vbTab
    temp = Mid(temp, 2)
  Loop
  Do Until Right(temp, 1) <> vbTab
    temp = Left(temp, Len(temp) - 1)
  Loop
  TrimLF = temp

End Function

Public Function GetLastRow(ByRef ws As Worksheet) As Long

    GetLastRow = ws.Cells(Rows.Count, COL_SRC_ORIGINAL).End(xlUp).row

End Function

Public Function ReturnMax(ByVal value1 As Long, ByVal value2 As Long) As Long

    If value1 >= value2 Then
        ReturnMax = value1
        Exit Function
    End If
    ReturnMax = value2

End Function

Public Function CleanChar(strData As String) As String

    Dim result As String: result = ""
    Dim i As Long
    For i = 1 To Len(strData)
        Dim currentChar As String
        currentChar = Mid$(strData, i, 1)
        If Asc(currentChar) < 0 Or 32 <= Asc(currentChar) Then
            '漢字のAscの返り値はマイナスに留意
            result = result & currentChar
        End If
    Next
    CleanChar = result

End Function

