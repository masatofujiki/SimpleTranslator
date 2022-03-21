VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "SimpleTranslator.xlsm.ProgressBarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public ProgressIsCancel As Boolean

Public Sub InitializeProgress(ByVal currentRow As Long, ByVal lastRow As Long)

    ProgressIsCancel = False
    ProgressBar1.value = 0
    ProgressBar1.Min = 0
    ProgressBar1.Max = lastRow
    ProgressLabel.Caption = "翻訳実行中"
    Caption = "翻訳実行中"

End Sub

Private Sub CancelButton_Click()

    ProgressIsCancel = True

End Sub

Public Sub UpdateProgress(ByVal currentRow As Long, ByVal lastRow As Long)

    ProgressBar1.value = currentRow - ROW_OFFSET
    ProgressLabel = "翻訳実行中 " & CStr(Int((currentRow - TranslatorSheet.ROW_OFFSET) / (lastRow - TranslatorSheet.ROW_OFFSET) * 100)) & "% 完了"
    DoEvents

End Sub

