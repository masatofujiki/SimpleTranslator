VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColouringFragmentedSentences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FIRST_ROW As Long = 2
Private Const SAFE_COLOR As Long = 16764006
Private Const CAUTION_COLOR As Long = 262143
Private Const ALERT_COLOR As Long = 16764159

Private srcSheet_ As Worksheet
Private currentRow_ As Long
Private lastRow_ As Long

Public Sub Initialize(ByVal lastRow As Long)

    Set srcSheet_ = ThisWorkbook.Worksheets(1)
    lastRow_ = lastRow

End Sub

Public Sub Paint()

    Call ClearCells
    Call PaintColor

End Sub

Private Sub ClearCells()

    currentRow_ = FIRST_ROW
    Do While (currentRow_ <= lastRow_)
        srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).Interior.ColorIndex = xlNone
        currentRow_ = currentRow_ + 1
    Loop

End Sub

Private Sub PaintColor()

    currentRow_ = FIRST_ROW
    Dim lastRow As String: lastRow = GetLastRow(srcSheet_)
    Do While (currentRow_ <= lastRow)
        Dim targetText As String
        targetText = srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).value
        If IsEndMark(targetText) Then
            GoTo CONTINUE
        End If
        If Right(targetText, 1) = "," Then
            srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).Interior.Color = ALERT_COLOR
            GoTo CONTINUE
        End If
        Dim nextTargetText As String
        nextTargetText = srcSheet_.Cells(currentRow_ + 1, TranslatorSheet.COL_SRC_ORIGINAL).value
        If UCase(Left(nextTargetText, 1)) = Left(nextTargetText, 1) Then
            srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).Interior.Color = SAFE_COLOR
            GoTo CONTINUE
        End If
        srcSheet_.Cells(currentRow_, TranslatorSheet.COL_SRC_ORIGINAL).Interior.Color = CAUTION_COLOR
CONTINUE:
        currentRow_ = currentRow_ + 1
    Loop

End Sub

Private Function IsEndMark(ByVal targetText As String) As Boolean

    Dim reg As RegExp: Set reg = New RegExp
    reg.IgnoreCase = True
    reg.Global = True
    reg.Pattern = "(\.""?|\?""?|:)$"

    If (reg.Test(targetText)) Then
        IsEndMark = True
        Exit Function
    End If
    IsEndMark = False

'    If Right(targetText, 1) = "." Or Right(targetText, 2) = ".""" Or Right(targetText, 2) = ".)" Or _
'       Right(targetText, 1) = "?" Or Right(targetText, 2) = "?""" Or Right(targetText, 2) = "?)" Or _
'       Right(targetText, 1) = ":" Then
'        IsEndMark = True
'        Exit Function
'    End If
'    IsEndMark = False

End Function

