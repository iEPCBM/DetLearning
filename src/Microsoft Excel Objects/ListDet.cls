VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2020, Rishat D. Kagirov (iEPCBM)
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without modification,
' are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice, this
' list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation and/or
' other materials provided with the distribution.
'
' 3. Neither the name of the copyright holder nor the names of its contributors may
' be used to endorse or promote products derived from this software without specific
' prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
' IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
' INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
' NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
' OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
' WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY
' OF SUCH DAMAGE.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public selectedLine As Integer
Public isRowSelected As Boolean

Private Sub btGenerate_Click()
    Dim mgDetMatrix() As Long

    setSetting stgSolZerosCount, 0
    setSetting stgSolIsRows, 0
    setSetting stgHasMatrixOptimized, 0

    Range(getSetting(stgRangeMatrix)).value = 0 'Zero fill
    ListDet.cmdCheckDet.Enabled = False
    With Range(getSetting(stgAddrResultsCell))
        .value = ""
        .Interior.Color = xlNone
    End With
    Range(getSetting(stgAddrOptimizedStatusCell)).value = ""
    Range(getSetting(stgAddrAnswer)).value = ""
    
    Dim alphabet As Variant
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    mgDetMatrix = generateDet
    For i = 0 To 3
        For j = 0 To 3
            Range(alphabet(j + 2) & Val(i + 3)).value = mgDetMatrix(i, j)
        Next j
    Next i
    
    setEnabledSelectors True
    Range(getSetting(stgAddrZeros)).value = getSetting(stgSolZerosCount)
    If getSetting(stgSolIsRows) Then
        Range(getSetting(stgAddrDirection)).value = getSetting(stgStrRow)
    Else
        Range(getSetting(stgAddrDirection)).value = getSetting(stgStrColumn)
    End If
End Sub

Private Sub btSettings_Click()
    Settings.Show
End Sub

Private Sub bttSel_Click()
    resetMatrixStyle
    If bttSel.value Eqv True Then
        selectColumn ("D")
    End If
End Sub

Private Sub cmdAboutProg_Click()
    formAbout.Show
End Sub

Private Sub cmdCheckDet_Click()
    If Val(Range(getSetting(stgAddrAnswer)).value) = Val(getDeterminant) Then
        With ListCongratulation
            .Visible = xlSheetVisible
            .Activate
        End With
        With Range(getSetting(stgAddrResultsCell))
            .value = getSetting(stgStrAnswerTrue)
            .Interior.Color = getSetting(stgColorAnswerTrue)
        End With
        cmdCheckDet.Enabled = False
    Else
        With Range(getSetting(stgAddrResultsCell))
            .value = getSetting(stgStrAnswerWrong)
            .Interior.Color = getSetting(stgColorAnswerWrong)
        End With
    End If
End Sub

Private Sub cmdPasteCol1_Click()
    cmdsPaste_Click False, 1
End Sub

Private Sub cmdPasteCol2_Click()
    cmdsPaste_Click False, 2
End Sub

Private Sub cmdPasteCol3_Click()
    cmdsPaste_Click False, 3
End Sub

Private Sub cmdPasteCol4_Click()
    cmdsPaste_Click False, 4
End Sub

Private Sub cmdPasteRow1_Click()
    cmdsPaste_Click True, 1
End Sub

Private Sub cmdPasteRow2_Click()
    cmdsPaste_Click True, 2
End Sub

Private Sub cmdPasteRow3_Click()
    cmdsPaste_Click True, 3
End Sub

Private Sub cmdPasteRow4_Click()
    cmdsPaste_Click True, 4
End Sub

Private Sub cmdResetGame_Click()
    resetGame
End Sub

Private Sub tbtSelCol1_Click()
    tbtsSelection_Click False, 1, tbtSelCol1
End Sub

Private Sub tbtSelCol2_Click()
    tbtsSelection_Click False, 2, tbtSelCol2
End Sub

Private Sub tbtSelCol3_Click()
    tbtsSelection_Click False, 3, tbtSelCol3
End Sub

Private Sub tbtSelCol4_Click()
    tbtsSelection_Click False, 4, tbtSelCol4
End Sub

Private Sub tbtSelRow1_Click()
    tbtsSelection_Click True, 1, tbtSelRow1
End Sub

Private Sub tbtSelRow2_Click()
    tbtsSelection_Click True, 2, tbtSelRow2
End Sub

Private Sub tbtSelRow3_Click()
    tbtsSelection_Click True, 3, tbtSelRow3
End Sub

Private Sub tbtSelRow4_Click()
    tbtsSelection_Click True, 4, tbtSelRow4
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not IsNumeric(Range(getSetting(stgAddrFactor)).value) Then
        Range(getSetting(stgAddrFactor)).value = 1
    End If
End Sub
