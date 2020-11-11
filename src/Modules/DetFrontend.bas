Attribute VB_Name = "DetFrontend"
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

Public Sub selectLine(ByVal isRow As Boolean, ByVal n As Integer)
    resetSelectionWithout isRow, n
    ListDet.selectedLine = n
    ListDet.isRowSelected = isRow
    Dim alphabet As Variant
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")
    If Not isRow Then
        selectColumn (alphabet(n + 1))
    Else
        selectRow (n + 2)
    End If
End Sub

Public Sub cmdsPaste_Click(ByVal isRow As Boolean, ByVal n As Integer)
    If getSetting(stgHasMatrixOptimized) = 0 And (ListDet.isRowSelected Eqv isRow) And Not (ListDet.selectedLine = n Or ListDet.selectedLine = -1) Then
        detSumToLine isRow, ListDet.selectedLine, n, Range(getSetting(stgAddrFactor)).value
        checkOptimization
    End If
End Sub

Public Sub tbtsSelection_Click(ByVal isRow As Boolean, ByVal n As Integer, ByVal obj_ As Object)
    If obj_.value Then
        selectLine isRow, n
    Else
        resetMatrixStyle
    End If
End Sub

Public Sub resetMatrixStyle()
    With ListDet
        .selectedLine = -1
        .Range(getSetting(stgRangeMatrix)).Interior.Color = xlNone
    End With
End Sub

Public Sub resetSelectionWithout(ByVal isRow As Boolean, ByVal n As Integer)
    resetMatrixStyle
    If isRow Then
        With ListDet
            .tbtSelCol1.value = False
            .tbtSelCol2.value = False
            .tbtSelCol3.value = False
            .tbtSelCol4.value = False
        End With
        If n = 1 Then
            With ListDet
                .tbtSelRow2.value = False
                .tbtSelRow3.value = False
                .tbtSelRow4.value = False
            End With
        End If
        If n = 2 Then
            With ListDet
                .tbtSelRow1.value = False
                .tbtSelRow3.value = False
                .tbtSelRow4.value = False
            End With
        End If
        If n = 3 Then
            With ListDet
                .tbtSelRow1.value = False
                .tbtSelRow2.value = False
                .tbtSelRow4.value = False
            End With
        End If
        If n = 4 Then
            With ListDet
                .tbtSelRow1.value = False
                .tbtSelRow2.value = False
                .tbtSelRow3.value = False
            End With
        End If
    Else
        With ListDet
            .tbtSelRow1.value = False
            .tbtSelRow2.value = False
            .tbtSelRow3.value = False
            .tbtSelRow4.value = False
        End With
        If n = 1 Then
            With ListDet
                .tbtSelCol2.value = False
                .tbtSelCol3.value = False
                .tbtSelCol4.value = False
            End With
        End If
        If n = 2 Then
            With ListDet
                .tbtSelCol1.value = False
                .tbtSelCol3.value = False
                .tbtSelCol4.value = False
            End With
        End If
        If n = 3 Then
            With ListDet
                .tbtSelCol1.value = False
                .tbtSelCol2.value = False
                .tbtSelCol4.value = False
            End With
        End If
        If n = 4 Then
            With ListDet
                .tbtSelCol1.value = False
                .tbtSelCol2.value = False
                .tbtSelCol3.value = False
            End With
        End If
    End If
End Sub

Public Sub setEnabledSelectors(ByVal isEnabled As Boolean)
    With ListDet
        .tbtSelRow1.Enabled = isEnabled
        .tbtSelRow2.Enabled = isEnabled
        .tbtSelRow3.Enabled = isEnabled
        .tbtSelRow4.Enabled = isEnabled
        .tbtSelCol1.Enabled = isEnabled
        .tbtSelCol2.Enabled = isEnabled
        .tbtSelCol3.Enabled = isEnabled
        .tbtSelCol4.Enabled = isEnabled

        .cmdPasteCol1.Enabled = isEnabled
        .cmdPasteCol2.Enabled = isEnabled
        .cmdPasteCol3.Enabled = isEnabled
        .cmdPasteCol4.Enabled = isEnabled
        .cmdPasteRow1.Enabled = isEnabled
        .cmdPasteRow2.Enabled = isEnabled
        .cmdPasteRow3.Enabled = isEnabled
        .cmdPasteRow4.Enabled = isEnabled
    End With
End Sub

Public Sub resetGame()
    setSetting stgSolZerosCount, 0
    setSetting stgSolIsRows, 0
    setSetting stgHasMatrixOptimized, 0

    Range(getSetting(stgRangeMatrix)).value = 0 'Zero fill
    ListDet.cmdCheckDet.Enabled = False
    setEnabledSelectors False
    With Range(getSetting(stgAddrResultsCell))
        .value = ""
        .Interior.Color = xlNone
    End With
    Range(getSetting(stgAddrOptimizedStatusCell)).value = ""
    Range(getSetting(stgAddrAnswer)).value = ""
    Range(getSetting(stgAddrZeros)).value = ""
    Range(getSetting(stgAddrDirection)).value = ""
    Range(getSetting(stgAddrFactor)).value = 1
    With ListDet
        .tbtSelCol1.value = False
        .tbtSelCol2.value = False
        .tbtSelCol3.value = False
        .tbtSelCol4.value = False
        .tbtSelRow1.value = False
        .tbtSelRow2.value = False
        .tbtSelRow3.value = False
        .tbtSelRow4.value = False
    End With
End Sub

Public Sub selectColumn(ByVal Col As String)
    Range(Col + "3:" + Col + "6").Interior.Color = Common.getSetting(stgSelColor)
End Sub

Public Sub selectRow(ByVal Row As Integer)
    Range("C" & Val(Row) & ":F" & Val(Row)).Interior.Color = Common.getSetting(stgSelColor)
End Sub

Public Sub checkOptimization()
    If DetOperations.isOptimizedByTask Then
        With ListDet
            .Range(getSetting(stgAddrOptimizedStatusCell)).value = getSetting(stgStrMatrixHasOptimized)
            .cmdCheckDet.Enabled = True
        End With
        setEnabledSelectors False
        setSetting stgHasMatrixOptimized, 1
    End If
End Sub
