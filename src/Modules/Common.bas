Attribute VB_Name = "Common"
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

Public Enum SETTING
    stgSelColor = 1
    stgPRNG
    stgMinZeros
    stgMaxZeros
    stgGenIterations
    stgTimer
    stgRangeMatrix
    stgAddrZeros
    stgAddrDirection
    stgAddrFactor
    stgStrColumn
    stgStrRow
    stgAddrAnswer
    stgAddrResultsCell
    stgStrAnswerTrue
    stgStrAnswerWrong
    stgColorAnswerTrue
    stgColorAnswerWrong
    stgAboutTitle
    stgAboutAutor
    stgAboutVersion
    stgAboutVersionCode
    stgStrMatrixHasOptimized
    stgAddrOptimizedStatusCell
    stgHasMatrixOptimized
    stgSolIsRows
    stgSolZerosCount
End Enum

Public Function getSetting(ByVal stg As SETTING) As String
    getSetting = listSettings.Range("B" & Val(stg)).value
End Function

Public Sub setSetting(ByVal stg As SETTING, ByVal strVal As String)
    listSettings.Range("B" & Val(stg)).value = strVal
End Sub

Public Sub displayMatrix(matrix() As Integer)
    Dim dStr As String
    dStr = "Matrix" + Chr$(13) & Chr$(10)
    For i = 0 To 3
        For j = 0 To 3
            dStr = dStr + CStr(matrix(i, j)) + "    "
        Next j
        dStr = dStr + Chr$(13) & Chr$(10) + Chr$(13) & Chr$(10)
    Next i
    MsgBox (dStr)
End Sub
