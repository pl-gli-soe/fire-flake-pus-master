Attribute VB_Name = "RunFromMCDModule"
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Public Sub sDUNS(mcd As CommonData, r As Range)

    r = mcd.duns

End Sub

Public Sub sSUPPLIER(mcd As CommonData, r As Range)

    r = mcd.supplierName

End Sub

Public Sub sF_U(mcd As CommonData, r As Range)
    r = mcd.fupCode
End Sub

Public Sub sA(mcd As CommonData, r As Range)
    r = mcd.fmaFupCode
End Sub

Public Sub sMISC(mcd As CommonData, r As Range)
    r = mcd.misc
End Sub

Public Sub sDOH(mcd As CommonData, r As Range)
    r = mcd.doh
End Sub

Public Sub sOS(mcd As CommonData, r As Range)
    r = mcd.os
End Sub


Public Sub sBANK(mcd As CommonData, r As Range)
    r = mcd.bank
End Sub

Public Sub sBBAL(mcd As CommonData, r As Range)
    r = mcd.bbal
End Sub

Public Sub sCBAL(mcd As CommonData, r As Range)
    r = mcd.cbal
End Sub

Public Sub sPCS_TO_GO(mcd As CommonData, r As Range)
    r = mcd.pcsToGo
End Sub

Public Sub sDK(mcd As CommonData, r As Range)
End Sub

Public Sub sMODE(mcd As CommonData, r As Range)
    r = mcd.mode
End Sub

Public Sub sMNPC(mcd As CommonData, r As Range)
    
End Sub

Public Sub sNCX(mcd As CommonData, r As Range)
    
End Sub

Public Sub sOBS(mcd As CommonData, r As Range)

End Sub

Public Sub sSTD_PACK(mcd As CommonData, r As Range)
    r = mcd.stdPack
End Sub

Public Sub soneJOB(mcd As CommonData, r As Range)

End Sub

Public Sub sIP(mcd As CommonData, r As Range)
End Sub

Public Sub sCOUNT(mcd As CommonData, r As Range)
    r = mcd.count_cmnt
End Sub

Public Sub sO(mcd As CommonData, r As Range)
    r = mcd.o_cmnt
End Sub

Public Sub sF(mcd As CommonData, r As Range)
    r = mcd.f_cmnt
End Sub

Public Sub sPART_NAME(mcd As CommonData, r As Range)
    r = mcd.partName
End Sub

Public Sub sQHD(mcd As CommonData, r As Range)
    r = mcd.qhd
End Sub

Public Sub sTT(mcd As CommonData, r As Range)
    r = mcd.ttime
End Sub

Public Sub sLOG(mcd As CommonData, r As Range)
    r = mcd.errorLog
End Sub

Public Sub sC(mcd As CommonData, r As Range)
    r = mcd.c
End Sub
