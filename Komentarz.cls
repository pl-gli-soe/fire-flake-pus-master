VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Komentarz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public typeOfComment As COMMENT_TYPE


Public Sub assignTransitsIntoComment(r As Range, c As Collection)

    If c.COUNT > 0 Then
        typeOfComment = IN_TRANSIT
        
        r.ClearComments
        txt = CStr(prepareComment(c))
        ile = CLng(UBound(Split(txt, Chr(10))))
        ile = ile / 10
        If Trim(txt) <> "" Then
            r.AddComment CStr(txt)
            ' wartosci dobrane heurstycznie
            r.comment.Shape.Width = 100
            r.comment.Shape.Height = 120 * (ile)
        End If
        
        
    End If
End Sub

Public Sub assignOneTransitIntoComment(r As Range, it As ITransit)

    ' w pizdu na okolo - ale jesli nie bedzie narzutu pamieciowego i czasowego to tak to zostawie :P

    Dim c As Collection
    Set c = New Collection
    c.Add it
    
    assignTransitsIntoComment r, c
    Set c = Nothing
End Sub

Public Sub makeCommentBasedOnPopData(ByRef r As Range, ByRef mcd As CommonData)
    
    If typeOfComment = DATA_FROM_POP Then
    
    
        cmnt_txt = assignHerePopData(mcd)
        r.ClearComments
        r.AddComment CStr(cmnt_txt)
        r.comment.Shape.Width = 300
        r.comment.Shape.Height = 250
    End If
End Sub

Private Function assignHerePopData(ByRef mcd As CommonData) As String
    assignHerePopData = ""
    
    assignHerePopData = "MS9POP00: " & Chr(10) & _
        CStr(mcd.plt) & ", " & CStr(mcd.pn) & Chr(10) & _
        "DUNS: " & CStr(mcd.duns) & Chr(10) & Chr(10) & _
        "COUNT_CMNT: " & CStr(mcd.count_cmnt) & Chr(10) & _
        "O_CMNT: " & CStr(mcd.o_cmnt) & Chr(10) & Chr(10) & _
        "CBAL: " & CStr(mcd.cbal) & ", BBAL: " & CStr(mcd.bbal) & Chr(10) & _
        "FUC: " & CStr(mcd.fupCode) & ", FMA FUC: " & CStr(mcd.fmaFupCode) & Chr(10) & _
        "LOG: " & CStr(mcd.errorLog) & Chr(10) & Chr(10) & _
        "PUS Limit: " & CStr(mcd.pusLimit) & ", RQM Limit: " & CStr(mcd.rqmLimit) & Chr(10) & Chr(10) & _
        "MODE: " & CStr(mcd.mode) & ", BANK: " & CStr(mcd.bank) & Chr(10) & _
        "OS: " & CStr(mcd.os) & ", DOH: " & CStr(mcd.doh) & Chr(10) & _
        "PNAME: " & CStr(mcd.partName) & Chr(10) & _
        "SUPP NAME: " & CStr(mcd.supplierName) & Chr(10) & _
        "PCS TO GO: " & CStr(mcd.pcsToGo) & ", STD PACK: " & CStr(mcd.stdPack) & Chr(10) & _
        "QHD: " & CStr(mcd.qhd) & ", TT: " & CStr(mcd.ttime) & Chr(10) & Chr(10) & _
        "ST: " & CStr(mcd.startType) & ", RT: " & CStr(mcd.runType) & ", LT: " & CStr(mcd.layoutType) & Chr(10)
        
End Function


Private Function prepareComment(c As Collection) As String

    ' c to kolekcja kolejnych tranzytow
    txt = ""
    Dim t As ITransit
    For Each t In c
        If t.mNotYetReceived Then
            txt = txt & _
            "eda: " & CStr(t.mDeliveryDate) & Chr(10) & _
            "eta: " & CStr(t.mDeliveryTime) & Chr(10) & _
            "sdate: " & CStr(t.mPickupDate) & Chr(10) & _
            "ptime: " & CStr(t.mPickupTime) & Chr(10) & _
            "name: " & CStr(t.mName) & Chr(10) & _
            "qty: " & CStr(t.mQty) & Chr(10) & _
            "route: " & CStr(t.mRoute) & Chr(10) & _
            "status: " & CStr(t.mst) & Chr(10) & _
            "trlr: " & CStr(t.mTRLR) & Chr(10) & _
            "--------------------" & Chr(10)
        End If
    Next t
    
    prepareComment = txt

End Function
