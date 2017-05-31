VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickupHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' obydwie kolekcje sa kontenerami typu ITransit
'Private ppum0coll As Collection
'Private po400coll As Collection
' slownik
Private dic As Dictionary
Private t As ITransit
Private ii As IItem
Public gotopo400 As Boolean

Private HOW_MANY_DAYS_FOR_PPUS0 As Range

Private ce As CatchError



'Public Function getppum0coll() As Collection
'    Set getppum0coll = ppum0coll
'End Function
'
'Public Function getpo400coll() As Collection
'    Set getpo400coll = po400coll
'End Function
'
'
'
'
'




Public Function getDictionary() As Dictionary
    Set getDictionary = dic
End Function

Public Sub migrateDataIntoCollections(ByRef nyr As Collection, ByRef ic As Collection)
    If dic.COUNT > 0 Then
        For Each key_element In dic.Keys
            
            Set t = dic.item(key_element)
            
            
            If CDate(CStr(t.mDeliveryDate) & " " & CStr(t.mDeliveryTime)) < CDate(Now) Then
                nyr.Add t
            Else
                ' iteracja jest do dupy
                ' czasochlonna
                ' cos lepszego mi trzeba
                ' ----------------------------------------------------------
                'Dim ii As IItem
                'For Each ii In ic
                '    If CStr(ii.getDate()) = CStr(t.mDeliveryDate) Then
                '        ii.addTtransit t
                '        Exit For
                '    End If
                'Next ii
                ' ----------------------------------------------------------
                
                
                ' proba przyspieszenia algorytmu poprzez wykorzystanie
                ' zlotego srodka
                mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek t, ic, Int(ic.COUNT / 2)
            End If
        Next key_element
    End If
End Sub

Private Sub mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek(ByRef t As ITransit, ic As Collection, podzial As Integer)
    
    Set ii = ic.item(podzial)
    
    If CStr(ii.getDate()) = CStr(t.mDeliveryDate) Then
        ii.addTtransit t
        ' Exit Sub
        'MsgBox ii.getInTransits.COUNT
        'MsgBox ic.Item(podzial).getInTransits.COUNT
    Else
    
        ' to rownanie miazdzy super szybko przeszukiwanie
        ' super fajnie ze jeden dzien to 1 liczba calkowita
        ' dzieki temu szybko mozemy sie przeniesc
        diff = CLng(ii.getDate) - CLng(t.mDeliveryDate)
        If CLng(ii.getDate) > CLng(t.mDeliveryDate) Then
            If podzial - diff > 0 Then
                mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek t, ic, Int(podzial - diff)
            'Else
            '    mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek t, ic, Int(Application.WorksheetFunction.RoundDown(CDbl(podzial / 2), 0))
            End If
        Else
            If podzial - diff <= ic.COUNT Then
                mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek t, ic, Int(podzial - diff)
            'Else
            '    mozeTrocheSzybciejPrzeszukamyIcKolekcjeMozeJakisZlotySrodek t, ic, Int(podzial + Application.WorksheetFunction.RoundUp(CDbl(podzial / 2#), 0))
            End If
        End If
    
    End If
    
End Sub

Public Sub takeDataFromPOP(m As MGO, mcd As CommonData)

    ' pierwszy asn
    For x = 0 To 1
        If CStr(m.pMS9POP00.transQTY(Int(x))) <> "" Then
            If CDate(m.convertToDateFromMS9POP00Date(CStr(m.pMS9POP00.transSDATE(Int(x))))) <= CDate(Format(mcd.pusLimit, "yyyy-mm-dd")) Then
            
                ' format pod czas
                txt_time = Left(m.pMS9POP00.transETA(Int(x)), 2) & ":" & Right(m.pMS9POP00.transETA(Int(x)), 2)
                If txt_time = "__:__" Then
                    txt_time = "00:00"
                ' tutaj fancy warunek gdyby ktos upadl na lep
                ' i dostal udaru zapisujac godzine 00:00 jako 24:00
                ElseIf txt_time = "24:00" Then
                    txt_time = "23:59"
                End If
                
                Set t = Nothing
                Set t = New Transit
                
                ' those 2 resume next is for strange behavior in mgo
                ' it's possible to have empty text in date & time
                On Error Resume Next
                t.mDeliveryDate = CDate(m.convertToDateFromMS9POP00Date(m.pMS9POP00.transEDA(Int(x))))
                On Error Resume Next
                t.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
                t.mNotYetReceived = True
                
                ' you can also say that name is sid
                t.mName = m.pMS9POP00.transBL(Int(x))
                t.mTRLR = m.pMS9POP00.transCONT(Int(x))
                t.mst = m.pMS9POP00.transST(Int(x))
                
                On Error Resume Next
                t.mPickupDate = CDate(m.convertToDateFromMS9POP00Date(CStr(m.pMS9POP00.transSDATE(Int(x)))))
                ' tutaj nie jest zdefiniowane
                ' t.mPickupTime = CDate(Format(Now, "hh:mm"))
                t.mQty = CLng(m.pMS9POP00.transQTY(Int(x)))
                t.mRoute = m.pMS9POP00.transROUTE(Int(x))
                t.mOnPo400 = False
                
                
                If Not dic.Exists(CStr(t.mName)) Then dic.Add CStr(t.mName), t
            End If
        End If
    Next x
    
End Sub

Public Sub takeDataFromPO400(m As MGO, mcd As CommonData)
    ' sprawdz ladnie prosze czy przypadkiem nie przeszlismy na ekran ms3p9800
    
    If ce.noNSRissue(m) Then
    
        Do
            ' mgo refresh
            m.refresh
            For x = 0 To 7
            
                ' czy pusta linia to jest
                If CStr(m.pMS9PO400.EDA(Int(x))) <> "______" Then
                
                    If CDate(m.convertToDateFromMS9PO400Date(m.pMS9PO400.SDATE(Int(x)))) <= CDate(Format(mcd.pusLimit, "yyyy-mm-dd")) Then
                
                    
                        ' format pod czas
                        txt_time = Left(m.pMS9PO400.ETA(Int(x)), 2) & ":" & Right(m.pMS9PO400.ETA(Int(x)), 2)
                        If txt_time = "__:__" Then
                            txt_time = "00:00"
                        ' tutaj fancy warunek gdyby ktos upadl na lep
                        ' i dostal udaru zapisujac godzine 00:00 jako 24:00
                        ElseIf txt_time = "24:00" Then
                            txt_time = "23:59"
                        End If
                        
                        Set t = Nothing
                        Set t = New Transit
                        
                        On Error Resume Next
                        t.mDeliveryDate = CDate(m.convertToDateFromMS9PO400Date(m.pMS9PO400.EDA(Int(x))))
                        
                        On Error Resume Next
                        t.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
                        t.mNotYetReceived = True
                        
                        ' you can also say that name is sid
                        t.mName = m.pMS9PO400.SID(Int(x))
                        t.mTRLR = m.pMS9PO400.TRLR(Int(x))
                        t.mst = m.pMS9PO400.st(Int(x))
                        
                        On Error Resume Next
                        t.mPickupDate = CDate(m.convertToDateFromMS9PO400Date(CStr(m.pMS9PO400.SDATE(Int(x)))))
                        ' tutaj nie jest zdefiniowane
                        ' t.mPickupTime = CDate(Format(Now, "hh:mm"))
                        t.mQty = CLng(m.pMS9PO400.qty(Int(x)))
                        t.mRoute = m.pMS9PO400.route(Int(x))
                        t.mOnPo400 = True
                        
                        
                        If Not dic.Exists(CStr(t.mName)) Then dic.Add CStr(t.mName), t
                        
                        
                        
                    End If
                
                End If
                
            Next x
        
            If m.getCode("ms9po400") = "I5487" Or m.getCode("ms9po400") = "I6155" Or m.getCode("ms9po400") = "I6293" Or m.getCode("ms9po400") = "I4528" Then
                Exit Do
            End If
        
            m.sendKeys "<pf8>"
        Loop While True
    Else
        mcd.errorLog = CStr(ce.getFeedbackInfo())
    End If
End Sub


Public Sub takeDataFromPPUM0(m As MGO, mcd As CommonData)

    
    Do
        m.refresh
        For x = 0 To 19 - 7
        
            If Trim(m.pZK7PPUM0.plt(Int(x))) <> "" Then
                If CDate(m.convertToDateFromZK7PPUM0Date(m.pZK7PPUM0.PU_DATE(Int(x)))) <= CDate(Format(mcd.pusLimit, "yyyy-mm-dd")) Then
            
                    Set t = Nothing
                    Set t = New Transit
                    
                    With m.pZK7PPUM0
                    
                        On Error Resume Next
                        t.mDeliveryDate = CDate(m.convertToDateFromZK7PPUM0Date(.DEL_DATE(Int(x))))
                        
                        On Error Resume Next
                        t.mDeliveryTime = CDate(Mid(.DEL_WINDW(Int(x)), 6, 2) & ":" & Right(.DEL_WINDW(Int(x)), 2))
                        t.mOnPo400 = False
                        t.mName = CStr(.PUS_NBR(Int(x)))
                        t.mNotYetReceived = True
                        t.mQty = CLng(.PART_QTY(Int(x)))
                        
                        On Error Resume Next
                        t.mPickupDate = CDate(m.convertToDateFromZK7PPUM0Date(.PU_DATE(Int(x))))
                        t.mRoute = CStr(.route(Int(x)))
                        
                        If Not dic.Exists(CStr(t.mName)) Then
                            If CDate(t.mDeliveryDate) >= CDate(Date) Then
                                dic.Add CStr(t.mName), t
                            End If
                        End If
                        Set t = Nothing
                    End With
                End If
            End If
        Next x
        
        If (m.getCode("zk7ppum0") = CStr("R6086")) Or (m.getCode("zk7ppum0") = CStr("R6123")) Then
            Exit Do
        End If
        
        m.sendKeys "<pf8>"
    Loop While True
End Sub


Public Sub takeDataFromPH100(m As MGO, mcd As CommonData)

    m.sendKeys "<Clear>"
    m.sendKeys "ms9ph100 <Enter>"
    With m.pMS9PH100
        .plt = mcd.plt
        .pn = mcd.pn
        .setDTYP "RECV", "ASN"
        .DATE_FROM = m.convertToMgoShortDate(CDate(Now - CLng(HOW_MANY_DAYS_FOR_PPUS0) - CLng(EXTRA_DAYS_FOR_HISTORY)))
        m.submit
        
        
        ' teraz bardzo wazna rzecz
        ' zeby to dzialalo poprawnie nalezy sprawdzic czy w ogole czesc jest znaleziona przez ekran historii
        
        ' check if there are no  issues like : I4017: PART NOT FOUND
        
        Dim ce As CatchError
        Set ce = New CatchError
        
        tylko_na_poczatku_jest_to_zerem_zwiazane_z_ukladem_ekranu_historii = 0
        If ce.catchIfPartNotFoundInHistoryScreen(.infocode) Then
            Do
                For x = (0 + tylko_na_poczatku_jest_to_zerem_zwiazane_z_ukladem_ekranu_historii) To 9
                    
                    ' kolejne wiersze z ph100 - ekran historii
                    If dic.COUNT > 0 Then
                        
                        
                        If dic.Exists(CStr(.SID(Int(x)))) Then
                        
                            Set t = dic.item(.SID(Int(x)))
                            
                            ' 3 warunki sprawdzajace czy recv jest na konkretny sid wraz z odowiednia wartoscia
                            ' TERAZ SUPER WAZNE - ZACHOWANIE LOGIKI DLA RECV JAK I DLA ASN JEST IDENTYCZNE WIEC
                            ' JESLI NIE ZAMIERZAM W POZNIEJSZYM CZASIE ZMIENIAC CZEGO KOLWIEK
                            ' ALBO JESLI CZEGOS REWOLUCYJNEGO SIE NIE DOWIEM - BEDZIE MOZNA SCIAGANC JEDEN IF
                            ' ROZROZNIAJACY ASN czy RECV
                            
                            
                            ' pierwszy warunek jest dla wszystkich bo w ogole sprawdzam co mam zrobic z dopasowanym pusem w nazwie
                            If Trim(CStr(.SID(Int(x)))) <> "" Then
                                If CStr(t.mName) = CStr(.SID(Int(x))) Then
                                
                                    ' sprawdzamy czy mamy jakis recving na tym pusie
                                    ' i jesli mamy to przechodzimy i patrzymy jaka jest wartosc recv
                                    If .DTYP(Int(x)) = "RECV" Then
                                    
                                        ' druga wersja to sprawdzenie czy recv jest na zero - to trzeba jako tako zignorowac
                                        ' w poprzedniej podwersji byla to druga opcja
                                        ' ale w sumie chodzi o to zeby sprawdzic to w pierwszej kolejnosci
                                        ' zeby potem druga opcja zwiazana z porownywaniem nie zmienialal nam flagi not yet recv
                                        If CLng(.qty(Int(x))) = 0 Then
                                            ' NOP
                                    
                                        ' jesli wartosc recv jest taka sama jak pusa to znaczy ze zostal on pieknie przyjety
                                        ElseIf CLng(.qty(Int(x))) = CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                            
                                        
                                        
                                        
                                        ' trzecia opcja to w ogole przyjecie pusa na wartosc inna niz wartosc pusa
                                        ElseIf CLng(.qty(Int(x))) <> CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                        
                                        End If
                                    ElseIf .DTYP(Int(x)) = "ASN" Then
                                        
                                        ' to jest zsynchronizowane z recv na zero
                                        ' tyle ze samo recv na zero to mozna ignorowac
                                        ' natomiast asn dr jest na potrzeby dopisania gdy transport nie przychodzi do fabyrki
                                        ' wartoscia zgodna z tym co bylo na pickup sheet
                                        If CLng(.qty(Int(x))) = CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                         ' druga wersja to sprawdzenie czy recv jest na zero - to trzeba jako tako zignorowac
                                        ElseIf CLng(.qty(Int(x))) = 0 Then
                                            ' NOP
                                        
                                        ' trzecia opcja to w ogole przyjecie pusa na wartosc inna niz wartosc pusa
                                        ElseIf CLng(.qty(Int(x))) <> CLng(t.mQty) Then
                                            t.mNotYetReceived = False
                                            t.mQty = 0
                                        
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Exit Sub
                    End If
                Next x
                
                If m.getCode("ms9ph100") = CStr("R6086") Then
                    Exit Do
                End If
                
                m.sendKeys "<pf8>"
                tylko_na_poczatku_jest_to_zerem_zwiazane_z_ukladem_ekranu_historii = 1
            Loop While True
        Else
            mcd.errorLog = CStr(ce.getFeedbackInfo())
        End If
    End With
End Sub

Private Sub Class_Initialize()
    Set HOW_MANY_DAYS_FOR_PPUS0 = ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0")
    Set ppum0coll = New Collection
    Set po400coll = New Collection
    Set dic = New Dictionary
    gotopo400 = False
    
    Set ce = New CatchError
End Sub

Private Sub Class_Terminate()
    Set ppum0coll = Nothing
    Set po400 = Nothing
    Set dic = Nothing
    Set ce = Nothing
End Sub
