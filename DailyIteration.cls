VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyIteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IIteration



' dzieki mcd iec jest zbedne poniewaz zawiera juz w srodku kluczowa informacyje
' Private iec As ITERATION_CONFIG
' mcd.iter_config !!!
Private mcd As CommonData


' kolekcja itemow- juz zunifikowanych - czyli po wyrzuceniu duplikatow
Private ic As Collection
Private ii As IItem

' handler pod prace z ekranami tranzytow
Private t As ITransit
' additional variable for synchro on between ms9po400 and zk7ppum0
Private ttmp As ITransit
' i tak jest potrzebna dodatkowa kolekcja tranzytow ktore nie zostaly bookniete a maja eda juz przeszla
Private notYetReceivedWithPastEdaCollection As Collection

Private push As PickupHandler

' layout pojedynczo dla kazdej iteracji aby mozna bylo robic continue z broken report
Private fireFlake_Layout As IFireFlakeLayout


Private imgo As MGO

Private HOW_MANY_DAYS_FOR_PPUS0 As Range


Private drh As DownloadingRulesHandler


Private ce As CatchError

Private Sub Class_Initialize()

    Set ic = New Collection
    Set mcd = New CommonData
    Set ii = New DailyItem
    ' Set fireFlake_Layout = New FireFlakeLayoutDaily
    Set notYetReceivedWithPastEdaCollection = New Collection
    
    Set HOW_MANY_DAYS_FOR_PPUS0 = ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0")
    
    
    Set push = New PickupHandler
    
    Set drh = New DownloadingRulesHandler
    
End Sub



Private Sub Class_Terminate()
    Set ic = Nothing
    Set mcd = Nothing
    Set ii = Nothing
    Set fireFlake_Layout = Nothing
    Set notYetReceivedWithPastEdaCollection = Nothing
    Set push = Nothing
    
    Set drh = Nothing
End Sub

Private Sub IIteration_addColorFlavour()

    ' tutaj bawimy sie tylko konkretnymi danymi
    fireFlake_Layout.buildIterationLayoutForCommonData ic, mcd, notYetReceivedWithPastEdaCollection
    fireFlake_Layout.buildIterationLayoutForCoverage ic, mcd
    ' fireFlake_Layout.autoFitContent

End Sub

Private Sub IIteration_connectFireFlakeLayout(ffll As IFireFlakeLayout)
    Set fireFlake_Layout = ffll
End Sub

Private Sub IIteration_createFullIteration(m As MGO)


    ' uruchomienie tej procedury znajduje sie w kazdej iteracji kolekcji kolekcji bo w sumie mozna to tak nazwac
    ' dlatego wydaje sie ze taka golizna tutaj wystepuje i ze nie trzeba zadnych dodatkowych arguemnt - just beautful :P


    ' ok zatem w tym przypadku musimy przejsc przez popa jako punkt pierwszy i uzupelnienie reszty reszty danych z common data
    ' lacznie z mgo narazie bedzie wygladac to tak ze glowna metoda znajduje sie nadrzednie choc faktycznie wcale nie jest tam potrzebna
    ' i tak rozwiazanie to usunelo z ff light refrencje dzieki czemu nie musismy przechodzic przez dwie hierarchie obiektow zagniezdzownych
    ' glowny new mgo znajduje sie w konstruktorze obiektu uchwytu fire flake light
    Set imgo = m
    imgo.mgoInit
    
    
    
    ' tutaj przez chwile bylo nalozenie sie wspolnych nazw metod :D
    ' ta metoda znaczyna juz prace z danymi z mgo
    ' sciaga wstepnie z popa to i owo
    fillRestOfCommonData imgo
    ' teraz zostaly tylko daty first runout i first warn
    ' aby je uzyskac musimy stworzyc dane pod coverage
    ' zatem musimy pobrac rqmsy oraz pusy
    ' metoda zawiera prace nad rqmsami jak i pusami
    If Trim(mcd.errorLog) = "" Then
        fillItemCollection imgo
    End If

End Sub

Private Sub fillItemCollection(m As MGO)


    ' kwestia komponentow
    ' przed addRqmsAndDatesIntoItems jestesmy jeszcze na ekranie ms9pop00
    ' i mozemy jeszcze zaciagnac dane rqms
    ' mozemy tez zrobic proste if else
    ' tj albo sciagamy dane z popa albo tylko z zk7pdrqm
    ' jesli user zarequestuje mniej niz 10 dni do ogladania nie ma potrzeby nawet zagladac na zk7pdrqm
    ' =============================================================================
    
    ' psuedo code
    ' =============================================================================
    ' if days < 10 or component
    '   ' here data downloaded from ms9pop00
    '   addRqmsAndDatesIntoItemsButFromPopScreen
    ' else
    '   ' here data downloaded from zk7pdrqm
    '   addRqmsAndDatesIntoItems
    ' =============================================================================

    ' this very first one is not only for the rqms but also provides dates
    'Public Enum ITERATION_CONFIG
    '    CONFIG_ASM
    '    CONFIG_POP
    '    CONFIG_M
    '    CONFIG_NULL
    'End Enum
    ' tutaj pieknie dopasowanie poniewaz addRqms... faktycznie jako drugi argument posiada typ ITERATION_CONFIG :D
    addRqmsAndDatesIntoItems m, mcd.iter_config
    
    ' teraz trzeba sprawdzic czy w ogole jest gdzie wsadzic pozostale dane
    ' to znaczy czy kolekcja itemow nie jest pusta tj. czy przynjamniej jeden rqm z data jest widoczny z perspektywy fire flake'a
    If ic.COUNT > 0 Then
        addPusesIntoItems m
    End If
    
    
    ' test_on_downloaded_data

End Sub

Private Sub test_on_downloaded_data()
    
    If ic.COUNT > 0 Then
        For Each ii In ic
            ' ' Debug.Print "date: " & CStr(ii.getDate) & " rqm: " & CStr(ii.getRqm) & " COUNT: " & CStr(ii.getInTransits.COUNT)
            
        Next ii
    End If
End Sub

Private Sub addRqmsAndDatesIntoItems(m As MGO, eic As ITERATION_CONFIG)


    ' narazie ten subroutine obsluguje tylko i wylacznie ekran daily rqm zk7pdrqm
    ' ale musis przecko byc tez naddatek zwiazany z plt comp
    ' i zatem to pomijamy w ogole
    
    If eic = CONFIG_ASM Then
        ' std assembly
        zk7pdrqm_logic m
    ElseIf eic = CONFIG_POP Then
        ' wlasciwie dopasowana struktura pod dane komponentowe
        ' bo nie ma zadnej lepszej alternatywy
        ms9pop00_logic m
    ElseIf eic = CONFIG_M Then
        ' dane manualne z zewnetrznego pliku
        ' narazie jeszcze brak idei :D
    ElseIf eic = CONFIG_NULL Then
        ' puste
        ' for now - no operation implemented
    ElseIf eic = CONFIG_Z Then
        zeros_logic DEFAULT_ZERO_RQMS
    Else
    
        MsgBox "Ten text nie moze sie pojawic!"
        ' End
    End If


    
    

End Sub

Private Sub zeros_logic(ile As Integer)
    
    Dim x As Integer
    Dim tmpdate As Date
    tmpdate = CDate(Format(Now, "yyyy-mm-dd"))
    For x = 0 To ile - 1
    
    
        If tmpdate <= mcd.rqmLimit Then
        
            ' jest OK nawet llimitacja rqmLimit juz ladnie lapie
            ' ' Debug.Print "date: " & currDate & ", RQM in TOTAL: " & CStr(m.pZK7PDRQM.PLT_TOTAL(Int(x), Int(PLT_TOTAL_ROW)))
            Set ii = Nothing
            Set ii = New DailyItem
            ii.setDate CDate(tmpdate + x)
            ii.setRqm CLng(0)
            
            ic.Add ii
        Else
            Exit For
        End If
        
    Next x
End Sub

Private Sub zk7pdrqm_logic(m As MGO)
    
    m.sendKeys "<Clear>"
    m.sendKeys "zk7pdrqm <Enter>"
    m.pZK7PDRQM.KANBAN = ""
    m.pZK7PDRQM.pn = mcd.pn
    
    m.pZK7PDRQM.plt = mcd.plt
    m.submit
    
    
    ' sprawdz teraz czy dany part number wywolal jakos akcje ze strony mgo - czy pojawily sie jakies dane
    Dim ce As CatchError
    Set ce = New CatchError
    
    
    ' czy dane sa dostepne
    If ce.dataAvail("zk7pdrqm", m.getCode("zk7pdrqm")) Then
        ' jesli dane znajduja sie w miejscu to wchodze w warunek i rozpoczynam prace nad danymi
        ' zatem zaczynamy od nowa
        ' implementacja w ff 3.95 na rqmsach jest masakryczna prawie ze nieczytelna
        ' trzeba wrocic do dobrego zwyczaju pisania kodu poprawnie
        
        ' czy daty nie sa puste
        If Not ce.catchIfInDailyRqmsThereAreNoDates(m.pZK7PDRQM.dates(0)) Then
        
        
            ' w tym miejscu algorytm przeszedl dwa sprawdzenia czy napewno ma podjac pobranie rqmsow
            ' sytuacja bardzo czysta i raczej niespodzianek nie powinno byc
            ' kod jest bardzo sterylny i wymaga spelniania bardzo restrykcyjnych warunkow
            ' musi byc naprawde odowiedni kod 22 2 aby w ogole moglo dojsc do logiki

            
            Dim currDate As Date, prevDate As Date
            ' initialize
            ' problem w tym ze nie tak latwo nalozyc poczatek istnienia tych dwoch zmiennych poniewaz koniec koncow
            ' the prevDate ma pokazywac nie stan biezacy ale stan z poprzedniej iteracji
            currDate = CDate(m.convertToDateFromZK7PDRQMDate(m.pZK7PDRQM.dates(0)))
            prevDate = CDate(Format(Now, "yyyy-mm-dd")) - 1
            ' zrobimy dwa scenariusze poniewaz jesli jest potrzeba F8 to jednak wiecej trzeba klikac i nie chcemy tego robic jesli faktycznie nie trzeba
            ' SCENARIO NO 1 - gdy nie ma tak duzo mY i wszystko miesci sie na raz bez potrzeby klikania F8
            '
            ' dodatkowo trzeba wziazc pod uwage ze obywa warunki tj. obydwa scenariusze sa do siebie ludzaco podobne i zakrawa i lamanie zasady DRY
            ' co nie jest zbyt si, jednak z drugiej strony wazy na jasnosci kodu i jego przejrzystosci dla przyszlosci gdy bede chcial optymalizowac go
            ' w inny mozliwy sposob
            ' mozliwe ze mozna by nadac funkcje komponentowe tutaj jeszcze
            ' ===============================================================================================================
            If m.getCode("zk7pdrqm") = "R6101" Then
            
            
                
            
                ' jestesmy w srodku scenariusza pierwszego tj. za kazdy razem PLT TOTAL bedzie w tym samym miejscu zatem juz teraz jeszcze przed iteracja
                ' uruchamiamy sprawdzenie na ktorym oto wierszu sie znajduje
                PLT_TOTAL_ROW = m.pZK7PDRQM.findPLTTOTALrow()
            
                Do
                    
                    If petlaForX04(m, prevDate, currDate, Int(PLT_TOTAL_ROW)) Then
                        Exit Do
                    End If
                    
                    
                    ' jesli inquiry complete nie trza przechodzic dalej
                    If m.getCode("zk7pdrqm") = "R6086" Then
                        Exit Do
                    End If
                    
                    
                    m.sendKeys "<pf11>"
                
                
                ' kod 6086 to inquiry complete nawet jesli mamy tylko jedna strone to przeleci przynajmniej raz iteracja
                Loop Until m.getCode("zk7pdrqm") = "R6120"
            
            ' SCENARIO NO 2 - czyli i F8 i F11
            ' -----------------------------------------------------------------------------------------------------------------------------------
            ElseIf m.getCode("zk7pdrqm") = "R6693" Then
            
                ' scenariusz drugi jest bardziej bogaty zwiazany jest z duza iloscia MY przez co dodatkowo w ruch wprawiany jest F8
                ' 1. ile razy F11
                ilerazyf11 = 0
                ' 2. ile razy F8 -> a raczej F8 do oporu
                Do
                    
                    kliknijF11 m, Int(ilerazyf11)
                    ' klikaj F8 do oporu
                    zjedzNaDolZaPomocaF8 m
                    
                    
                    
                    PLT_TOTAL_ROW = m.pZK7PDRQM.findPLTTOTALrow()
                    
                    If petlaForX04(m, prevDate, currDate, Int(PLT_TOTAL_ROW)) Then
                        Exit Do
                    End If
                    
                    
                    ' jesli inquiry complete nie trza przechodzic dalej
                    If (m.getCode("zk7pdrqm") = "R6086") Or (Trim(m.getCode("zk7pdrqm")) = "R6120") Then
                        Exit Do
                    End If
                
                    m.sendKeys "<Enter>"
                    ilerazyf11 = ilerazyf11 + 1
                    
                Loop Until Trim(m.getCode("zk7pdrqm")) = "R6120"
            End If
            ' ===============================================================================================================
        Else
            mcd.errorLog = CStr(ce.getFeedbackInfo())
        End If
    Else
        ' plan jest taki ze jesli danych nie - ma trzeba je jakos zastapic - mozna np zerami po prostu zgodnie z iloscia danych
        ' jaka requestowal user
        ' aktualnie implementacja jest pusta to znaczy ze musze dac sam sobie na pozniej dac znac ze to jest do uzupelnienia :D
        ' ---------------------------------------------------------------------------------------------------
        mcd.errorLog = CStr(ce.getFeedbackInfo())
        
        
        ' PLEASE INSERT CODE HERE FOR CUSTOM DAILY RQMS - NO DATA IN MGO - OBSOLETE FOR NOW - this place according to above code should never start.
        ' ---------------------------------------------------------------------------------------------------
        
            ' koniec koncow rezygnuje z tej koncpecji, poniewaz kod bedzie zbyt gleboko zagniezdzony
            ' ostatecznie odseparowalem logike w poziomie wyzej dzieki temu kod stanie sie nieco bardziej klarowny
        
        ' ---------------------------------------------------------------------------------------------------
    End If
End Sub

Private Function petlaForX04(m As MGO, prevDate As Date, currDate As Date, PLT_TOTAL_ROW As Integer) As Boolean

    ' blokuje to co nie pasuje
    ' Debug.Assert mcd.plt <> "RU"
    
    petlaForX04 = False
    Dim ce As CatchError
    Set ce = New CatchError
    
    For x = 0 To 4
    
    
        If Not ce.catchIfInDailyRqmsThereAreNoDates(m.pZK7PDRQM.dates(Int(x))) Then
            currDate = CDate(m.convertToDateFromZK7PDRQMDate(m.pZK7PDRQM.dates(Int(x))))
            
            
            
            ' when we have downdays we need to add fake data (e.g. Saturday and Sunday)
            optionalFillWithZeroRqmsOnDownDays prevDate, currDate
            
            ' srawdzamy czy miescimy sie w limicie ilosc requirementsow jakie chcemy widziec w raporcie
            If currDate <= mcd.rqmLimit Then
            
                ' jest OK nawet llimitacja rqmLimit juz ladnie lapie
                ' ' Debug.Print "date: " & currDate & ", RQM in TOTAL: " & CStr(m.pZK7PDRQM.PLT_TOTAL(Int(x), Int(PLT_TOTAL_ROW)))
                Set ii = Nothing
                Set ii = New DailyItem
                ii.setDate currDate
                ii.setRqm CLng(m.pZK7PDRQM.PLT_TOTAL(Int(x), Int(PLT_TOTAL_ROW)))
                
                ic.Add ii
            Else
                ' tutaj jest caly sens limit rqm zawarty - gdyby nie wczesniejsze wyjscie z petli wszystko zatrzymaloby sie na tym samym miejscu
                petlaForX04 = True
                Exit For
            End If
            
            prevDate = currDate
        Else
            mcd.errorLog = CStr(ce.getFeedbackInfo())
            Exit Function
        End If
    Next x
    
    
    
End Function

Private Sub optionalFillWithZeroRqmsOnDownDays(prevDate As Date, currDate As Date)


    ' pierwszy warunek sprawdzajcy czy nie jestesmy u zarrania dziejow iteracji
    ' wtedy i tak nie wystepuje prevDate wiec nie ma zaczynac akcji - warunek brzegowy
    
    ' warunek ktory sprawdza czy mamy jakis downdays
    ' -------> POTATO <-----------
    If CLng(currDate) - CLng(prevDate) > 1 Then
        
        For x = (CLng(prevDate) + 1) To (CLng(currDate) - 1)
        
        
            If CDate(x) <= mcd.rqmLimit Then
                Set ii = Nothing
                Set ii = New DailyItem
                ii.setDate CDate(x)
                ' there's no prod on dates without data on mgo
                ' so simply downday
                ii.setRqm CLng(0)
                
                ic.Add ii
            Else
                Exit Sub
            End If
        Next x
    End If

End Sub

Private Sub kliknijF11(m As MGO, qty As Integer)

    If qty > 0 Then
        For x = 1 To qty
            m.sendKeys "<pf11>"
        Next x
    End If
End Sub

Private Sub zjedzNaDolZaPomocaF8(m As MGO)

    While m.getCode("zk7pdrqm") = "R6693"
        m.sendKeys "<pf8>"
    Wend
End Sub



Private Sub ms9pop00_logic(m As MGO)

    ' assumption: we're still on ms9pop00 just double check it:
    Dim ce As CatchError
    Set ce = New CatchError
    
    If CStr(m.actualScreen) = "MS9POP00" Then
    
        ' slusznie jedynie tutaj kod moze sie uruchomic jesli wszystko poszlo tak jak nalezy
        
        Dim currDate As Date, prevDate As Date
        ' initialize
        ' problem w tym ze nie tak latwo nalozyc poczatek istnienia tych dwoch zmiennych poniewaz koniec koncow
        ' the prevDate ma pokazywac nie stan biezacy ale stan z poprzedniej iteracji
        prevDate = CDate(Format(Now, "yyyy-mm-dd")) - 1
        
        take10FromPOP m, m.pMS9POP00, currDate, prevDate
    
    Else
        ' test nie poszedl pomyslnie cos jest nie tak skoro bylem na ekranie ms9pop00 a teraz juz nie jestem :/
        ce.sth_went_wrong_with_mgo_on_ms9pop00_screen_download_rqm_attempt
        mcd.errorLog = CStr(ce.getFeedbackInfo())
        Exit Sub
    End If
End Sub

Private Sub take10FromPOP(ByRef m As MGO, ByRef pop As MS9POP00, ByRef cd As Date, ByRef pd As Date)

    Dim x As Integer
    Dim czy_klikneli_juz_f11 As Boolean
    czy_klikneli_juz_f11 = False
    With pop
        ' pop screen have always 10 rows for balance
        ' after 5th need to click F10
        For x = 1 To 5
            cd = CDate(m.convertToDateFromMS9POP00ShortDate(m.pMS9POP00.balance_on_pop_data(Int(x), 1)))
            optionalFillWithZeroRqmsOnDownDays pd, cd
            
            
            ' srawdzamy czy miescimy sie w limicie ilosc requirementsow jakie chcemy widziec w raporcie
            If cd <= mcd.rqmLimit Then
            
                ' jest OK nawet llimitacja rqmLimit juz ladnie lapie
                ' ' Debug.Print "date: " & currDate & ", RQM in TOTAL: " & CStr(m.pZK7PDRQM.PLT_TOTAL(Int(x), Int(PLT_TOTAL_ROW)))
                Set ii = Nothing
                Set ii = New DailyItem
                ii.setDate cd
                ii.setRqm CLng(pop.balance_on_pop_data(Int(x), 2))
                
                ic.Add ii
            Else
                Exit For
            End If
        
            ' after 5th need to click F11
            If Not czy_klikneli_juz_f11 Then
                If x = 5 Then
                    m.sendKeys "<pf11>"
                    czy_klikneli_juz_f11 = True
                    x = 1
                End If
            End If
            
            pd = cd
        Next x
    
    End With
End Sub

Private Sub addPusesIntoItems(m As MGO)



    ' metoda prywatna opierajaca swoja prace na dwoch strategiach
    ' jedna z nich opiera sie na zk7ppus0 & ms9ph100 & zk7prosl
    ' drugi - przypadek dla ktorego czesci nie posiadaja pusow - czyli tylko asny - czyli wystarczy jak przejrze ino ms9po400
    
    ' waznym elementem jest brak koniecznosci sprawdzenia czy mozna w ogole dane wsadzac - zrobilem to na poziomie metody wyzej
    ' jestem nieco niestabliny poniewaz raz robie tak a raz
    ' zakladam ze jednak chyba powinienen wszystko robic od srodka
    ' ale zalozmy ze jest to szczegolny przypadek dla ktorego nie bede specjalnie wchodzil w ta metode
    
    
    
    ' =====================================
    ' sprawa dodatkowa jest czasami dziwne zachowanie ekranu ms9po400
    ' tj. zamiast wchodzic na niego bezposrednio zaczyna przechodzic
    ' na jakis ekran warningowy, ktory uswiadamia ze limit asnow zostal przekroczony
    ' z tego co pamietam, albo raczej to co dla mnie pamieta task outlooka to:
    '
    ' MS3P9800
    ' ASNS EXCEED TABLE SIZE OF 900
    '
    ' =====================================
    
    
    ' powazny blad
    ' Set push = New PickupHandler
    
    
    m.sendKeys "<Clear>"
    m.sendKeys "zk7ppus0 <Enter>"
    
    m.pZK7PPUS0.plt = mcd.plt
    m.pZK7PPUS0.pn = mcd.pn
    ' wersja wstepna zaklada zmienna globalna HOW_MANY_DAYS_FOR_PPUS0 jako stala 20 w Global Module
    ' bedzie potem mozliwosci modyfikowania tej wartosci z formularza startowego :D
    ' powyzszy komentarz zostal juz chyba uniewazniony poniewaz form juz udostepnia wraz z sheetem register
    ' mozliwosc odpowiedniej malwersacji na danych
    ' m.pZK7PPUS0.DELIVERY_DATE_FROM = m.convertToMgoLongDate(CDate(Now - HOW_MANY_DAYS_FOR_PPUS0))
    m.pZK7PPUS0.DELIVERY_DATE_FROM = m.convertToMgoLongDate(CDate(Now - HOW_MANY_DAYS_FOR_PPUS0))
    m.submit
    m.sendKeys "<pf12>"
    
    
    
    ' bardzo proste sprawdzenie czy moge sciagnac pusy, czy raczej musze pracowac na asnach
    If CStr(m.actualScreen()) = "ZK7PPUM0" Then
    
        ' tutaj zabawa z pusami
        ' ======================================
        
        ' te 3 procedury sa schowane w fillTransitCollectionByPuses
        ' szczegolnie nalezy zwrocic uwage na ph100 - ekran historii
        ' jest to dosyc gleboko osadzona implementacja w logice
        ' push.takeDataFromPPUM0 m
        ' push.takeDataFromPH100 m, mcd
        ' push.fillCollections notYetReceivedWithPastEdaCollection, ic
        
        
        fillTransitCollectionByPuses m, push
        ' niestety okazalo sie ze nie moge byc az tak restrykcyjny jesli chodzi o zaciaganie danych tylko i wylacznie z PPUS0
        ' musze znow wejsc na ekran ms9po400 i sprawdzic przykladowo help shippy badz
        ' tak czy siak sprawdzic asny poniewaz czesc jest w trakcie zmiany dostawcy z lokalnego na osea
        ' i na odwrot
        
        ' cudowne uproszczenie implementacji dzieki zmiennej typu slownik
        fillTransitCollectionFromMs9po400 m, push, mcd
        
        ' ======================================
    Else
    
        
    
        ' tutaj sprawa duzo prostsza poniewaz sciagamy tylko i wylacznie asny
        ' bierzemy pod uwage tylko i wylaczenie dane z ms9po400
        ' ======================================
        
        ' cudowne uproszczenie implementacji dzieki zmiennej typu slownik
        push.gotopo400 = True
        fillTransitCollectionFromMs9po400 m, push, mcd
        
        ' ======================================
    End If
    
    migrateDataFromDictionaryToCollections push, notYetReceivedWithPastEdaCollection, ic
    
    
End Sub

Private Sub migrateDataFromDictionaryToCollections(ByRef push As PickupHandler, ByRef nyr As Collection, ByRef ic As Collection)
    push.migrateDataIntoCollections nyr, ic
End Sub


Private Sub fillTransitCollectionByPuses(ByRef m As MGO, ByRef push As PickupHandler)


    ' kwestia zaimplementowania osobnej klasy sprawujacej wladze nad ekranami ppum0 oraz ph100
    ' ktora na wyjsciu daje dwa obiekty - uzupelnia odpowiednio itransity dla itemow std
    ' oraz wypelnia not yet recevied collection
    
    
    push.takeDataFromPPUM0 m, mcd
    push.takeDataFromPH100 m, mcd
    ' push.fillCollections notYetReceivedWithPastEdaCollection, ic
    

End Sub

Private Sub fillTransitCollectionFromMs9po400(ByRef m As MGO, ByRef push As PickupHandler, ByRef mcd As CommonData)

    If CBool(push.gotopo400) Then
    
        m.sendKeys "<Clear>"
        m.sendKeys "ms9po400 <Enter>"
        m.pMS9PO400.plt = mcd.plt
        m.pMS9PO400.pn = mcd.pn
        m.pMS9PO400.KANBAN = ""
        m.submit
        
        push.takeDataFromPO400 m, mcd
    
    End If
    
End Sub


Private Sub fillRestOfCommonData(m As MGO)

    ' already in place filled during init thing
    '=====================================================
    ' PLT
    ' PN
    ' comment
    ' c
    ' iter_conf
    ' rowFlag
    ' laytout type
    ' start type - start from the beg or continue
    ' run type - daily for exmaple
    ' rqm limit
    ' pus limit
    '=====================================================
    
    m.sendKeys "<Clear>"
    m.sendKeys "ms9pop00 <Enter>"
    m.pMS9POP00.DS = "6"
    m.pMS9POP00.pn = mcd.pn
    m.pMS9POP00.plt = mcd.plt
    m.submit
        
    Dim ce As CatchError
    Set ce = New CatchError
    
    If ce.dataAvail("ms9pop00", m.getCode("ms9pop00")) Then
        ' silowe castowanie - tak jakby co :)
        mcd.bank = checkIsLong(m.pMS9POP00.bank)
        mcd.bbal = checkIsLong(m.pMS9POP00.bbal)
        mcd.cbal = checkIsLong(m.pMS9POP00.cbal)
        mcd.doh = checkIsDouble(m.pMS9POP00.firstDOH)
        mcd.duns = CStr(m.pMS9POP00.duns)
        mcd.misc = checkIsLong(m.pMS9POP00.misc)
        mcd.os = checkIsLong(m.pMS9POP00.os)
        mcd.pcsToGo = checkIsLong(m.pMS9POP00.PCS_TO_GO)
        mcd.qhd = checkIsLong(m.pMS9POP00.qhd)
        mcd.stdPack = checkIsLong(m.pMS9POP00.STD_PACK)
        mcd.ttime = checkIsLong(m.pMS9POP00.TT)
        
        ' czego braklo i trzeba dopisac
        mcd.supplierName = CStr(m.pMS9POP00.SUPPLIER)
        mcd.partName = CStr(m.pMS9POP00.PART_NAME)
        mcd.fupCode = CStr(m.pMS9POP00.F_U)
        mcd.fmaFupCode = CStr(m.pMS9POP00.a)
        mcd.mode = CStr(m.pMS9POP00.mode)
        mcd.count_cmnt = CStr(m.pMS9POP00.COUNT)
        mcd.o_cmnt = CStr(m.pMS9POP00.O)
        ' to jest jeszcze do poprawy
        mcd.f_cmnt = CStr(m.pMS9POP00.F)
        
        
        
        ' teraz przejdzmy do czesci z asnami / pusami
        ' i jesli jest mozliwosc to posciagaj dane od razu tutaj
        
        ' wyliczamy tak jak w przypadku std tablicy od zera
        If CStr(m.pMS9POP00.transQTY(2)) = "" Then
            push.gotopo400 = False
            
            ' tutaj znajduja sie przynajmniej 2 asny
            
            ' ten warunek sprawdza czy w ogole cos mamy
            If CStr(m.pMS9POP00.transQTY(0)) = "" Then
                ' tutaj nic nie ma wiec nawet nie uruchamiaj
                ' nop at the end
            Else
                push.takeDataFromPOP m, mcd
                
            End If
        ElseIf CStr(m.pMS9POP00.transQTY(2)) <> "" Then
            push.gotopo400 = True
        End If
    Else
        mcd.errorLog = CStr(ce.getFeedbackInfo())
    End If
    
    
    
    
    
    
    
    ' brakuje jeszcze tylko dat put first runout i first warning
End Sub

Private Function checkIsLong(s As String) As Long

    s = Replace(s, ".", ",")
    
    If Trim(s) = "" Then
        checkIsLong = CLng("0")
    Else
        checkIsLong = CLng(Trim(s))
    End If
        
End Function

Private Function checkIsDouble(s As String) As Double


    s = Replace(s, ".", ",")
    
    If Not IsNumeric(s) Then
        checkIsDouble = CDbl("0")
    ElseIf Trim(s) = "" Then
        checkIsDouble = CDbl("0")
    Else
        checkIsDouble = CDbl(Trim(s))
    End If
        
End Function


Private Sub IIteration_czesciowyFillNaTejIteracji(r As Range, lt As LAYOUT_TYPE, st As START_TYPE, limit As Date, rqm_limit As Date, rowFlag)

    ' plt & pn - dzieki tym dwom paramsom jestem w stanie odtworzyc caly coverage przy pomocy oczywiscie rowniez danych z mgo
    
    '=====================================================
    mcd.plt = Trim(r)
    mcd.pn = Trim(r.Offset(0, 1))
    '=====================================================
    
    ' dodatkowy komentarz wejsciowy co by moze latwiej sie filtrowalo na gotowym raporcie
    mcd.comment = Trim(r.Offset(0, 2))
    mcd.additionalCommnet = Trim(r.Offset(0, 3))
    
    'component -> flaga czy w inne ekrany mam zagldac
    ' OBSOLETE - najpierw spradzamy config
    ' potem sprawdzamy czy  ten konkretny plant to komponent
    ' mcd.c = doubleCheckOnPlant(r.Offset(0, 4), r)
    
    ' config enum - nie pamietam czy od zera czy od jedynki
    ' potem temat bede meczyl :D
    ' teraz wazna rzecz drh.getIUterConfig robi juz dwie pieczenie na jednym ruszcie
    ' wyrzuca do iter_config nowa wartosc
    ' plus modyfikuje argument dzieki temu ze jest z przyrostkiem byref
    '
    ' 2015-06-17
    ' kolejny temat wykorzystania drh w celu uzupelniania info o nieznanym plancie do tej pory z perspektywy fire flake
    ' trzeba w prosty sposob zapytac jakiego rodzaju plantem jest plant zaproponowany przez usera!
    
    
    
    
    ' OBSOLETE! <- zrozumialem co bylo nie tak w kodzie - szczegolnie z getIterConfig(r)
    ' ---------------------------------------------------------------------------------------------------------------------------
    ' ---------------------------------------------------------------------------------------------------------------------------
    ' z tym kodem to jest sroga masakra
    ' jest dwa razy pod rzad uzyty ta sama methoda
    ' poniewaz raz korzysta z jednej czesci a raz z drugiej w skrajnym przypadku, gdy plant nie zostal rozpoznany
    ' nie ukrywam ze dodatkowym problem jest to ze dynamicznie wywolywany jest form
    ' i moze temu - tak czy siak ten kod jest na tyle elastyczny ze podwojne jego wywolnie nie robi roznicy
    ' za kadym razem jest osiagniecie celu i gdy juz on jest drugi raz kod w metodzie nie jest wykonywany
    
    ' ta dwie linie musza sie powtarzac!
    ' mcd.iter_config = drh.getIterConfig(r)
    ' using here handikap vba
    
    ' do not delete this repeatition!
    
    '
    ' powyzszy komentarz trwal w wersji 05 poniewaz dwa razy uruchamialem kod get iter confing
    ' z powodu srogiego babola w srodku tego kodu
    ' ponizej juz widac ze uruchamiam to ostrojstwo ino raz
    ' ---------------------------------------------------------------------------------------------------------------------------
    ' ---------------------------------------------------------------------------------------------------------------------------
    
    ' implementacja zostala zmieniona i poprawiona!
    ' 2015-08-03 - ostatnia zmiana
    ' wykorzystanie w prawidlowy sposob wartosci eic
    ' czesniej funkcja zapominala zwrocic jaka kolwiek wartosc
    ' dlatego pojawial sie dziwny blad w pozniejszym kodzie!
    ' gdyby nie msgbox ktore mowil ze nigdy nie powinien sie pojawic ciezko by bylo mi znalezc ten blad
    mcd.iter_config = drh.getIterConfig(r)
    ' juz iter config ready zatem mozna jeszcze przypisac
    ' nalezy pamietac o tym ze ifComponent jest dosc mocno powiazany z pracy z poprzednia funkcja getIterConfig z racji tego
    ' wynik czy dany plant jest componentowy wynika z wartosci jaka przechwyci ogolna konfiguracja dane part numberu, czy tez
    ' bardziej akurat tutaj plantu, poniewaz w glownej mierze to plant decyduje czy mamy do czynienia z asm, czy moze z comp, czy moze jeszcze z czyms innym
    mcd.c = drh.ifComponent(mcd, r)
    
    
    ' flaga pod wiersz - szczegolnie przydatne dla odtworzenia danych przy popsustym raporcie
    ' plus 2 poniewaz raport daily w output zaczyna sie od 5 wiersza - nieco gorzej bedzie z hourly ale to juz inna historia
    mcd.rowFlag = CLng(rowFlag) + 4
    
    
    ' limitacje i konfiguracje
    mcd.layoutType = lt
    mcd.startType = st
    mcd.runType = DAILY
    mcd.pusLimit = CDate(limit)
    mcd.rqmLimit = CDate(rqm_limit)
End Sub




' metoda ta sprawdza czy napewno dany plant powinien zagladac w takie a nie inne screeny na mgo
' ta procedura bedzie teraz pracowac w klasie DownloadingRuleshandler
'Private Function doubleCheckOnPlant(c As Range, plt As Range)
'    ' wstepna uproszczona implementacja poniewaz nie chce mi sie narazie logiki pisac :D
'    doubleCheckOnPlant = Trim(c)
'End Function

Private Function IIteration_getCommonData() As CommonData
    Set IIteration_getCommonData = mcd
End Function

Private Function IIteration_getItem() As IItem
    Set IIteration_getItem = ii
End Function

Private Function IIteration_getItemsCollection() As Collection
    Set IIteration_getItemsCollection = ic
End Function

Private Function IIteration_getLayout() As Variant
    Set IIteration_getLayout = fireFlake_Layout
End Function

Private Sub IIteration_setCommonData(cd As CommonData)
    Set mcd = cd
End Sub

Private Sub IIteration_setItem(i As IItem)
    Set ii = i
End Sub

Private Sub IIteration_setItemsCollections(c As Collection)
    Set ic = c
End Sub
