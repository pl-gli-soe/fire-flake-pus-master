VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyListReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' obiekt uruchomiony z tej klasy jest zagniezdzony jako komponent obeiktu klasy FireFlakeLight
' czyli mamy tutaj do czynienia z druga warstwa pracy makra fire flake
' ten oto obiekt w ktorym teraz sie znajdujemy to DailyReport lapiacy metody po interfejsie IReport
' jego glownym komponentami jest kolekcja iteracji oraz uchwyt pojdedynczej iteracji aby latwiej pracowalo sie w petlach
' for each zeby obiekt od razu byl zdefiniowany
' wazna sprawa jest rowniez to aby pamietac ze kolekcja iteracji ma kolejne swoje elementy jako IIteration dzieki temy interfejsowi
' mozna chwytac dane z daily, hourly i weekly bez najmniejszego bolu.


Implements IReport

Private iter As IIteration
Private itersCollection As Collection
Private box_for_itersCollection As FireFlakePMCollection
Private ffld As IFireFlakeLayout
Private dc As IDynamicColors


Private mlt As LAYOUT_TYPE
Private mst As START_TYPE

Private ce As CatchError


Private Sub Class_Initialize()

    Set mgo_handler = Nothing
    Set iter = New DailyIteration
    Set itersCollection = New Collection
    Set dc = New DailyDynamicColors
    
    Set box_for_itersCollection = New FireFlakePMCollection
    
    Set ce = New CatchError
End Sub

Private Sub Class_Terminate()
    Set mgo_handler = Nothing
    Set iter = Nothing
    Set itersCollection = Nothing
    Set box_for_itersCollection = Nothing
    Set ffld = Nothing
    Set dc = Nothing
    
    Set ce = Nothing
End Sub

Private Sub IReport_adjustLayoutAfterAllJobs()
    
    Dim rng As Range, sh As Worksheet

    If mst = FROM_THE_BEGINNING Then
        If mlt = LIST_LAYOUT Then
            
            Set sh = ThisWorkbook.ActiveSheet
            Set rng = sh.Range("a3").End(xlToRight)
            Set rng = sh.Range(sh.Cells(4, 2), sh.Cells(4, rng.Column - 1))
            rng.EntireColumn.AutoFit
        End If
    End If
End Sub

Private Sub IReport_clearAndStop()

End Sub

Private Sub IReport_fakeDataChangeToRecalcFirstRunoutFormula()
    Dim rng As Range, ir As Range
    
    If mlt = LIST_LAYOUT And (mst = FROM_THE_BEGINNING Or mst = CONTINUE_BROKEN_ONE) Then
        
        dc.assignDynamicColorsrange
        Set ir = dc.getDataScope().item(1, 1)
        Do
            
            tmp = ir
            ir = -1
            DoEvents
            ir = tmp
            
            Set ir = ir.Offset(1, 0)
        Loop Until Trim(ir.Parent.Cells(ir.Row, 2)) = ""
    End If
End Sub

Private Sub IReport_fullyRefresh()

    dc.assignDynamicColorsrange
    dc.recalcColors
End Sub

Private Function IReport_getCollectionOfIterations() As Collection
    Set IReport_getCollectionOfIterations = Nothing
End Function

Private Sub IReport_partialyRefresh()

    dc.assignDynamicColorsrange
    dc.recalcColors
End Sub

Private Sub IReport_pause()

End Sub


' metoda connect with input list bardziej by nazwa pasowala
' jest to piewrsza metoda zaraz po uruchomieniu pre konstrukora
' mozna powiedziec ze to jest ten wlasciwy konstruktor
' mysle ze warto z mojej strony pomyslec nad pewna konwencja
' zapisu i kazdy projekt jaki zaczynam zaczac od interfejsu w stylu IObject
' ktora wymuszac bedzie uruchomieniu swoistego dynamicznego konstrktora jednak z drugiej strony
' jak stowrzyc interfejs w vba ktory ma dynamiczna ilosc argumentow :D
Private Sub IReport_repInit(zakres As Range, limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, rqm_limit As Date, Optional mrowFlag As Long)

    If ce.inputDataAvail(zakres) Then

        ' nalezy podkreslic ze wszystkie te dane ktore sa do siebie przypisywane to w dalszym ciagu zonglowanie danych offline
        ' czy nie ma jeszcze wartosci dodanej
        ' jest to tylko przygotowanie kolekcji pod dane z mgo - zeby bylo spojnie
        ' plus dodatkowe pola przetrzymujace info jakie to wlasciwie raport ma byc
        mlt = lt
        mst = st
        
        ' sprawa zwiazana przede wszystkim z koniecznoscia rozroznienia kiedy mamy do czynienia
        ' z elementami raport new czy continue?
        Set ffld = New FireFlakeLayoutPM
        ffld.connectWithIterationBox box_for_itersCollection
        
        ' pod lt i st kryje sie info od init args
        ffld.init lt, st
        
    
        Dim r As Range
        Set itersCollection = Nothing
        Set itersCollection = New Collection
        
        ' z racji tego ze jak filtruje to nie patrze na wszystkie musi byc niezalezna flaga sprawdzajaco ktory
        ' pn z kolej jest brany pod uwage
        ' zaczynamy od pierwszego
        ' rowFlag = 1 - tylko w przypadku gdy pracuje od poczatku
        ' w przypadku continue broken report moze byc odrobine inaczej
        
        If CLng(mrowFlag) > 0 Then
            ' ta tutaj logika przesuwa mimo wszystko sie na jeden wiersz w dol
            ' okazuje sie ze parametr z klasy cebeer nie do konca jest w linii ze swoja nazwa
            ' w rzeczywistosci pokazuje ostatni element juz wpisanych danych
            rowFlag = mrowFlag - 3
        Else
            rowFlag = zakres.item(1, 1).Row - 1
        End If
        
        For Each r In zakres
        
            ' ten prosty warunek rozwiazuje problem
            ' schowanych wierszy przez filtr
            
            If Not r.EntireRow.Hidden Then
        
                Set iter = Nothing
                Set iter = New DailyIteration
                iter.connectFireFlakeLayout ffld
                
                
                ' do dnia ostatniego pazdziernika 2014 rozwiazanie krzyzowalo sie pomiedzy komponentami obiektowki co nieco myli podczas realizacji
                ' i wracania poraz kolejny do implementacji i nie wiedziec co sie dzieje do okola
                
                ' zatem co by nie meczyc tematu za bardzo nalezy dodac suba w implementacji iterfejsu co robi czesciowe uzupelnianie danych
                ' w obiekcie typu IIteration
                ' a jest nim:
                ' sub powstaly dnia 31 oct 2014: czesciowyFillNaTejIteracji jako komponenet subowy obiekty typu iiteracja
                ' ======================================
                iter.czesciowyFillNaTejIteracji r, lt, st, limit, rqm_limit, rowFlag
                ' inkrementacja zawsze po wykonaniu czesciowego wypelnienia
                rowFlag = rowFlag + 1
                
                ' dodajemy szczatkowo iteracje tylko wejsciowymi danymi
                ' praca dalsza na juz istniejacych obiektach - z kwesita po kolei dodawania kolejno rqm, pus/asn / balance
                itersCollection.Add iter
                
                Set iter = Nothing
            End If
        Next r
    End If
    

End Sub



' metoda najwazniejsza jesli chodzi o ffa
' czy run calego raportu
' to jest glowna metoda uruchmiana pod daily i uklad listowy
' uklad list jest dobry poniewaz mozna wtedy bezproblemowo ustawiac filtrowanie na danych
Private Sub IReport_repRun(imgo As MGO, sh As StatusHandler)


    ' no i teraz lecimy dalej
    ' chodzi teraz o wypelnienie reszty reszta
    ' mamy niepelny fill (wlasciwie to wypelniony nawet czesciowo tylko commondata co by dalej pojsc w ogole)
    ' na kolekcjach ale jest gotowe rusztowanie - czeka puste kolejne itemy rqm intransit i ebal.
    ' tj. ilosc itemow oraz przygotowane puste kontenery na dane z mgo.
    

    
    ' jednak warunkiem podstawowym, ktory w ogole pozwala na uruchomienie raportu jest fakt ze w naszej kolekcji w ogole wystepuje jakis item
    ' dzieki temu ze wykonalismy rusztowanie w repInit
    ' w repRun nie musimy poslugiwac sie juz tak duza iloscia
    If itersCollection.COUNT > 0 Then
    
    
        ' troche zamieszenia wprowadze tym status handlerem
        ' ale nie chce znowu wrzucac tutaj tego jako pola niech pracuje ino tutaj
        sh.init_statusbar CLng(itersCollection.COUNT)
    
        If MGO_active(imgo) Then
        
        
            ' to je ino handler
            ' nie trzeba robic zadnego new DailyIteration
            ' poniewaz koniec koncow itersCollection zawiera juz szczatkowe info w sobie
            ' teraz juz trzeba tylko uzupelnic brakujace dane w juz istniejacych obiektach o odpowiednich typach.
            ' ograniczamy poza tym argumentacje w subie z 5 do jednego :)
            Dim mitem As IIteration
        
        
        
            ' i teraz male przesuniecie tej metody na sama gore
            ' aby latwiej znalezc pola ktore iteresuja pierwsza formule ebal ffa
            ffld.buildLayoutForCommonData
            
            iledni = 0
            ' ' Debug.Print "0. start " & CStr(Now)
        
            For Each mitem In itersCollection
            
                ' mgo active wersja 24 poprawiona czyli teraz bede sprawdzal czy mgo jest aktywne na kazdej iteracji
                If MGO_active(imgo) Then

                    mitem.createFullIteration imgo

                    sh.progress_increase

                Else
                    End
                End If
                
            Next mitem
            
            ' poczekamy az sie wszystko wypelni
            ' i potem znow petla layoutowa tym razem
            box_for_itersCollection.c = itersCollection
            box_for_itersCollection.calc_collection
            
            For Each mitem In itersCollection
                
                mitem.addColorFlavour
            Next mitem
            
            ffld.buildLayoutForDatesLabels itersCollection
            ffld.buildLayoutTheThinLines itersCollection ' ten argument to wlasciwie tak just in case
            ffld.endingBalanceBackgroundColor
            ffld.closingDayLines
            ffld.finalAdjustOnColumnsWidth
            ffld.addFilterOnListLayout
            ffld.freezePanesAtTheEnd mlt
            
            ffld.shrinkColumnsWidthForRqmItEbal
            '' Debug.Print "4. end before assigning dynamic colors " & CStr(Now)
            
            dc.assignDynamicColorsrange
            dc.recalcColors
            
            '' Debug.Print "5. after final recalc " & CStr(Now)
            
            ' ffld.autoFitContent
        End If
    End If

End Sub

Private Sub IReport_setCollectionOfIterations(coi As Collection)
    Set itersCollection = coi
End Sub
