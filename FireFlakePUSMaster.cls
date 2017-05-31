VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakePUSMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private rep As IReport
Private imgo As MGO


' CeBeEr - continue broken report
Private cebeer As ContinueBrokenReportHandler
Private init_sh As Worksheet
Private continue_this_sheet As Worksheet

' komorka a2 w arkuszu input
Private begOfInputList As Range
' koniec listy wejsciowej za pomoca metody end xlDown
Private endOfInputList As Range

' caly zakres danych liczac tylko z perspektywy
' pierwszej tylko kolumny
Private firstColumnScopeOnInitWorksheet As Range


' alias do komorki w arkuszu register
Private HOW_MANY_DAYS_FOR_PPUS0 As Range

' on the run property
Private ontherun As Range


' obiekt gragiczny reprezentujacy status
Private sh As StatusHandler


' jak pieknie widac na zalaczonym obrazku - pseudo konstruktor klasy ff light
' sluzy tylko i wylacznie do okreslenia zakresow wejsciowych
' dosyc zabawne myslac z perspektywy nazwy tej metody
' ale to raczej wina ze nie ma mozliwosci argumentowania konstruktora w vba
' jesli wrocisz poprzeczke wyzej do ModuleMain to zobaczysz ze new FireFlakeLight jest metoda grzecznosciowa
' i glownym elementem uruchomieniowym jest run na odpowiednim raporcie jak na przyklad
' Public Sub runDaily(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Date)
' ^^^^^^^^^^^^^^^^^^^^^
' mozna rzec ze to jest nasz konstruktor juz bardziej przyjmuje wejscie od usera i na tej bazie mozna juz pracowac
' koniec koncow taka a nie inna logika dzialania tego tutaj ponizej pseudo konstruktora jest miejsce do initu
' to uzupelniam nim to co nie wymaga customowej akcji ze strony usera.
Private Sub Class_Initialize()
    Set rep = Nothing
    Set imgo = New MGO
    
    Set sh = New StatusHandler
    sh.init_statusbar 100
    sh.show
    
    
    Set HOW_MANY_DAYS_FOR_PPUS0 = ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0")
    Set ontherun = ThisWorkbook.Sheets("register").Range("ontherun")
    ontherun = 1
    
    Set begOfInputList = ThisWorkbook.Sheets("input").Range("a2")
    If begOfInputList.Offset(1, 0) <> "" Then
        Set endOfInputList = begOfInputList.End(xlDown)
        Set firstColumnScopeOnInitWorksheet = Range(begOfInputList, endOfInputList)
    Else
        Set endOfInputList = begOfInputList
        Set firstColumnScopeOnInitWorksheet = begOfInputList
    End If
    
    
    Set init_sh = begOfInputList.Parent
    Set continue_this_sheet = Nothing
    Set cebeer = New ContinueBrokenReportHandler
    
End Sub

Private Sub Class_Terminate()
    Set rep = Nothing
    Set imgo = Nothing
    ontherun = 0
    
    
    sh.hide
    Set sh = Nothing
    Set cebeer = Nothing
    Set wh = Nothing
    MsgBox "ready!"
End Sub




' run on std reports
' dodatakowo trzeba bedzie wziazc pod uwage ze raczej nie bede chcial zadnych wiecej dodatkow
' jesli chodzi o ta odslone ff'a chodzi o to aby zyl on w zgodzie z jego wlasnym nazewnictwem tj. LIGHT!
' problem pojawia sie i tak na konfigurowalnosci fire flakea mam 3 metody ktore niezaleznie pracuja od
' siebie chyba zeby stworzyc im wspolne komponenty - i raczej w ta strone bede szedl
' poniewaz kazdy z tych runow koniec koncow bedzie mial wiele cech wspolnych jak chocby algorytm wyciagania
' asn/ pus z po400 czy tez ppum0 - no i tez koniec weryfikacja danych na podstawie ekranu histori
' musze ponownie sprawdzic na jakiej zasadzie powinno to dzialac
Public Sub runDaily(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Date)

    ' zrobimy najpierw schemat w komentarzu
    ' zaczynamy od tego jaki to jest rodzaj startu
    
    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        
        ' zanim zaczniemy raport dobrze by bylo niezaleznie zrobic walidacje listy wejsciowej
        
        
        
        Set rep = New DailyListReport
        ' metoda connect with input list bardziej by nazwa pasowala
        ' teraz wazna sprawa dzialanie calych iteracji opierac sie badzie na obiektach chwytajacych
        ' i tutaj uwaga juz na startcie to jest na metodzie initujacej uruchamiam kolekcje wszsytkich iteracji
        ' oczywiscie na poczatku wypelnione zostana tylko dane wejsciowe w postaci pn & plt itd
        ' cala reszta pozostanie nothing dopoty dopoki nie uruchomie nastpenych metod.
        
        
        
        
        
        ' trzy glowne metody pracy na kazdym raporcie light fire flake
        ' 1. init - przygotowuje rusztowanie pod wszystkie dane
        
        
        ' OBSOLETE?
        ' 2. connect with mgo chyba nie wymaga szerwszego komentarza
        '   uruchamia nam przy okazji wszelkie pomocniki obiektowe ktore pozwalaja nam zapomniec o pracy na adresach ale bezposrednio
        '   odnosiemy sie nazwami do pol sesji mgo
        
        ' 3. repRun jest to juz gotowa metoda ktora czerpie z dwoch poprzednich czesciowo wypelniona kolekcja pustymi obiektami zostaje zapelniona wartosciami
        '   z sesji mgo.
        
        rep.repInit firstColumnScopeOnInitWorksheet, limit, lt, st, CDate(daily_rqm_limit)
        rep.repRun imgo, sh
        
        
        
        
        ' =======================================================================================================
        
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        ' a wiec z tego co rozumiem continue jest realizowany
        ' poza scopem tej procedury!
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If
    
    ' to jest troche chore ale narazie nie mam lepszego pomyslu
    rep.fakeDataChangeToRecalcFirstRunoutFormula
    ' rep.adjustLayoutAfterAllJobs

End Sub

Public Sub runHourly(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Integer)

    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        MsgBox "no implmentation in light! Should never appear!"
        
        ' =======================================================================================================
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If

End Sub

Public Sub runWeekly(limit As Date, lt As LAYOUT_TYPE, st As START_TYPE, daily_rqm_limit As Integer)

    If st = FROM_THE_BEGINNING Then
        ' tutaj wystepuje metoda ktora nie przejmuje sie danymi do okola ale lista wejsciowa
        ' oczywiscie catch error musi byc na miejscu aby latwo by wylapac czy to aby napewno to o co nam chodzi
        ' =======================================================================================================
        
        MsgBox "no implmentation in light! Should never appear!"
        
        ' =======================================================================================================
        
    ElseIf st = CONTINUE_BROKEN_ONE Then
        ' forbidden state
        ' tutaj musi byc zlapany blad
        MsgBox "ten msgbox nigdy nie moze sie pojawic"
    End If

End Sub

Public Sub continueBrokenReport(lt As LAYOUT_TYPE, st As START_TYPE)

    ' =======================================================================================================
    ' =======================================================================================================
    ' chcemy ogarnac jak najszybciej pierwsza wersje do testu - zatem pomijamy inne implementacje
    MsgBox "not implemented yet!"
    End
    ' =======================================================================================================
    ' =======================================================================================================

    ' nie ptorzebuje zadnych dodatkowych danych
    ' sprawdz czy lista wejsciowa dalej jest zsynchr razem ze zepsutym reportem
    
    ' tutaj warunek dla st w ogole nie jest potrzebny :D
    ' poniewaz procedura i tak nie przyjmuje zadnego argumentu :D
    ' lecimy od razu z algorytm odzyskiwania informacji jak dalej wykonac niedokonczony raport ff
    ' =======================================================================================================
    '
    ' SUPER WAZNE
    ' continue broken report dziala tylko i wylacznie daily i list!
    ' stad if nawet w ogole nie jest brane pod uwage sprawdzania
    '
    ' =======================================================================================================
    
    ' troche kontrowersyjna ta zmienna stworzona tylko na potrzeby tej procedury wlasciwie
    ' wiec musze sie zastanowic nad sensem jej istnienia jako pola
    ' a wrzucic jako zmienna lokalna
    Set continue_this_sheet = ActiveSheet
    ' init_sh - tutaj mam liste czesci
    ' teraz te dwa akursze nalezy odpowiednio sprawdzic
    ' CeBeEr - continue broken report
    ' Private cebeer As ContinueBrokenReportHandler
    
    ' sprawdzamy czy mozemy w ogole kontynuowac ten raport
    ' -4 = dane za stare
    ' -3 = to nie jest rarpot DAILY - sprawdzanie ogólenie routingu na a1
    ' -2 = to nie jest raport LIST - sprawdzanie ogólenie routingu na a1
    ' -1 = nie mozna kontynuowac tego reportu - super critical nawet nie przeszedl logiki
    ' 0 = nie ma dopasowania danych miedy lista wejsciowa a raportem
    ' 1 = wyrzucamy ostatnia czesc i lecimy z kontynuacja
    Dim d As Dictionary
    Set d = New Dictionary
    d.Add -4, "Data at least from yesterday, report can not be completed"
    d.Add -3, "It's not even Daily FF Report"
    d.Add -2, "This is not LIST Layout"
    d.Add -1, "Ciritical error, run on continue broken report can not be completed"
    d.Add 0, "Init data are not inline with broken report"
    d.Add 1, "Config OK, we can go on with report"
    
    
    Dim wynik_funkcji As Integer
    wynik_funkcji = cebeer.checkInputAndActiveSheetBeforeJunction(init_sh, continue_this_sheet)
    If Int(wynik_funkcji) = 1 Then
    
        ' ok wszystkie dane powinny byc ok
        ' zatem pora przejsc do do uruchomienia raportu z odpowiedniego miejsca
        ' teoretycznie sprawa jest prosta - odpwiednie przesuniecie rowFlag i po roboecie ? :>
        
        ' najpiew usune ostatni wiersz, gdy podczas zatrzymywania 0 zazwyczaj konczy sie w polowie zatem
        ' nie bedziemy juz wchodzic w szczegoly ostatniej iteracji
        ' po prostu zaczniemy ja od poczatku
        cebeer.deleteLastRow continue_this_sheet
        cebeer.prepareStart
        
        Dim rng As Range
        Set rng = cebeer.getTheBeginning
        
        ' ok w input zaczynamy od 2 wiersza
        ' lista od 5 wiersza roznica oczywista
        Set rng = cebeer.prepareRestOfData()
        
        
        
        
        ' kopiuj wklej z runRepDaily
        Set rep = New DailyListReport
        ' metoda connect with input list bardziej by nazwa pasowala
        ' teraz wazna sprawa dzialanie calych iteracji opierac sie badzie na obiektach chwytajacych
        ' i tutaj uwaga juz na startcie to jest na metodzie initujacej uruchamiam kolekcje wszsytkich iteracji
        ' oczywiscie na poczatku wypelnione zostana tylko dane wejsciowe w postaci pn & plt itd
        ' cala reszta pozostanie nothing dopoty dopoki nie uruchomie nastpenych metod.
        
        
        ' trzy glowne metody pracy na kazdym raporcie light fire flake
        ' 1. init - przygotowuje rusztowanie pod wszystkie dane
        
        ' 2. connect with mgo chyba nie wymaga szerwszego komentarza
        '   uruchamia nam przy okazji wszelkie pomocniki obiektowe ktore pozwalaja nam zapomniec o pracy na adresach ale bezposrednio
        '   odnosiemy sie nazwami do pol sesji mgo - ale to chyba jest juz OBSOLETE?
        
        ' 3. repRun jest to juz gotowa metoda ktora czerpie z dwoch poprzednich czesciowo wypelniona kolekcja pustymi obiektami zostaje zapelniona wartosciami
        '   z sesji mgo.
        Dim pus_limit As Date, daily_rqm_limit As Date
        pus_limit = cebeer.setPusLimit(continue_this_sheet.Range("A1"))
        daily_rqm_limit = cebeer.setRqmLimit(continue_this_sheet.Range("A1"))
        rep.repInit rng, CDate(pus_limit), lt, st, CDate(daily_rqm_limit), cebeer.getTheBeginning.Row
        rep.repRun imgo, sh
        
        ' wsadzone swiezo - by przeliczyc fakowo first runouty
        rep.fakeDataChangeToRecalcFirstRunoutFormula
    Else
        MsgBox CStr(d.item(Int(wynik_funkcji)))
    End If
    
    
    
    
    ' =======================================================================================================
End Sub
