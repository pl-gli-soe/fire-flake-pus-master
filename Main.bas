Attribute VB_Name = "Main"
' jest to glowna metoda ktora bedzie stand alone w przypadku podpinania do guzika tj.
' bedzie wpisana jako jedyna do subroutine ktore bedzie podpiete do guzika juz bezposrednio bez zbednych dodatkowych zabiegow
' jej wszsytkie argumenty maja w pelni ogarnac konfiguracje run ff light
' myslalem nad kombinacja alpejska w stylu zeby user mial mozliwosc konfigurowalnosci widoku ale chyba bylo by to przedobrzone :D
' jeszcze obaczym - napisal to ja 2014 wrzesien 22.
Public Sub runReport(t As RUN_TYPE, l As LAYOUT_TYPE, st As START_TYPE, p_limit As Date, daily_rqm_limit As Date)


    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ff_pm As FireFlakePUSMaster
    Set ff_pm = New FireFlakePUSMaster
    Set sh = New StatusHandler


    If st = FROM_THE_BEGINNING Then
    
    
        ' na koniec koncow light ma byc fire flak'iem tylko i wylacznie
        ' tematem daily - wersja odchudzona bez zbednych dodatkow
        ' nawet dochodzac do skrajnosci w ktorej dodalem tutaj pelno elementow
        ' layoutu jako box, list, czy coverage
        ' zostajmy tylko i wylacznie z tematem:
        ' daily i list!
        ' no i oczywiscie nowinka techniczna continue with broken :D
        ' ale to jak widac znadjuje sie po drugiej stronie barykady
        ' sam bedzie musial rozpoznac format danych jak i layout
        If t = DAILY Then
            ff_pm.runDaily CDate(p_limit), l, st, CDate(daily_rqm_limit)
        ElseIf t = HOURLY Then
            ff_pm.runHourly p_limit, l, st, CDate(daily_rqm_limit)
        ElseIf t = WEEKLY Then
            ff_pm.runWeekly p_limit, l, st, CDate(daily_rqm_limit)
        End If
    ElseIf st = CONTINUE_BROKEN_ONE Then
    
        ' ten tutaj jest bystry na tyle zeby sam siebie skonfigurowac i pociagnac temat samemu
        ' :)
        ' ff_pm.continueBrokenReport LIST_LAYOUT, st
    End If
    
    
    Set ff_pm = Nothing
    Set sh = Nothing
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' poza wszelkim zasiegiem obietkowych zagrywek - po prostu przelicz arkusz
    ' reset_report_inner
End Sub




Public Sub run_ff(ictrl As IRibbonControl)
    MainForm.show
End Sub

Public Sub reset_report_inner()

    ' dodatkowo przyda sie:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' tu ma byc reset jako taki dla odswiezenia dynamicznych kolorow
    ' teraz kwestia tylko z jakim rodzajem raportu mamy do czynienia
    Dim ash As Worksheet, dc As IDynamicColors
    If ThisWorkbook.FullName = ActiveWorkbook.FullName Then
        Set ash = ActiveSheet
        
        If CStr(ash.Range("b4")) = "Part #" And CStr(ash.Range("c4")) = "Plant" Then
            Set dc = New DailyDynamicColors
            dc.assignDynamicColorsrange
            dc.recalcColors
    
            
        End If
        
    Else
        Set ash = Nothing
    End If
    
    
    
End Sub

Public Sub reset_report(ictrl As IRibbonControl)
    reset_report_inner
End Sub
