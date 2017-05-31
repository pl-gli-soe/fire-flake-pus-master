VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContinueBrokenReportHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private init_sh As Worksheet
Private active_sh As Worksheet

Private input_rng_flag As Range
Private report_rng_flag As Range

' na trzecim wierszu powstawiane sa
' kolejne daty coverage
' pierwsza data musi byc data dzisiejsza
Private first_date_from_list_raport As Range

Private poczatek_danych_dla_continue_broken_report As Range



Public Function prepareRestOfData() As Range
    Set prepareRestOfData = Nothing
    
    wiersz = poczatek_danych_dla_continue_broken_report.Row - 2
    
    Set input_rng_flag = init_sh.Cells(1, 1)
    x = 1
    Do
        Set input_rng_flag = input_rng_flag.Offset(1, 0)
        
        If Not input_rng_flag.EntireRow.Hidden Then
            x = x + 1
        End If
    Loop Until x = wiersz
    
    ' Set prepareRestOfData = input_rng_flag
    
    
    If Trim(input_rng_flag.Offset(1, 0)) <> "" Then
        Set input_rng_flag = init_sh.Range(input_rng_flag, input_rng_flag.End(xlDown))
        Set prepareRestOfData = input_rng_flag
    Else
        Set prepareRestOfData = input_rng_flag
    End If

End Function


Public Function setPusLimit(r As Range) As Date
    ' setPusLimit
    
    arr = Split(CStr(r), ";")
    item = CStr(arr(5))
    
    item = Replace(item, "PUS_LIM:", "")
    
    setPusLimit = CDate(item)
End Function

Public Function setRqmLimit(r As Range) As Date
    ' setRqmLimit
    
    arr = Split(CStr(r), ";")
    item = CStr(arr(4))
    item = Replace(item, "RQM_LIM:", "")
    
    setRqmLimit = CDate(item)
End Function

Public Function getTheBeginning() As Range
    Set getTheBeginning = poczatek_danych_dla_continue_broken_report
End Function

Public Sub deleteLastRow(ByRef actives As Worksheet)

    ' checkInput.. juz zostalo wypisane zatem pomimo przesuniecia argumentu zostawie to tak jak jest
    ' tj nie bede znow nadpisywal tych pol
    'Set active_sh = actives
    ' Set report_rng_flag = active_sh.Range("b5")
    
    Set report_rng_flag = active_sh.Range("b5")
    If active_sh.Range("b6") <> "" Then
        Set report_rng_flag = report_rng_flag.End(xlDown)
    End If
    
    report_rng_flag.EntireRow.Delete
    
    
    
End Sub

Public Sub prepareStart()
    
    
    Set report_rng_flag = active_sh.Range("b5")
    If active_sh.Range("b6") <> "" Then
        Set report_rng_flag = report_rng_flag.End(xlDown)
    End If
    
    Set poczatek_danych_dla_continue_broken_report = report_rng_flag
End Sub



' -4 = dane za stare
' -3 = to nie jest rarpot DAILY - sprawdzanie ogólenie routingu na a1
' -2 = to nie jest raport LIST - sprawdzanie ogólenie routingu na a1
' -1 = nie mozna kontynuowac tego reportu - super critical nawet nie przeszedl logiki
' 0 = nie ma dopasowania danych miedy lista wejsciowa a raportem
' 1 = wyrzucamy ostatnia czesc i lecimy z kontynuacja

Public Function checkInputAndActiveSheetBeforeJunction(ByRef inits As Worksheet, ByRef actives As Worksheet) As Integer
    checkInputAndActiveSheetBeforeJunction = -1
    
    Set init_sh = inits
    Set active_sh = actives
    
    
    ' ok no to lecimy
    ' wiadomka ze uklad na input liscie jest zawsze taki sam
    ' kolumna1 - plt
    ' kolumna2 - part number
    
    ' z arkuszem raportu jest juz troche wiecej zabawy
    If active_sh.Range("a1").Value Like "*REPORT;DAILY;*" Then
        If active_sh.Range("a1").Value Like "*REPORT;DAILY;LIST;*" Then
        
            ' teraz ostatni warunek - czyli sprawdzenie czy pierwszy item na liscie to dzien dzisiejszy
            Set first_date_from_list_raport = active_sh.Range("a1")
            arr = Split(first_date_from_list_raport, ";")
            data_wyprodukowania_raportu = arr(3)
            
            If CStr(Format(Now, "yyyy-mm-dd")) = CStr(Format(CDate(data_wyprodukowania_raportu), "yyyy-mm-dd")) Then
            
            
                ' ok teraz mozemy zaczac proces sprawdzania listy wejsciowej
                ' konkretne niezmienne komorki acz wydaje mi sie ze nie powinno to tak wygladac jesli bede chcial srogo zmienic layout to sie wysypie
                '
                ' narazie to tak zostawie
                ' ale musze pamietac
                ' ze tutaj slabo wpisane zostalo na stale i mody na layoucie spowoduje errory
                Set input_rng_flag = init_sh.Range("a2")
                If Not input_rng_flag.EntireRow.Hidden Then
                    ' Set input_rng_flag = init_sh.Range("a2")
                    ' no additional operation needed
                ElseIf input_rng_flag.EntireRow.Hidden Then
                    Do
                        Set input_rng_flag = input_rng_flag.Offset(1, 0)
                    Loop Until (Not input_rng_flag.EntireRow.Hidden)
                End If
                Set report_rng_flag = active_sh.Range("b5")
                
                Do
                    
                    If CStr(input_rng_flag) & CStr(input_rng_flag.Offset(0, 1)) = CStr(report_rng_flag.Offset(0, 1)) & CStr(report_rng_flag) Then
                        ' nop
                    Else
                        checkInputAndActiveSheetBeforeJunction = 0
                        Exit Function
                    End If
                
                    Set input_rng_flag = input_rng_flag.Offset(1, 0)
                    If Not input_rng_flag.EntireRow.Hidden Then
                        ' Set input_rng_flag = init_sh.Range("a2")
                        ' no additional operation needed
                    ElseIf input_rng_flag.EntireRow.Hidden Then
                        Do
                            Set input_rng_flag = input_rng_flag.Offset(1, 0)
                        Loop Until (Not input_rng_flag.EntireRow.Hidden)
                    End If
                    Set report_rng_flag = report_rng_flag.Offset(1, 0)
                    
                    ' teraz wazna rzecz
                    ' ze jesli mamy do czynienia z broken report
                    ' to jest tak ze lista wejsciowa bedzie wieksza od zatrzymanego raportu
                    ' wstepne sprawdzenie zatem musi byc takie, ze lista juz wygenerowannych
                    ' pnow musi sie zgadzac z lista wejsciowa
                    ' nie jest porownanie kompletne ale mysle ze wystaraczajace
                    
                Loop Until Trim(report_rng_flag) = ""
                
                checkInputAndActiveSheetBeforeJunction = 1
            Else
                checkInputAndActiveSheetBeforeJunction = -4
                Exit Function
            End If
            
        Else
            checkInputAndActiveSheetBeforeJunction = -2
            Exit Function
        End If
    Else
        ' to nawet nie jest daily report
        checkInputAndActiveSheetBeforeJunction = -3
        Exit Function
    End If
    
End Function
