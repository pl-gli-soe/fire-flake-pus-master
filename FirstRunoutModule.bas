Attribute VB_Name = "FirstRunoutModule"
' to jest jedna z najwazniejszych funkcji calego projektu
' zatem totez dlatego ma nawet swoj wlasny modul
' co by nie bylo watpliwosci
' gdybym chcial powaznie
' rozmyslac o ewentualnej przyszlej rozbudowie tej oto funkcji
Public Function firstRunout(r As Range, caly_zasieg As Range) As String
    firstRunout = ""
    
    'wiersz_w_ktorym_znajduja_sie_daty = 3
    'wwkzsd = wiersz_w_ktorym_znajduja_sie_daty
    
    Dim sh As Worksheet, rng As Range, ebal_flag As Range
    Set sh = r.Parent
    If sh.Range("a1") Like "*;LIST;*" Then
        Set rng = sh.Range("a4")
        Do
            Set rng = rng.Offset(0, 1)
                
            If rng = "" Then
                firstRunout = "no data"
                Exit Function
            End If
        Loop Until CStr(rng) = CStr(ff_pm.STR_EBAL)
    End If
    
    ' prosciej byc nie moze
    Set ebal_flag = rng
    
    ' ' Debug.Print ebal_flag
    
    Do
    
        If CStr(sh.Cells(4, ebal_flag.Column)) = ff_pm.STR_EBAL Then
            If sh.Cells(r.Row, ebal_flag.Column) < 0 Then
            
                For q = 0 To -100 Step -1
                    If IsDate(ebal_flag.Offset(-1, q)) Then
                        firstRunout = ebal_flag.Offset(-1, q)
                        Exit Do
                    End If
                    
                    If ebal_flag.Offset(-1, q).Column = 1 Then
                        Exit Do
                    End If
                Next q
            End If
        End If
        Set ebal_flag = ebal_flag.Offset(0, 1)
    Loop Until Trim(ebal_flag) = ""
    
    If firstRunout = "" Then

        firstRunout = CStr(CDate(Date + 1000))
    End If
    
    
    ' to jest przeklamanie!
    ' firstRunout = ebal_flag.Offset(-1, -5)
    
End Function
