VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailyDynamicColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IDynamicColors


Private ash As Worksheet
Private flagaLiczenia As Boolean
Private zasiegDoNadaniaDynamicznychKolorow As Range
Private zasiegCBala As Range
Private jakiOffsetDoPierwszegoRqm As Integer
Private r As Range



' colors
Private redColor As Variant
Private pinkColor As Variant
Private blackColor As Variant
Private yellowColor As Variant
Private greyColor As Variant


Private Sub Class_Initialize()

    If CStr(ThisWorkbook.FullName) = CStr(ActiveWorkbook.FullName) Then
        Set ash = ActiveSheet
        flagaLiczenia = True
        
        redColor = ThisWorkbook.Sheets("register").Range("red").Interior.Color
        pinkColor = ThisWorkbook.Sheets("register").Range("pink").Interior.Color
        blackColor = ThisWorkbook.Sheets("register").Range("black").Interior.Color
        yellowColor = ThisWorkbook.Sheets("register").Range("yellow").Interior.Color
        greyColor = ThisWorkbook.Sheets("register").Range("grey").Interior.Color
        
        
        
    Else
        Set ash = Nothing
        flagaLiczenia = False
    End If
End Sub

Private Sub IDynamicColors_assignDynamicColorsrange()

    If flagaLiczenia Then
        ' Set zasiegDoNadaniaDynamicznychKolorow = ash.Range("b4")
        ' tutjaj zalozenie jest takie ze pracujemy na liscie bez wiekszego rozstrzasania
        Set zasiegDoNadaniaDynamicznychKolorow = przeliczZasieg(4, 2)
        Set zasiegCBala = przeliczCBala(4, 4)
    End If
    
    ' ' Debug.Print zasiegDoNadaniaDynamicznychKolorow.Address
End Sub

Private Function przeliczCBala(r, c) As Range

    
    ' Set przeliczZasieg = Nothing
    Dim rr As Range
    ' MsgBox ActiveSheet.Name
    Set rr = Cells(Int(r), Int(c))
    
    
    ' przy okazji policze offset
    ' na poczatku zaporowa wartosc
    ' jesli -1 jeden pozostanie to znaczy ze cos poszlo nie tak
    jakiOffsetDoPierwszegoRqm = -1
    jakiOffsetDoPierwszegoRqm = Int(goToFirstRqmField(rr, Int(c)))
    
    
    Dim poczatek As Range
    Dim ost_wiersz_pierwszej_kolumny As Range
    
    
    ' ale to jest uklad jak zaczynamy od 5 wiersza co nie jest prawdda
    ' moze byc tak ze mamy tylko jeden wiersz danych
    'If Trim(rr.Offset(2, 0)) <> "" Then
    '    Set ost_wiersz_pierwszej_kolumny = rr.End(xlDown)
    'End If
    

    Set ost_wiersz_pierwszej_kolumny = rr.End(xlDown)

    
    'Do
    '    Set rr = rr.Offset(0, 1)
    'Loop While Trim(CStr(rr)) <> "Past due rqms"
    
    Set poczatek = rr.Offset(1, 0)
    
    Set przeliczCBala = Range(poczatek, ost_wiersz_pierwszej_kolumny)
End Function

Private Function goToFirstRqmField(ByRef rr As Range, c As Integer) As Integer
    goToFirstRqmField = -1
    
    Dim tmp As Range
    Set tmp = rr
    Do
        If Trim(tmp) = "Requirements" Then
            goToFirstRqmField = Int(tmp.Column - c)
            Exit Function
        End If
        Set tmp = tmp.Offset(0, 1)
    Loop Until Trim(tmp) = ""
End Function

Private Function przeliczZasieg(r, c) As Range

    
    ' Set przeliczZasieg = Nothing
    Dim rr As Range
    ' MsgBox ActiveSheet.Name
    Set rr = Cells(Int(r), Int(c))
    
    Dim poczatek As Range
    Dim ost_wiersz_pierwszej_kolumny As Range
    Dim ost_kolumna As Range
    
    
    ' ale to jest uklad jak zaczynamy od 5 wiersza co nie jest prawdda
    ' moze byc tak ze mamy tylko jeden wiersz danych
    'If Trim(rr.Offset(2, 0)) <> "" Then
    '    Set ost_wiersz_pierwszej_kolumny = rr.End(xlDown)
    'End If
    
    Set ost_wiersz_pierwszej_kolumny = rr.End(xlDown)
    Set ost_kolumna = rr.End(xlToRight)
    
    Do
        Set rr = rr.Offset(0, 1)
    Loop While Trim(CStr(rr)) <> "Past due rqms"
    
    Set poczatek = rr.Offset(1, 1)
    
    Set poczatek = Range(poczatek, Cells(ost_wiersz_pierwszej_kolumny.Row, ost_kolumna.Column))
    Set przeliczZasieg = poczatek
End Function


Private Function IDynamicColors_getDataScope() As Range
    Set IDynamicColors_getDataScope = zasiegDoNadaniaDynamicznychKolorow
End Function

Private Sub IDynamicColors_recalcColors()


    If flagaLiczenia Then
    
        If jakiOffsetDoPierwszegoRqm > -1 Then
        
            For Each r In zasiegCBala
                If CLng(r) < 0 Then
                    r.Interior.Color = CLng(redColor)
                    r.Font.Color = CLng(blackColor)
                ElseIf CLng(r.Offset(0, jakiOffsetDoPierwszegoRqm)) > CLng(r) Then
                    r.Interior.Color = CLng(pinkColor)
                    r.Font.Color = CLng(blackColor)
                Else
                    r.Interior.Color = Cells(4, r.Column).Interior.Color
                    r.Font.Color = Cells(4, r.Column).Font.Color
                End If
            Next r
            
        End If
    
        If (Not zasiegDoNadaniaDynamicznychKolorow Is Nothing) Then
    
        
            For Each r In zasiegDoNadaniaDynamicznychKolorow
                If CStr(Cells(4, r.Column)) = CStr("Ending Balance") Then
                    
                    ' pierwszy warunek spelniony to jest kolor na ebal
                    ' teraz przechodzimy i sprawdzamy jaka wartosc ma ending balance i jesli ma cos konkret
                    ' to dajemy odpowiednio red or pink
                    
                    If CLng(r) < 0 Then
                        r.Interior.Color = CLng(redColor)
                        r.Font.Color = CLng(blackColor)
                    ElseIf CLng(r.Offset(0, 1)) > CLng(r) Then
                        r.Interior.Color = CLng(pinkColor)
                        r.Font.Color = CLng(blackColor)
                    Else
                        r.Interior.Color = Cells(4, r.Column).Interior.Color
                        r.Font.Color = Cells(4, r.Column).Font.Color
                    End If
                End If
            Next r
        End If
    End If
End Sub
