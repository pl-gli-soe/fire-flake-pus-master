VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakePMCollection"
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



' poniewaz sama kolekcja w przypadku tak skomplikowanego tworu jak PUS Master
' - nieregularne ksztalty raportu (odbiega od powtarzalnej formy)
' brak formy tabelarycznej, badz z offsetem 0,3 jak w przypadku std listowego fire flake'a


' opis
' ----------------------------------------------------------------

' niniejsza klasa jest obudowa dla pierwotnej kolekcji obslugujacej dane
' w FFL

' wczesniej kolekcja nie potrzebowala dodatkowych danych, zatem sama z siebie
' w zupelnosci wystarczala.
' jednak podniesienie wyrafinowania prezentowania informacji na raporcie
' (nieregularne ukazywania nazw pusow)

' ----------------------------------------------------------------


Private kolekcja As Collection
Private rusztowanie As FireFlakePMScaffold


Public Property Get scaffold() As FireFlakePMScaffold
    Set scaffold = rusztowanie
End Property


Public Property Let c(mc As Collection)
    Set kolekcja = mc
End Property

Public Property Get c() As Collection
    Set c = kolekcja
End Property


Public Sub calc_collection()
    
    ' pierwszy podstawowy warunek sprawdzajcy czy jest w ogole co liczyc
    ' w konstruktorze klasy jest przypisanie do nothing, zatem jesli nie
    ' swiadomego przypisania kolekcji danych do obiektu, to nie mamy na co liczyc.
    If Not kolekcja Is Nothing Then
    
        Set rusztowanie = New FireFlakePMScaffold
        rusztowanie.make kolekcja
    End If
End Sub



Private Sub Class_Initialize()
    Set kolekcja = Nothing
    Set rusztowanie = Nothing
End Sub

Private Sub Class_Terminate()
    Set kolekcja = Nothing
    Set rusztowanie = Nothing
End Sub

