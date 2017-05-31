VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakePMScaffold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' referencje pod kolekcje
Private iter As IIteration
Private item As IItem
Private it As ITransit
Private sitem As FireFlakePMScaffoldItem
Private sitemCollection As Collection

Private ILE_DNI As Long


Public Property Get GET_ILE_DNI() As Long
    GET_ILE_DNI = ILE_DNI
End Property


Public Property Get SCAFFOLD_ITEM(x) As FireFlakePMScaffoldItem
    
    Set SCAFFOLD_ITEM = sitemCollection.item(x)
End Property



' jestesmy w klasie ktora uruchamia obiekt rusztowania w klasie FireFlakePMCollection
' jest to wlasciwie boxowanie poniewaz klasa fire flake pm coll sama siebie za wiele nie robi
' oprocz dodatkowego sprawdzenia jak ma sie instancje tejze tutaj klasy
' glowna procedura uruchomieniowa po konstruktorze jest sub make - przeglada on kolekcje danych
' ktora wczesniej logika wyodrebnila z danych z mgo z odpowiednich ekranow
' uklad ma byc taki, ze kazda labelka it to nowy PUS
Public Sub make(main_collection As Collection)


    Set sitemCollection = New Collection

    
    ' algorytm n^2
    'For Each iter In c
    '    For Each item In iter
    '    Next item
    'Next c
    
    
    
    Dim item_i As Integer, maks As Long
    
    ' zaczynamy od przejrzenia pierwszego dnia
    ' dla kazdego dnia musismy podliczyc ile w ogole wystapilo tranzytow sumarycznie
    ' nie bedziemy robic zadnych fancy utlizacji obszaru raportu
    ' prosto lecimy po wszystkim i tyle tez ustawiamy kolumn danych
    item_i = 1
    maks = 0
    
    For Each iter In main_collection
        If maks < iter.getItemsCollection.COUNT Then
            maks = iter.getItemsCollection.COUNT
        End If
    Next iter
    
    ' przeliczenie ilosci itemow sumarycznie
    ILE_DNI = maks
    
    Do
        Set sitem = New FireFlakePMScaffoldItem
        sitem.pozycja = item_i
        sitem.ilosc_pol_tranzytow = 1
        sitem.ile_kolumn = 3
        
        
        ' item_i = 1
        For Each iter In main_collection
        
            If item_i <= iter.getItemsCollection.COUNT Then
                Set item = iter.getItemsCollection.item(item_i)
                
                
                
                
                
                
                If Not item.getInTransits Is Nothing Then
                
                
                    ' obsoletowa logika ustawiajaca ilosc kolumn pod max ilosc tranzytow dla jednego pnu
                    ' zalozenia projektowe beda jednak inne
                    ' --------------------------------------------------------------------------------------
                    ' --------------------------------------------------------------------------------------
                    'If Int(item.getInTransits.COUNT) > sitem.ilosc_pol_tranzytow Then
                    '
                    '    sitem.ilosc_pol_tranzytow = Int(item.getInTransits.COUNT)
                    'End If
                    ' --------------------------------------------------------------------------------------
                    ' --------------------------------------------------------------------------------------
                    
                    ' nowa logika po prostu wszystko sumuje
                    ' ======================================================================================
                    If IsNumeric(item.getInTransits.COUNT) Then
                        If item.getInTransits.COUNT > 0 Then
                        
                            For Each it In item.getInTransits
                                sitem.proba_dodania_pusa_jako_nowej_labelki it
                            Next it
                            
                            
                        End If
                        
                        
                    End If
                    ' ======================================================================================
                    
                    
                Else
                    sitem.ilosc_pol_tranzytow = 1
                End If
                
            Else
                Exit For
            End If
        Next iter
        
        sitem.dodanie_manuala
        
        sitem.inner_calc
        
        sitemCollection.Add sitem
        
        item_i = item_i + 1
        
        If CLng(item_i) > CLng(ILE_DNI) Then
            Exit Do
        End If
    Loop While True
    
    
End Sub
