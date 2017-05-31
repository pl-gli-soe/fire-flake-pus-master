VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicEventOnControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta wlasciwie jest kontenerem przechowujacym metody do wywolania w przypadku wystapienia zdarzenia
' this class is in fact container for methods which will be called in case of proper event

' but anyway
' bardziej comment dla mnie niz dla kogo innego, gdyz poniewaz zakladam
' ze malo kto bedzie zainteresowany customowym przechwytywaniem zdarzen z poziomu vba
' nawet nie wiedzialem ze sa tu takie mozliwosci (jeszcze nie tak dawno temu)
' myslalem ze eventy ograniczaja sie tylko do formow samych w sobie
' potem potrafilem budowac eventy na samuch customowych obiektach
' a teraz nie dosc ze dynamicznie buduje formy to at runtime wpisuje te obiekty do kolekcji "eventowej"
'
' jesli chodzi o wykorzystanie tej klasy DynamicEventOnControls (DEOC)
' to jest to glowne pole dla DownloadingRulesHandler (drh) which has collection for creating at runtime some events

Option Explicit

Private WithEvents handler As MSForms.ComboBox
Attribute handler.VB_VarHelpID = -1
Private WithEvents nameh As MSForms.TextBox
Attribute nameh.VB_VarHelpID = -1
Private WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Private pEventSelected As Control
Private handler_name As Integer



Private Sub handler_Change()
    ' ' Debug.Print "tutaj arkusz powinien sie dopasowac"
    
    ThisWorkbook.Sheets("register").Cells(Int(handler_name) + 2, CONFIG_REG_PLT_COLUMN + 1) = CStr(handler.Value)
    lbl.Caption = "adjusted for " & CStr(handler.Value)
    
    
    ' prawie jak konstruktor - tyle ze nie kasuje danych nieodswiezanych :)
    ' co smieszne implementacja wyglada tak samo jak w new DRH
    'drh.inicjacja
    'drh.okresl_wielkosc_forma_i_przesun_guziki_w_dol
    'DownloadingRulesForm.Repaint
End Sub


Public Sub setHandlerName(s As Integer)
    handler_name = Int(s)
End Sub

Public Function getHandlerName() As String
    getHandlerName = handler_name
End Function


Public Property Set Control(p As MSForms.ComboBox)
    Set handler = p
End Property

Public Property Set textBoxControl(p As MSForms.TextBox)
    Set nameh = p
End Property

Public Property Set setLbl(p As MSForms.Label)
    Set lbl = p
End Property




Private Sub nameh_Change()
    ThisWorkbook.Sheets("register").Cells(Int(handler_name) + 2, CONFIG_REG_PLT_COLUMN) = CStr(nameh.Value)
End Sub
