VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPUD0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sess0 As Object

Public Property Let sesja(ByRef arg As Object)
    Set sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = sess0
End Property

Public Property Let pus(arg As String)
    sess0.screen.putString CStr(arg), 4, 39
End Property

Public Property Get pus() As String
    pus = Trim(sess0.screen.getString(4, 39, 9))
End Property

Public Property Let plt(arg As String)
    sess0.screen.putString CStr(arg), 4, 7
End Property



' musisz tutaj mateuszek pamietac ze mamy tutaj surowe cztery liczby bez znaku ":"
Public Property Get fst_pickup_time_window() As String
    fst_pickup_time_window = Trim(sess0.screen.getString(6, 39, 4))
End Property





