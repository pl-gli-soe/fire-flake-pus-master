VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PDRQM"
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

Public Property Get infocode() As String
    infocode = Trim(sess0.screen.getString(22, 2, 5))
End Property

Public Property Get plt() As String
    plt = Trim(sess0.screen.getString(3, 8, 2))
End Property

Public Property Let plt(arg As String)
    sess0.screen.putString CStr(arg), 3, 8
End Property

Public Property Get pn() As String
    pn = Trim(sess0.screen.getString(4, 8, 8))
End Property

Public Property Let pn(arg As String)
    sess0.screen.putString CStr(arg), 4, 8
End Property

Public Property Get KANBAN() As String
    KANBAN = Trim(sess0.screen.getString(5, 10, 4))
End Property

Public Property Let KANBAN(arg As String)
    sess0.screen.putString CStr(arg), 5, 10
End Property

Public Property Get PART_DESC() As String
    PART_DESC = Trim(sess0.screen.getString(4, 41, 10))
End Property



Public Function dates(SCREEN_COLUMN As Integer) As String
    dates = Trim(sess0.screen.getString(8, 24 + 8 * SCREEN_COLUMN, 5))
End Function



' metoda zostala zubozona i uzalezniona od metody ponizej, lecz dzieki temu kod zostanie zoptymalizowany i przyspieszony
Public Function PLT_TOTAL(SCREEN_COLUMN As Integer, ByRef plttotalrow As Integer) As String
    PLT_TOTAL = Trim(sess0.screen.getString(plttotalrow, 22 + 8 * SCREEN_COLUMN, 7))
End Function


' pomimo tego ze jest to metoda publiczna
' jest scisle polaczona z logika PLT_TOTAL
' sluzy tylko i wylacznie podliczeniu w ktorym wieszu znajduje sie label PLT TOTAL
' moglo by byc tak jak jest w jezykach funkcyjnych to jest definicja metody ktora jest z miesjca uruchamiana, ale nie jestem do konca pewien ze nigdzie indziej jej
' wykorzystac nie bede
Public Function findPLTTOTALrow() As Integer

    For x = 9 To 21
        If Trim(sess0.screen.getString(x, 2, 9)) Like "*PLT TOTAL*" Then
            findPLTTOTALrow = x
            Exit For
        End If
    Next x
End Function
