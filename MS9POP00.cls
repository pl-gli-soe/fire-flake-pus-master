VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9POP00"
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

Public Property Let pn(arg As String)
    sess0.screen.putString CStr(arg), 3, 22
End Property

Public Property Get pn() As String
    pn = Trim(sess0.screen.getString(4, 22, 8))
End Property

Public Property Let plt(arg As String)
    sess0.screen.putString CStr(arg), 3, 13
End Property

Public Property Get plt() As String
    plt = Trim(sess0.screen.getString(4, 13, 2))
End Property

Public Property Let DS(arg As String)
    sess0.screen.putString CStr(arg), 2, 18
End Property

Public Property Get DS() As String
    DS = Trim(sess0.screen.getString(2, 18, 1))
End Property

Public Property Let duns(arg As String)
    sess0.screen.putString CStr(arg), 3, 62
End Property

Public Property Get duns() As String
    duns = Trim(sess0.screen.getString(7, 7, 9))
End Property

Public Property Get SUPPLIER() As String
    SUPPLIER = Trim(sess0.screen.getString(7, 17, 20))
End Property

Public Property Let F_U(arg As String)
    sess0.screen.putString CStr(arg), 3, 35
End Property

Public Property Get F_U() As String
    F_U = Trim(sess0.screen.getString(4, 35, 2))
End Property

Public Property Let a(arg As String)
    sess0.screen.putString CStr(arg), 3, 53
End Property

Public Property Get a() As String
    a = Trim(sess0.screen.getString(4, 53, 2))
End Property


' pierwszy doh
Public Property Let firstDOH(arg As String)
    sess0.screen.putString CStr(arg), 3, 43
End Property

Public Property Get firstDOH() As String
    firstDOH = Trim(sess0.screen.getString(4, 43, 3))
End Property

' drugi doh
Public Property Let secDOH(arg As String)
    sess0.screen.putString CStr(arg), 3, 47
End Property

Public Property Get secDOH() As String
    secDOH = Trim(sess0.screen.getString(3, 47, 3))
End Property

Public Property Get misc() As String
    misc = Trim(sess0.screen.getString(5, 73, 7))
End Property

Public Property Get os() As String
    os = Trim(sess0.screen.getString(6, 67, 7))
End Property

Public Property Get bank() As String
    bank = Trim(sess0.screen.getString(6, 54, 7))
End Property


Public Property Get bbal() As String
    If Trim(sess0.screen.getString(5, 15, 1)) = "-" Then
        bbal = "-" & Trim(sess0.screen.getString(5, 7, 8))
    Else
        bbal = Trim(sess0.screen.getString(5, 7, 8))
    End If
End Property


Public Property Get cbal() As String
    If Trim(sess0.screen.getString(5, 30, 1)) = "-" Then
        cbal = "-" & Trim(sess0.screen.getString(5, 22, 8))
    Else
        cbal = Trim(sess0.screen.getString(5, 22, 8))
    End If
End Property


Public Property Get PCS_TO_GO() As String
    PCS_TO_GO = Trim(sess0.screen.getString(6, 23, 7))
End Property

Public Property Get DK() As String
    DK = Trim(sess0.screen.getString(8, 6, 8))
End Property

Public Property Get mode() As String
    mode = Trim(sess0.screen.getString(8, 21, 2))
End Property

Public Property Get TT() As String
    TT = Trim(sess0.screen.getString(8, 31, 8))
End Property

Public Property Get MNPC() As String
    MNPC = Trim(sess0.screen.getString(9, 7, 7))
End Property

Public Property Get NCX() As String
    NCX = Trim(sess0.screen.getString(10, 6, 1))
End Property

Public Property Get OBS() As String
    OBS = Trim(sess0.screen.getString(10, 14, 1))
End Property

Public Property Get STD_PACK() As String
    STD_PACK = Trim(sess0.screen.getString(10, 30, 9))
End Property

Public Property Get oneJOB() As String
    oneJOB = Trim(sess0.screen.getString(12, 7, 11))
End Property

Public Property Get IP() As String
    IP = Trim(sess0.screen.getString(12, 27, 5))
End Property

Public Property Get COUNT() As String
    COUNT = Trim(sess0.screen.getString(14, 8, 45))
End Property

Public Property Get PART_NAME() As String
    PART_NAME = Trim(sess0.screen.getString(6, 2, 10))
End Property

Public Property Get qhd() As String
    qhd = Trim(sess0.screen.getString(5, 48, 7))
End Property

Public Property Get O() As String
    O = Trim(sess0.screen.getString(15, 4, 25))
End Property

Public Property Get F() As String
    F = Trim(sess0.screen.getString(15, 43, 38))
End Property


Public Property Get infocode() As String
    infocode = Trim(sess0.screen.getString(23, 2, 5))
End Property

Public Function transQTY(i As Integer) As String
    transQTY = Trim(sess0.screen.getString(17 + (Int(i) * 2), 2, 9))
End Function

Public Function transSDATE(i As Integer) As String
    transSDATE = Trim(sess0.screen.getString(17 + (Int(i) * 2), 18, 8))
End Function

Public Function transPCS(i As Integer) As String
    transPCS = Trim(sess0.screen.getString(17 + (Int(i) * 2), 27, 3))
End Function

Public Function transWGT(i As Integer) As String
    transWGT = Trim(sess0.screen.getString(17 + (Int(i) * 2), 31, 6))
End Function

Public Function transMD(i As Integer) As String
    transMD = Trim(sess0.screen.getString(17 + (Int(i) * 2), 38, 2))
End Function

Public Function transCD(i As Integer) As String
    transCD = Trim(sess0.screen.getString(17 + (Int(i) * 2), 41, 1))
End Function

Public Function transEDA(i As Integer) As String
    transEDA = Trim(sess0.screen.getString(17 + (Int(i) * 2), 45, 8))
End Function

Public Function transETA(i As Integer) As String
    transETA = Trim(sess0.screen.getString(17 + (Int(i) * 2), 54, 2)) & ":" & Trim(sess0.screen.getString(17 + (Int(i) * 2), 56, 2))
End Function

Public Function transST(i As Integer) As String
    transST = Trim(sess0.screen.getString(17 + (Int(i) * 2), 60, 1))
End Function

Public Function transBL(i As Integer) As String
    transBL = Trim(sess0.screen.getString(17 + (Int(i) * 2), 63, 9))
End Function

Public Function transROUTE(i As Integer) As String
    transROUTE = Trim(sess0.screen.getString(17 + (Int(i) * 2), 73, 7))
End Function

Public Function transCMNT(i As Integer) As String
    transCMNT = Trim(sess0.screen.getString(18 + (Int(i) * 2), 39, 30))
End Function

Public Function transDUNS(i As Integer) As String
    transDUNS = Trim(sess0.screen.getString(18 + (Int(i) * 2), 71, 9))
End Function

Public Function transCONT(i As Integer) As String
    transCONT = Trim(sess0.screen.getString(18 + (Int(i) * 2), 2, 35))
End Function

' -------------------------- rqms on pop ----------------------------------------------------------
Public Function balance_on_pop_data(wiersz As Integer, col As Integer) As String
    'row 8 - 12
    'col 40, 46, 56, 66, 73, size 8
    Dim kolumna As Integer
    Select Case col
        Case 1
            kolumna = 40
        Case 2
            kolumna = 46
        Case 3
            kolumna = 56
        Case 4
            kolumna = 64
        Case 5
            kolumna = 73
    End Select
    
    balance_on_pop_data = Trim(sess0.screen.getString(7 + wiersz, kolumna, 8))
End Function

