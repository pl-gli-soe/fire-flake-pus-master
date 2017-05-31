VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakePMScaffoldItem"
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



' description
' ------------------------------------------------------------------------------

' scaffold item rozni sie od zwyklego itemu tym ze item zwykly odnosi sie do
' elementu z mgo, natomiast item scaffoldowy odnosi calosciowo do layoutu
' buduje ilosci kolumn dla wszystkich pnow oraz ustawia odpowiednio formuly

' ------------------------------------------------------------------------------

Public pozycja As Integer
Public ilosc_pol_tranzytow As Integer
Public ile_kolumn As Integer

Private formula As String

Private pus_dictionary As Dictionary



' maly zarcik ;)
Public Property Get PUSSY() As Dictionary
    Set PUSSY = pus_dictionary
End Property


Public Sub dodanie_manuala()

    
    If Not pus_dictionary.Exists(ff_pm.STR_MANUAL) Then
        pus_dictionary.Add ff_pm.STR_MANUAL, 1
    Else
        pus_dictionary(ff_pm.STR_MANUAL) = pus_dictionary(ff_pm.STR_MANUAL) + 1
    End If
End Sub



Public Sub proba_dodania_pusa_jako_nowej_labelki(it As ITransit)
    
    If Not pus_dictionary.Exists(CStr(it.mName)) Then
        pus_dictionary.Add CStr(it.mName), 1
    Else
        pus_dictionary(CStr(it.mName)) = pus_dictionary(CStr(it.mName)) + 1
    End If
End Sub


Private Sub make_std_formula()

    'for_first_time_formula = "=" & CStr(ofst_CBAL) & "-" & CStr(ofst_MISC) & "-RC[-3]+RC[-4]+RC[-1]-RC[-2]"
    'for_regular_formula = "=RC[-3]+RC[-1]-RC[-2]"
    
    If ilosc_pol_tranzytow > 0 Then
        inner_sit = ""
        For x = 1 To ilosc_pol_tranzytow
            inner_sit = inner_sit & "+RC[" & CStr(Int(-1 * x)) & "]"
        Next x
        
        
        ' calc OK
        inner_srqm = "-RC[" & CStr(Int(-1 * ilosc_pol_tranzytow - 1)) & "]"
        inner_sebal = "+RC[" & CStr(Int(-1 * ilosc_pol_tranzytow - 2)) & "]"
        
        formula = CStr(inner_srqm) & CStr(inner_sit) & CStr(inner_sebal)
    ElseIf ilosc_pol_tranzytow = 0 Then
        
        formula = "-RC[-1]+RC[-2]"
    End If
    
End Sub


Private Sub make_first_formula()

    'for_first_time_formula = "=" & CStr(ofst_CBAL) & "-" & CStr(ofst_MISC) & "-RC[-3]+RC[-4]+RC[-1]-RC[-2]"
    'for_regular_formula = "=RC[-3]+RC[-1]-RC[-2]"
    
    If ilosc_pol_tranzytow > 0 Then
        inner_sit = ""
        For x = 1 To ilosc_pol_tranzytow
            inner_sit = inner_sit & "+RC[" & CStr(Int(-1 * x)) & "]"
        Next x
        
        
        ' calc OK
        inner_srqm = "-RC[" & CStr(Int(-1 * ilosc_pol_tranzytow - 1)) & "]"
        inner_past_due_rqm = "-RC[" & CStr(Int(-1 * ilosc_pol_tranzytow - 2)) & "]"
        inner_nyrecv = "+RC[" & CStr(Int(-1 * ilosc_pol_tranzytow - 3)) & "]"
        
        formula = CStr(inner_srqm) & CStr(inner_sit) & CStr(inner_past_due_rqm) & CStr(inner_nyrecv)
    ElseIf ilosc_pol_tranzytow = 0 Then
        
        formula = "-RC[-1]-RC[-2]+RC[-3]"
    End If
    
End Sub


Public Sub inner_calc()
    
    ' 2 : RQM I EBAL
    
    ' uproszczenie koncepcji
    
    
    
    ilosc_pol_tranzytow = pus_dictionary.COUNT
    
    If pozycja = 1 Then
        make_first_formula
    Else
        make_std_formula
    End If
    
    ile_kolumn = 2 + ilosc_pol_tranzytow
End Sub

Public Function get_formula() As String
    
    get_formula = CStr(formula)
End Function

Private Sub Class_Initialize()
    Set pus_dictionary = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set pus_dictionary = Nothing
End Sub
