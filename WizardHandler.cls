VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WizardHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public plt As String
Public ic As ITERATION_CONFIG

Private cfgnum As Integer ' <- wartosc ta jest jednoczesnie wartoscia ktora znajduje sie w s.offset(0,5)
Private plt_type As String
Private mdrh As DownloadingRulesHandler

Private curr_pn As Range


Public Function getCurrPn() As Range
    Set getCurrPn = curr_pn
End Function

Public Sub setCurrPn(arg As Range)
    Set curr_pn = arg
End Sub



' te dwie zmienne beda manipulowane tylko z punktu widzenia metod setplt i setplttype
' ktore dodatkowo moga byc uzyte tylko w srodku implementacji wizardform podczas uruchmienia metody submit
' inaczej pozostana puste i samo uzycie metod zwaracajacych:
' Public Function getNewIterConfing() As String
' Public Function getNewSOffset0_5() As String
' zwroci dalej stare wartosci ujemny - tak jakby brak zmiany :D

Public Function getDrh() As DownloadingRulesHandler
    Set getDrh = mdrh
End Function

Public Sub setPltType(pt As String)
    plt_type = pt
End Sub

Public Function getPltType() As String
    getPltType = plt_type
End Function

Public Sub submit_changes_via_form(ByRef drh As DownloadingRulesHandler, ByRef s As Range, cfg_num As Integer)

    ' plt is public so it's visible and can be assigned quicker!
    ' assign to field which will be available on whole scope of this type object
    ' cfg_num is just local var
    cfgnum = cfg_num
    Set mdrh = drh
    
    Set curr_pn = s
    
    With WizardForm
    
        .connectWithWizardHandler Me
    
        .LabelPlt.Caption = plt
        
        If cfg_num = -1 Then
            .LabelDesc.Caption = ""
        ElseIf cfg_num = -2 Then
            .LabelDesc.Caption = ""
        Else
            MsgBox "adding new plt definition in wizardhandler object implementation went wrong!"
            End
        End If
        .ComboBoxType.Clear
        
        Dim i As Range
        For Each i In drh.getRqmSource()
            .ComboBoxType.AddItem CStr(i)
        Next i
        .show
    End With
    
End Sub



Public Function getNewIterConfing() As String



    
    getNewIterConfing = CStr(cfgnum)
    
    If plt <> "" And plt_type <> "" Then
    
        Dim i As Range
        For Each i In mdrh.getRqmSource()
            If CStr(i) = CStr(plt_type) Then
                getNewIterConfing = i.Offset(0, 2)
            End If
        Next i
    
    End If
End Function
        
Public Function getNewSOffset0_5() As String

    

    getNewSOffset0_5 = CStr(cfgnum)
    
    If plt <> "" And plt_type <> "" Then
    
    
        Dim i As Range
        For Each i In mdrh.getRqmSource()
            If CStr(i) = CStr(plt_type) Then
                getNewSOffset0_5 = i.Offset(0, 2)
            End If
        Next i
    
    End If
End Function
