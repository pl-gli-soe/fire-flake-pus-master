VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WizardForm 
   Caption         =   "Quick Wizard"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   OleObjectBlob   =   "WizardForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WizardForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wizard_handler As WizardHandler

Private Sub Btncancel_Click()
    End
End Sub

Public Sub connectWithWizardHandler(wh As WizardHandler)
    Set wizard_handler = wh
End Sub


Private Sub BtnSubmit_Click()
    hide
    submit_changes_and_change_register
End Sub

Private Sub submit_changes_and_change_register()

    With wizard_handler

        .setPltType Me.ComboBoxType.Value
    
    
        Dim i As Range, czy_jest_taki_plt As Boolean
        For Each i In .getDrh.getPlants()
            If CStr(i) = CStr(.plt) Then
                czy_jest_taki_plt = True
                
            End If
        Next i
        
        If Not czy_jest_taki_plt Then
            ' przelecielismy wszystkie plty i nic zatem trzeba dodac nowy
            .getDrh.getPlants().End(xlDown).Offset(1, 0).Value = .plt
            .getDrh.getPlants().End(xlDown).Offset(0, 1).Value = .getPltType
            
            wizard_handler.getCurrPn.Offset(0, 5) = CStr(wizard_handler.getNewSOffset0_5())
        End If
    End With
    

End Sub

