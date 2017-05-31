VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DownloadingRulesForm 
   Caption         =   "Downloading Rules Form"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6315
   OleObjectBlob   =   "DownloadingRulesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DownloadingRulesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Global Const OFFSET_FOR_NEW_PLT = 20
'Global Const INIT_RULES_FORM_WIDTH = 320
'Global Const INIT_RULES_FORM_HEIGHT = 108
'
'Global Const INIT_RULES_FORM_BTN_TOP = 54
'
'Global Const TEXTBOX_PLT_0_LEFT = 12
'Global Const TEXTBOX_RQM_0_LEFT = 54
'
'Global Const TEXTBOX_PLT_0_TOP = 30
'Global Const TEXTBOX_RQM_0_TOP = 30
'
'Global Const TEXTBOX_PLT_0_W = 36
'Global Const TEXTBOX_RQM_0_W = 72
'
'Global Const TEXTBOX_PLT_0_H = 18
'Global Const TEXTBOX_RQM_0_H = 18
'
'Global Const LABEL_CMNT_0_L = 132
'Global Const LABEL_CMNT_0_T = 30
'Global Const LABEL_CMNT_0_W = 174
'Global Const LABEL_CMNT_0_H = 18



Private Sub BtnAdd_Click()

    ' MsgBox drh.whereAmI
    
    drh.whereAmI = drh.whereAmI + 1
    drh.ile_bedzie_plantow = drh.whereAmI
    drh.okresl_wielkosc_forma_i_przesun_guziki_w_dol
    drh.iteracja Int(drh.whereAmI)
    
    Me.Repaint
    

End Sub

Private Sub BTNHide_Click()
    Me.hide
End Sub

Private Sub BtnRemove_Click()
    ' MsgBox drh.whereAmI
    Me.Controls.Remove "Kombo" & CStr(drh.whereAmI)
    Me.Controls.Remove "Label" & CStr(drh.whereAmI)
    Me.Controls.Remove "Name" & CStr(drh.whereAmI)
    
    
    ThisWorkbook.Sheets("register").Cells(Int(drh.whereAmI) + 2, CONFIG_REG_PLT_COLUMN + 1).Clear
    ThisWorkbook.Sheets("register").Cells(Int(drh.whereAmI) + 2, CONFIG_REG_PLT_COLUMN).Clear
    
    drh.whereAmI = drh.whereAmI - 1
    drh.ile_bedzie_plantow = drh.whereAmI
    drh.okresl_wielkosc_forma_i_przesun_guziki_w_dol
    Me.Repaint
End Sub
