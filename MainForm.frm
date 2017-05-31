VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Main Form"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClose_Click()
    Application.ThisWorkbook.Close
End Sub

Private Sub BtnDownloadingRules_Click()
    
    ' DownloadingRulesForm
    
    
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
    
    
    
    
    drh.inicjacja
    drh.ile_bedzie_plantow = CLng(drh.getPlants().COUNT)
    drh.okresl_wielkosc_forma_i_przesun_guziki_w_dol
    

    For x = 1 To drh.ile_bedzie_plantow
    
    
    
        drh.iteracja Int(x)
        
        
    Next x
    
    'tb.Left = TEXTBOX_PLT_0_LEFT
    'tb.Top = TEXTBOX_PLT_0_TOP + OFFSET_FOR_NEW_PLT
    
    
    ' wlasciwie show bedzie na samym koncu
    DownloadingRulesForm.show
End Sub

Private Sub BTNHide_Click()
    Me.hide
End Sub

Private Sub BtnMoreLess_Click()
    If Me.BtnMoreLess.Caption Like "*More*" Then
        Me.Height = 500
        Me.BtnMoreLess.Caption = "Less"
    Else
        Me.Height = 110
        Me.BtnMoreLess.Caption = "More"
    End If
        
End Sub

Private Sub BtnMoveAllToLeft_Click()

    fillAllPopsDataByChar "x"
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init

End Sub

Private Sub BtnMoveAllToRight_Click()


    fillAllPopsDataByChar ""
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub fillAllPopsDataByChar(ch As String)
    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
    
    Do
        If CLng(r.Interior.Color) <> CLng(ThisWorkbook.Sheets("register").Range("black")) Then ' as black
            r.Offset(0, 1) = ch
        End If
        Set r = r.Offset(1, 0)
    Loop While r <> ""
End Sub

Private Sub btnMoveToLeft_Click()
    change_register_workhseet "x"
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub BtnMoveToRight_Click()
    change_register_workhseet ""
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
End Sub

Private Sub BtnRunDaily_Click()


    Application.EnableEvents = False
    Me.hide
    
    Dim wybor_typu_layoutu As LAYOUT_TYPE
    
    wybor_typu_layoutu = LIST_LAYOUT
    
    If Me.DTPickerPUSLimit.Enabled Then
        ThisWorkbook.Sheets("register").Range("pusLimit") = CDate(Me.DTPickerPUSLimit.Value)
    Else
        ThisWorkbook.Sheets("register").Range("pusLimit") = CDate(Me.DTPickerPUSLimit.Value) + 100
    End If
    
    If Me.DTPickerRQMLimit.Enabled Then
        ThisWorkbook.Sheets("register").Range("rqmLimit") = CDate(Me.DTPickerRQMLimit.Value)
    Else
        ThisWorkbook.Sheets("register").Range("rqmLimit") = CDate(Me.DTPickerRQMLimit.Value) + 100
    End If
    
    
    ' wydaje mi sie ze brakuje tutaj synchro ilosci dni historii
    ' ------------------------------------------------------------
     ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0") = CLng(Me.ComboBoxHistoryLimit.Value)
    ' ------------------------------------------------------------
    
    
    ThisWorkbook.Sheets("register").Range("LAYOUT_TYPE") = wybor_typu_layoutu
    ThisWorkbook.Sheets("register").Range("RUN_TYPE") = DAILY
    ThisWorkbook.Sheets("register").Range("START_TYPE") = FROM_THE_BEGINNING
    
    runReport DAILY, wybor_typu_layoutu, FROM_THE_BEGINNING, CDate(ThisWorkbook.Sheets("register").Range("pusLimit")), CDate(ThisWorkbook.Sheets("register").Range("rqmLimit"))
    Application.EnableEvents = True
End Sub

Private Sub CheckBoxPUSLimit_Click()
    If Not Me.CheckBoxPUSLimit.Value Then
        Me.DTPickerPUSLimit.Enabled = False
    Else
        Me.DTPickerPUSLimit.Enabled = True
    End If
End Sub

Private Sub CheckBoxRQMLimit_Click()

    If Not Me.CheckBoxRQMLimit.Value Then
        Me.DTPickerRQMLimit.Enabled = False
    Else
        Me.DTPickerRQMLimit.Enabled = True
    End If
End Sub

Private Sub CheckBoxWeekNum_Click()
    With Me.CheckBoxWeekNum
        If .Value = True Then
            ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 1
        Else
            ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 0
        End If
    End With
End Sub

Private Sub CommandButton1_Click()
    Me.hide
    MsgBox "not yet implemented"
    Me.show
End Sub

Private Sub change_register_workhseet(s As String)

    Dim r As Range

    If s = "" Then
        
        For x = 0 To Me.ListBoxInCellLeft.ListCount - 1
            If Me.ListBoxInCellLeft.Selected(x) Then
                tmp = Me.ListBoxInCellLeft.List(x)
                
                
                
                Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
                
                Do
                    If CLng(r.Interior.Color) <> CLng(ThisWorkbook.Sheets("register").Range("black")) Then ' as black
                        If CStr(tmp) = CStr(r) Then
                            r.Offset(0, 1) = s
                        End If
                    End If
                    Set r = r.Offset(1, 0)
                Loop While r <> ""
                
            End If
        Next x
    ElseIf s = "x" Then
    
    
        For x = 0 To Me.ListBoxInCommentRight.ListCount - 1
            If Me.ListBoxInCommentRight.Selected(x) Then
                tmp = Me.ListBoxInCommentRight.List(x)
                
                
                
                Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
                
                Do
                    If CLng(r.Interior.Color) <> CLng(ThisWorkbook.Sheets("register").Range("black")) Then ' as black
                        If CStr(tmp) = CStr(r) Then
                            r.Offset(0, 1) = s
                        End If
                    End If
                    Set r = r.Offset(1, 0)
                Loop While r <> ""
                
            End If
        Next x
    End If
End Sub




Private Sub ComboBoxColorLayout_Change()


    ThisWorkbook.Sheets("register").Range("actualColorLayoutChoice") = Me.ComboBoxColorLayout.Value
    
    Set item_on_layout_color_list = ThisWorkbook.Sheets("register").Range("m10")
    Do
        If Trim(CStr(item_on_layout_color_list)) = Trim(CStr(ThisWorkbook.Sheets("register").Range("actualColorLayoutChoice"))) Then
            ' jestesmy w odpowiednim miejscu zeby zmienic aktualne kolory layoutu
            ThisWorkbook.Sheets("register").Range("primary").Interior.Color = item_on_layout_color_list.Offset(0, 1).Interior.Color
            ThisWorkbook.Sheets("register").Range("secondary").Interior.Color = item_on_layout_color_list.Offset(0, 2).Interior.Color
            ThisWorkbook.Sheets("register").Range("weekendColor").Interior.Color = item_on_layout_color_list.Offset(0, 3).Interior.Color
            
            Me.TextBoxPrimaryColor.backColor = ThisWorkbook.Sheets("register").Range("primary").Interior.Color
            Me.TextBoxSecondaryColor.backColor = ThisWorkbook.Sheets("register").Range("secondary").Interior.Color
            Me.TextBoxWeekendColor.backColor = ThisWorkbook.Sheets("register").Range("weekendColor").Interior.Color
            Me.TextBoxMinusColor.backColor = ThisWorkbook.Sheets("register").Range("minus").Interior.Color
            Me.TextBoxWarningColor.backColor = ThisWorkbook.Sheets("register").Range("warning").Interior.Color
        End If
        
        Set item_on_layout_color_list = item_on_layout_color_list.Offset(1, 0)
    Loop Until Trim(item_on_layout_color_list) = ""
    
End Sub

Private Sub ComboBoxHistoryLimit_Change()
    ThisWorkbook.Sheets("register").Range("HOW_MANY_DAYS_FOR_PPUS0") = _
        Me.ComboBoxHistoryLimit.Value
End Sub

Private Sub UserForm_Initialize()
    
    
    
    
    Set drh = New DownloadingRulesHandler

    ' dates now
    Me.DTPickerPUSLimit = Now
    Me.DTPickerRQMLimit = Now
    
    Me.Height = 110
    Me.Width = 365



    ' week #
    Me.CheckBoxWeekNum.Value = True
    ThisWorkbook.Sheets("register").Range("weekNumOnTop") = 1

    


    ' history limit
    ' ============================================
    Me.ComboBoxHistoryLimit.Clear
    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("BegOfHistoryLimitRange")
    
    Do
        Me.ComboBoxHistoryLimit.AddItem CStr(r)
        Set r = r.Offset(1, 0)
    Loop While r <> ""
    
    Me.ComboBoxHistoryLimit = assign_default_value()
    ' ============================================
    


    ' limitacje
    Me.DTPickerPUSLimit.Enabled = False
    Me.DTPickerRQMLimit.Enabled = False
    Me.CheckBoxPUSLimit.Value = False
    Me.CheckBoxRQMLimit.Value = False
    
    
    
    ' tutaj zabawa z konfiguracja danych z popa
    set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init
    
    
    
    
    ' teraz zabawa z kolorami
    Me.ComboBoxColorLayout.Clear
    Dim item_on_layout_color_list As Range
    Set item_on_layout_color_list = ThisWorkbook.Sheets("register").Range("m10")
    Do
    
        Me.ComboBoxColorLayout.AddItem CStr(item_on_layout_color_list)
        Set item_on_layout_color_list = item_on_layout_color_list.Offset(1, 0)
    Loop Until Trim(item_on_layout_color_list) = ""
    
    Me.ComboBoxColorLayout.Value = CStr(ThisWorkbook.Sheets("register").Range("actualColorLayoutChoice"))
    
    Set item_on_layout_color_list = ThisWorkbook.Sheets("register").Range("m10")
    Do
        If Trim(CStr(item_on_layout_color_list)) = Trim(CStr(ThisWorkbook.Sheets("register").Range("actualColorLayoutChoice"))) Then
            ' jestesmy w odpowiednim miejscu zeby zmienic aktualne kolory layoutu
            ThisWorkbook.Sheets("register").Range("primary").Interior.Color = item_on_layout_color_list.Offset(0, 1).Interior.Color
            ThisWorkbook.Sheets("register").Range("secondary").Interior.Color = item_on_layout_color_list.Offset(0, 2).Interior.Color
            ThisWorkbook.Sheets("register").Range("weekendColor").Interior.Color = item_on_layout_color_list.Offset(0, 3).Interior.Color
            
            Me.TextBoxPrimaryColor.backColor = ThisWorkbook.Sheets("register").Range("primary").Interior.Color
            Me.TextBoxSecondaryColor.backColor = ThisWorkbook.Sheets("register").Range("secondary").Interior.Color
            Me.TextBoxWeekendColor.backColor = ThisWorkbook.Sheets("register").Range("weekendColor").Interior.Color
            Me.TextBoxMinusColor.backColor = ThisWorkbook.Sheets("register").Range("minus").Interior.Color
            Me.TextBoxWarningColor.backColor = ThisWorkbook.Sheets("register").Range("warning").Interior.Color
        End If
        
        Set item_on_layout_color_list = item_on_layout_color_list.Offset(1, 0)
    Loop Until Trim(item_on_layout_color_list) = ""
    
End Sub

Private Function assign_default_value()

    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("BegOfHistoryLimitRange")
    
    Do
        If r.Offset(0, -1).Value = "default" Then
            assign_default_value = r
            Exit Function
        End If
        Set r = r.Offset(1, 0)
    Loop While r <> ""

End Function


Private Sub set_pop_data_left_right_thing_take_data_from_regiser_worksheet_on_init()

    Me.ListBoxInCellLeft.Clear
    Me.ListBoxInCommentRight.Clear

    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Range("begOfPopParams")
    
    Do
    
        If r.Offset(0, 1) = "x" Then
            Me.ListBoxInCellLeft.AddItem CStr(r)
        Else
            Me.ListBoxInCommentRight.AddItem CStr(r)
        End If
        Set r = r.Offset(1, 0)
    Loop While r <> ""
End Sub

