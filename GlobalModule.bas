Attribute VB_Name = "GlobalModule"
#If VBA7 Then
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As LongPtr
#Else
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
#End If


'r = "Requirements"
'r.Offset(0, 1) = "In Transit"
'r.Offset(0, 2) = "Ending Balance"
Global Const STR_RQM = "Requirements"
Global Const STR_IT = "In Transit"
Global Const STR_EBAL = "Ending Balance"

Global Const STR_MANUAL = "Manual"

Global Const OFFSET_FOR_NEW_PLT = 20
Global Const INIT_RULES_FORM_WIDTH = 320
Global Const INIT_RULES_FORM_HEIGHT = 88

Global Const INIT_RULES_FORM_BTN_TOP = 36

Global Const TEXTBOX_PLT_0_LEFT = 12
Global Const TEXTBOX_RQM_0_LEFT = 54

Global Const TEXTBOX_PLT_0_TOP = 10
Global Const TEXTBOX_RQM_0_TOP = 10

Global Const TEXTBOX_PLT_0_W = 36
Global Const TEXTBOX_RQM_0_W = 72

Global Const TEXTBOX_PLT_0_H = 18
Global Const TEXTBOX_RQM_0_H = 18

Global Const LABEL_CMNT_0_L = 132
Global Const LABEL_CMNT_0_T = 10
Global Const LABEL_CMNT_0_W = 174
Global Const LABEL_CMNT_0_H = 18

Global Const CONFIG_REG_PLT_COLUMN = 18

Global Const DEFAULT_ZERO_RQMS = 10

Global Const EXTRA_DAYS_FOR_HISTORY = 20

Global Const GLOBAL_COLUMN_WIDTH = 4




' delay time
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' pobierz dynamiczna library
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)


Global Const LETTERS = 26
Global Const MAX_COLUMNS = 16384 ' ostatnia kolumna
Global Const C_HOUR = (0.041 + 0.001 * (2 / 3))
Global Const INITIAL_TIMING_FOR_ONE_PN = 6

Global sh As StatusHandler
Global drh As DownloadingRulesHandler

Public Enum ENUM_LEFT_RIGHT_LISTBOX
    MOVE_TO_LEFT_LISTBOX
    MOVE_TO_RGHT_LISTBOX
End Enum


Public Enum RUN_TYPE
    DAILY
    WEEKLY
    HOURLY
End Enum

Public Enum LAYOUT_TYPE
    LIST_LAYOUT
    COV_LAYOUT
    BOX_LAYOUT
End Enum

Public Enum START_TYPE
    FROM_THE_BEGINNING
    CONTINUE_BROKEN_ONE
End Enum


Public Enum ITERATION_CONFIG
    CONFIG_ASM
    CONFIG_POP
    CONFIG_M
    CONFIG_NULL
    CONFIG_Z
End Enum

Public Enum COMMENT_TYPE
    IN_TRANSIT
    DATA_FROM_POP
End Enum


Public Function MGO_active(m As MGO) As Boolean


    MGO_active = False
    
    If m Is Nothing Then
        MGO_active = False
        MsgBox "mgo class is nothing!"
        Exit Function
    End If
    
    
    
    If m.actualScreen <> "" Then
        MGO_active = True
    End If
End Function


Public Function chrx(col As Integer, Optional ByRef s As box) As String

    If col <= MAX_COLUMNS And col > 0 Then
    If s Is Nothing Then
        Set s = New box
    End If
    
    If col > LETTERS Then
        s.counter = s.counter + 1
        If s.counter = 26 Then
        ' wersja prostsza
            s.counter = 0
            s.scope = s.scope + 1
        End If
        chrx = chrx(col - LETTERS, s)
    Else
        If s.counter = 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + s.counter) + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope <> 0 Then
            chrx = chrx + Chr(64 + s.scope) + Chr(64 + s.counter) + Chr(64 + col)
        End If
    End If
    Else
        MsgBox "out of scope mf! MAX_COLUMNS = 16384"
    End If
    
   
End Function


Public Sub refresh_register_worksheet()

End Sub






