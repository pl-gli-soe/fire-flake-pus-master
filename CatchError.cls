VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CatchError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private feedback_info As String


Public Function getFeedbackInfo() As String
    getFeedbackInfo = CStr(feedback_info)
End Function

Public Function noNSRissue(m As MGO) As Boolean
    If UCase(CStr(m.actualScreen)) <> UCase(CStr(m.pMS3P9800.screenName)) Then
        noNSRissue = True
    Else
        noNSRissue = False
        feedback_info = feedback_info & "MS3P9800 screen issue / nsr issue"
    End If
End Function

Public Function inputDataAvail(r As Range) As Boolean


    inputDataAvail = False

    If Not r Is Nothing Then
        inputDataAvail = True
    Else
        inputDataAvail = False
        feedback_info = feedback_info & "no input data!"
    End If
End Function

Public Function dataAvail(screen_name As String, infocode As String) As Boolean
    
    
    
    dataAvail = True
    If screen_name = "ms9pop00" Then
        ' I4266: DATA NOT FOUND FOR SCAN REQUEST - for example
        If infocode = "I4266" Then
            dataAvail = False
            feedback_info = feedback_info & "ms9pop00 data not available"
        End If

    ElseIf screen_name = "zk7pdrqm" Or screen_name = "zk7pwrqm" Then
    
        ' R6116: PART NUMBER NOT FOUND
        ' R6105: MUST INQUIRE BEFORE PF11 CAN BE USED
        If infocode = "R6116" Or infocode = "R6105" Then
            dataAvail = False
            feedback_info = feedback_info & "zk7pdrqm/zk7pwrqm data not available"
        End If
    End If
End Function


Public Function catchIfInDailyRqmsThereAreNoDates(strFromMgo As String) As Boolean
    If strFromMgo = "" Then
        catchIfInDailyRqmsThereAreNoDates = True
    Else
        catchIfInDailyRqmsThereAreNoDates = False
        feedback_info = feedback_info & "Daily Rqm Issue"
    End If
End Function

Public Function catchIfPartNotFoundInHistoryScreen(strFromMgo As String) As Boolean
    catchIfPartNotFoundInHistoryScreen = False
    
    
    If strFromMgo <> "I4017" Then
        catchIfPartNotFoundInHistoryScreen = True
    Else
        catchIfPartNotFoundInHistoryScreen = False
        feedback_info = feedback_info & "PART NOT FOUND IN MS9PH100"
    End If
End Function


Public Function catchIfMGOIsNotActive(sess0) As Boolean
    catchIfMGOIsNotActive = False
    
    If sess0 Is Nothing Then
        catchIfMGOIsNotActive = True
        feedback_info = feedback_info & "MGO is not active"
    End If
End Function

Private Sub Class_Initialize()
    feedback_info = ""
End Sub

Public Sub sth_went_wrong_with_mgo_on_ms9pop00_screen_download_rqm_attempt()
    feedback_info = feedback_info & "unrecognized failure on ms9pop00 during rqm downloading for special config"
End Sub
