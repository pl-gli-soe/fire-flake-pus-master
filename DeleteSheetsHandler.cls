VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteSheetsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub deleteCurrentSheet()
    Application.DisplayAlerts = False
    
    If (ActiveSheet.Name Like "*input*") Or (ActiveSheet.Name Like "*register*") Then
        MsgBox "you can't delete this sheet!"
    Else
        ActiveSheet.Delete
    End If
    
    Application.DisplayAlerts = True
End Sub

Public Sub deleteAllSheets()
    ret = MsgBox("Delete all?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Application.DisplayAlerts = False
        
        
        Dim sh As Worksheet
        x = 1
        Do
            If (Sheets(x).Name Like "*input*") Or (Sheets(x).Name Like "*register*") Then
                x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.COUNT
        Application.DisplayAlerts = True
    End If
End Sub
