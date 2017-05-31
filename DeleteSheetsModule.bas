Attribute VB_Name = "DeleteSheetsModule"
Public Sub del_curr_sh(ictrl As IRibbonControl)
    Dim dh As DeleteSheetsHandler
    Set dh = New DeleteSheetsHandler
    dh.deleteCurrentSheet
End Sub

Public Sub del_all_shs(ictrl As IRibbonControl)
    Dim dh As DeleteSheetsHandler
    Set dh = New DeleteSheetsHandler
    dh.deleteAllSheets
End Sub
