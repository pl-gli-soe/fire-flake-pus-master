VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TheLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta ma pomoc w budowie layoutu
' nadanie standardowych struktur na arkuszu
Public Sub initLayout()
    ' Sheets("daily").Activate
    ' Sheets("register").Range("lastRow") = last_row() + 4
    ' Sheets("register").Range("allParts") = last_row() - 5
    Sheets.Add Before:=Worksheets(Worksheets.COUNT)
    ' If IsMissing(for_template) Then
        ' ActiveSheet.NAME = "daily " & CStr(Sheets.COUNT)
    Sheets("register").Range("sheetName") = ActiveSheet.Name
    'ElseIf for_template = True Then
    '    ActiveSheet.NAME = "template " & CStr(Sheets.COUNT)
    '    Sheets("register").Range("sheetName") = ActiveSheet.NAME
    'End If
    
    
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
    ' ILayout_ClearLayoutWithData
    ' ILayout_GenerujTabeleDlaListyDanychOgolnych
    ' ILayout_GenerujReszteRusztownia
End Sub

Public Sub mergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
    Range(CStr(chrx(mfirstcolumn)) & CStr(mrow) & ":" & CStr(chrx(mfirstcolumn + 2)) & CStr(mrow)).Merge
End Sub

Public Sub mergeXCellsForDate(mrow As Integer, mfirstcolumn As Integer, x As Integer)
    Range(CStr(chrx(mfirstcolumn)) & CStr(mrow) & ":" & CStr(chrx(mfirstcolumn + (x - 1))) & CStr(mrow)).Merge
End Sub

Public Sub backColor(rng As Range, kolor As Long)
    With rng.Interior
        .Pattern = xlSolid
        .Color = kolor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Public Sub changeTxtOrientation(rng As Range, angle As Integer)

    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = Int(angle)
        
    End With
End Sub

Public Sub fillSolidFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillRightThinLine(rng As Range, kolor As Long)


    'rng.Borders(xlDiagonalDown).LineStyle = xlNone
    'rng.Borders(xlDiagonalUp).LineStyle = xlNone


    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

Public Sub fillBottomThinLine(rng As Range, kolor As Long)


    'rng.Borders(xlDiagonalDown).LineStyle = xlNone
    'rng.Borders(xlDiagonalUp).LineStyle = xlNone


    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

Public Sub fillSolidGridLines(rng As Range, kolor As Long)
    
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Public Sub fillThinFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub fillThinGridLines(rng As Range, kolor As Long)

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

