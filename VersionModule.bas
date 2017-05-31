Attribute VB_Name = "VersionModule"
' Version Module




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



' no to start
' modul ten powstal od czwartej generacji ffa
' w mysl o tym czego wlasciwie potrzebuje uzytkownik musze pomyslec ile faktycznie jest narzedzi potrzebnych a ile nie
' odchudzenie ffa rowniez wchodzi w gre aby chodzil odrobine szybciej
Public Sub msgbox_about(ictrl As IRibbonControl)



    ' now in transit class we have new added lines: on error resume next
    
    ' ------------------------------------------------------------
    ' On Error Resume Next
    ' t.mDeliveryDate = CDate(m.convertToDateFromMS9POP00Date(m.pMS9POP00.transEDA(Int(x))))
    ' On Error Resume Next
    ' t.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
    ' t.mNotYetReceived = True
    ' ...
    ' On Error Resume Next
    ' t.mPickupDate = CDate(m.convertToDateFromMS9POP00Date(CStr(m.pMS9POP00.transSDATE(Int(x)))))
    ' ------------------------------------------------------------
    version_6_03 = "This is Fire Flake PM - the 6th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "fix on not yet recv implemneted into first end balance formula" & Chr(10) & _
        "VERSION 6.03" & Chr(10)
        
    MsgBox CStr(version_6_03)
    
    version_6_02 = "This is Fire Flake PM - the 6th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 6.02" & Chr(10)
    
    version_6_01 = "This is Fire Flake PM - the 6th generation of this tool" & Chr(10) & _
        "wstepne zmiany z FFL na FFPM" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 6.01" & Chr(10)
    
    version_0_01 = "This is Fire Flake PM - the 6th generation of this tool" & Chr(10) & _
        "" & Chr(10) & _
        "VERSION 6.01" & Chr(10)
    
    ' --------------------------
    
End Sub


