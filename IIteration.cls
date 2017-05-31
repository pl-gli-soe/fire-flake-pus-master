VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function getItem() As IItem

End Function

Public Sub setItem(i As IItem)

End Sub

Public Function getItemsCollection() As Collection

End Function

Public Sub setItemsCollections(c As Collection)

End Sub

Public Function getCommonData() As CommonData

End Function

Public Sub setCommonData(cd As CommonData)

End Sub

Public Sub czesciowyFillNaTejIteracji(ByRef r As Range, lt As LAYOUT_TYPE, st As START_TYPE, limit As Date, rqm_limit As Date, rowFlag)

End Sub

Public Sub createFullIteration(m As MGO)

End Sub


Public Sub addColorFlavour()

End Sub

Public Function getLayout()
    
End Function

Public Sub connectFireFlakeLayout(ffll As IFireFlakeLayout)
End Sub
