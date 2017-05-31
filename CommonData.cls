VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public plt As String
Public pn As String
Public comment As String
Public additionalCommnet As String

' component
Public c As String


' some pop data
Public duns As String
Public pcsToGo As Long
Public doh As Double

' dates
Public firstRunout As Date
Public firstWarning As Date

' limitations
Public pusLimit As Date
Public rqmLimit As Date

' pop - the rest
Public stdPack As Long
Public bbal As Long
Public cbal As Long
Public ttime As Long
Public bank As Long
Public os As Long
Public misc As Long
Public qhd As Long
Public iter_config As ITERATION_CONFIG
Public runType As RUN_TYPE
Public layoutType As LAYOUT_TYPE
Public startType As START_TYPE


' dane z popa ktore jeszcze nie byly brane pod uwage z perspektywy tego typu obiektu
Public fupCode As String
Public fmaFupCode As String
Public supplierName As String
Public mode As String
Public partName As String

' comment from mgo pop
Public count_cmnt As String
Public o_cmnt As String
Public f_cmnt As String







' wiersz na ktorym znajduje sie iteracja
Public rowFlag As Long


' error log feedback
Public errorLog As String


