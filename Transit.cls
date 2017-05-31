VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ITransit



' tutaj komentarz maly poniewaz boolean posiada tylko dwa stany
' stanem domyslnym w ff 3 generacji jest not yet received jako true
' jednak moze to powodowac pewne dziwne rzeczy w stylu ogromnych ilosci pusow czy tez asnow
' z drugiej strony mozna przeciez traktowac wszystkie asny przyszlosci jako not booked yet
' i nie wyklucza to logiki :D
' a wiec zostaniemy przy takim samym rozwiazaniu jakie mialo miejsce w ffach trzeciej generacji
Private not_yet_received_flag As Boolean

' ta zmienna ma podchwytliwa nazwe bo bardziej chodzi o po400 czyli asnypusy ktore sa w tranzycie jeszcze
' pomaga w tym w odniesieniu sie do tranzytow zduplikowanych
Private on_pop As Boolean


' to teraz mysl moja dazy do tego zeby z tych dat
' ponieaz data moze przechowywac i date sama w sobie jak i czas
' dzieki czemu mozemy ograniczyc ilosc zmiennych
Private pickup_date As Date
Private delivery_date As Date

Private pickup_time As Date
Private delivery_time As Date

Private qty As Long

Private route As String
Private st As String
' trlr - tj. container rozwijajac skrot
Private cont As String

' moj name dla kazdego tranzytu
Private nm As String

Private Property Let ITransit_mDeliveryDate(RHS As Date)
    delivery_date = RHS
End Property

Private Property Get ITransit_mDeliveryDate() As Date
    ITransit_mDeliveryDate = delivery_date
End Property

Private Property Let ITransit_mDeliveryTime(RHS As Date)
    delivery_time = RHS
End Property

Private Property Get ITransit_mDeliveryTime() As Date
    ITransit_mDeliveryTime = delivery_time
End Property



' also you can say that this is SID
Private Property Let ITransit_mName(RHS As String)
    nm = RHS
End Property

Private Property Get ITransit_mName() As String
    ITransit_mName = nm
End Property

Private Property Let ITransit_mNotYetReceived(RHS As Boolean)
    not_yet_received_flag = RHS
End Property

Private Property Get ITransit_mNotYetReceived() As Boolean
    ITransit_mNotYetReceived = not_yet_received_flag
End Property

Private Property Let ITransit_mOnPo400(RHS As Boolean)
    on_pop = RHS
End Property

Private Property Get ITransit_mOnPo400() As Boolean
    ITransit_mOnPop = on_pop
End Property

Private Property Let ITransit_mPickupDate(RHS As Date)
    pickup_date = CDate(RHS)
End Property

Private Property Get ITransit_mPickupDate() As Date
    ITransit_mPickupDate = CDate(pickup_date)
End Property

Private Property Let ITransit_mPickupTime(RHS As Date)
    pickup_time = CDate(RHS)
End Property

Private Property Get ITransit_mPickupTime() As Date
    ITransit_mPickupTime = pickup_time
End Property

Private Property Let ITransit_mQty(RHS As Long)
    qty = CLng(RHS)
End Property

Private Property Get ITransit_mQty() As Long
    ITransit_mQty = CLng(qty)
End Property

Private Property Let ITransit_mRoute(RHS As String)
    route = CStr(RHS)
End Property

Private Property Get ITransit_mRoute() As String
    ITransit_mRoute = CStr(route)
End Property

Private Property Let ITransit_mST(RHS As String)
    st = CStr(RHS)
End Property

Private Property Get ITransit_mST() As String
    ITransit_mST = CStr(st)
End Property

Private Property Let ITransit_mTRLR(RHS As String)
    cont = CStr(RHS)
End Property

Private Property Get ITransit_mTRLR() As String
    ITransit_mTRLR = CStr(cont)
End Property
