VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PackNumber As String
Private m_PackEvent As String
Private m_Chart As String
Private m_PatName As String
Private m_UserName As String
Private m_RecordDateTime As String
Private m_GroupRH As String
Private m_Supplier As String
Private m_DateExpiry As String
Private m_SampleID As String
Private m_crt As Boolean
Private m_cco As Boolean
Private m_cen As Boolean
Private m_crtr As Boolean
Private m_ccor As Boolean
Private m_cenr As Boolean
Private m_Barcode As String
Private m_Checked As Boolean
Private m_Notes As String
Private m_EventStart As String
Private m_EventEnd As String
Private m_OrderNumber As String
Private m_Screen As String
Private m_Reason As String
Private m_ISBT128 As String


Public Sub Save()

      Dim sql As String
      Dim Reason As String
      Dim UID As String
      Dim Notes As String
      Dim EvStart As String
      Dim EvEnd As String

10    On Error GoTo Save_Error

20    Reason = AddTicks(m_Reason)
30    UID = AddTicks(m_UserName)
40    Notes = AddTicks(m_Notes)

50    If IsDate(m_EventStart) Then
60        EvStart = Format$(m_EventStart, "'dd/MMM/yyyy HH:nn:ss'")
70    Else
80        EvStart = "null"
90    End If
100   If IsDate(m_EventEnd) Then
110       EvEnd = Format$(m_EventEnd, "'dd/MMM/yyyy HH:nn:ss'")
120   Else
130       EvEnd = "null"
140   End If

150   sql = "IF EXISTS (SELECT * FROM Latest WHERE " & _
            "           ISBT128 = '" & m_ISBT128 & "' " & _
            "           AND BarCode  = '" & m_Barcode & "' ) "
160   sql = sql & "  UPDATE Latest " & _
            "  SET ISBT128 = '" & m_ISBT128 & "', " & _
            "  Number = '" & m_PackNumber & "', " & _
            "  [Event] = '" & m_PackEvent & "', " & _
            "  PatID = '" & m_Chart & "', " & _
            "  PatName = '" & AddTicks(m_PatName) & "', " & _
            "  [Operator] = '" & UID & "', " & _
            "  [DateTime] = '" & Format$(Now, "dd/MMM/yyyy HH:nn:ss") & "', " & _
            "  GroupRh = '" & m_GroupRH & "', " & _
            "  Supplier = '" & m_Supplier & "', " & _
            "  LabNumber = '" & m_SampleID & "', " & _
            "  crt = '" & IIf(m_crt, 1, 0) & "', " & _
            "  cco = '" & IIf(m_cco, 1, 0) & "', " & _
            "  cen = '" & IIf(m_cen, 1, 0) & "', " & _
            "  crtr = '" & IIf(m_crtr, 1, 0) & "', " & _
            "  ccor = '" & IIf(m_ccor, 1, 0) & "', " & _
            "  cenr = '" & IIf(m_cenr, 1, 0) & "', " & _
            "  Checked = '" & IIf(m_Checked, 1, 0) & "', " & _
            "  Notes = '" & Notes & "', " & _
            "  EventStart = " & EvStart & ", " & _
            "  EventEnd = " & EvEnd & ", " & _
            "  OrderNumber = '" & m_OrderNumber & "', " & _
            "  Screen = '" & AddTicks(m_Screen) & "', " & _
            "  Reason = '" & AddTicks(Reason) & "' "
170   sql = sql & "  WHERE ISBT128 = '" & m_ISBT128 & "' " & _
            "  AND BarCode  = '" & m_Barcode & "' "
180   sql = sql & "ELSE " & _
            "  INSERT INTO Latest " & _
            "  ([Number], [BarCode], [DateExpiry], [Event], [PatID], [PatName], [Operator], [GroupRH], [Supplier], [LabNumber], " & _
            "  [crt], [cco], [cen], [crtr], [ccor], [cenr], [Checked], [Notes], [EventStart], [EventEnd], " & _
            "  [OrderNumber], [Screen], [Reason], [ISBT128], [DateTime]) " & _
            "  VALUES " & _
            "  ('" & m_PackNumber & "', " & _
            "  '" & m_Barcode & "', " & _
            "  '" & Format$(m_DateExpiry, "dd/MMM/yyyy HH:mm") & "', " & _
            "  '" & m_PackEvent & "', " & _
            "  '" & m_Chart & "', " & _
            "  '" & AddTicks(m_PatName) & "', " & _
            "  '" & UID & "', " & _
            "  '" & m_GroupRH & "', " & _
            "  '" & m_Supplier & "', " & _
            "  '" & m_SampleID & "', " & _
            "  '" & IIf(m_crt, 1, 0) & "', " & _
            "  '" & IIf(m_cco, 1, 0) & "', " & _
            "  '" & IIf(m_cen, 1, 0) & "', " & _
            "  '" & IIf(m_crtr, 1, 0) & "', " & _
            "  '" & IIf(m_ccor, 1, 0) & "', " & _
            "  '" & IIf(m_cenr, 1, 0) & "', " & _
            "  '" & IIf(m_Checked, 1, 0) & "', "
190   sql = sql & "  '" & Notes & "', " & _
            "  " & EvStart & ", " & _
            "  " & EvEnd & ", " & _
            "  '" & m_OrderNumber & "', " & _
            "  '" & AddTicks(m_Screen) & "', " & _
            "  '" & AddTicks(Reason) & "', " & _
            "  '" & m_ISBT128 & "', " & _
            "  '" & m_RecordDateTime & "')"

200   CnxnBB(0).Execute sql

210   sql = "INSERT INTO Product " & _
            "SELECT * FROM Latest " & _
            "WHERE  ISBT128 = '" & m_ISBT128 & "' " & _
            "AND BarCode  = '" & m_Barcode & "' "
220   CnxnBB(0).Execute sql

230   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "Product", "Save", intEL, strES, sql

End Sub




Public Property Get PackNumber() As String
10    PackNumber = m_PackNumber
End Property

Public Property Get PackEvent() As String
10    PackEvent = m_PackEvent
End Property

Public Property Get Notes() As String
10    Notes = m_Notes
End Property

Public Property Get EventStart() As String
10    EventStart = m_EventStart
End Property
Public Property Get EventEnd() As String
10    EventEnd = m_EventEnd
End Property

Public Property Get OrderNumber() As String
10    OrderNumber = m_OrderNumber
End Property
Public Property Get Screen() As String
10    Screen = m_Screen
End Property

Public Property Get Reason() As String
10    Reason = m_Reason
End Property
Public Property Get Supplier() As String
10    Supplier = m_Supplier
End Property

Public Property Get SampleID() As String
10    SampleID = m_SampleID
End Property

Public Property Get RecordDateTime() As String
10    RecordDateTime = m_RecordDateTime
End Property

Public Property Get GroupRh() As String
10    GroupRh = m_GroupRH
End Property



Public Property Get Chart() As String
10    Chart = m_Chart
End Property


Public Property Get PatName() As String
10    PatName = m_PatName
End Property



Public Property Get UserName() As String
10    UserName = m_UserName
End Property



Public Property Get ISBT128() As String
10    ISBT128 = m_ISBT128
End Property


Public Property Get Barcode() As String
10    Barcode = m_Barcode
End Property

Public Property Get DateExpiry() As String
10    DateExpiry = m_DateExpiry
End Property


Public Property Get crt() As Boolean
10    crt = m_crt
End Property
Public Property Get cco() As Boolean
10    cco = m_cco
End Property
Public Property Get cen() As Boolean
10    cen = m_cen
End Property

Public Property Get crtr() As Boolean
10    crtr = m_crtr
End Property
Public Property Get ccor() As Boolean
10    ccor = m_ccor
End Property
Public Property Get cenr() As Boolean
10    cenr = m_cenr
End Property
Public Property Get Checked() As Boolean
10    Checked = m_Checked
End Property

Public Property Let PackNumber(ByVal sNewValue As String)
10    m_PackNumber = sNewValue
End Property
Public Property Let Chart(ByVal sNewValue As String)
10    m_Chart = sNewValue
End Property

Public Property Let PatName(ByVal sNewValue As String)
10    m_PatName = sNewValue
End Property


Public Property Let UserName(ByVal sNewValue As String)
10    m_UserName = sNewValue
End Property



Public Property Let RecordDateTime(ByVal sNewValue As String)
10    m_RecordDateTime = sNewValue
End Property




Public Property Let Barcode(ByVal sNewValue As String)
10    m_Barcode = sNewValue
End Property

Public Property Let DateExpiry(ByVal sNewValue As String)
10    m_DateExpiry = sNewValue
End Property

Public Property Let ISBT128(ByVal sNewValue As String)
10    m_ISBT128 = sNewValue
End Property

Public Property Let GroupRh(ByVal sNewValue As String)
10    m_GroupRH = sNewValue
End Property
Public Property Let EventStart(ByVal sNewValue As String)
10    m_EventStart = sNewValue
End Property

Public Property Let EventEnd(ByVal sNewValue As String)
10    m_EventEnd = sNewValue
End Property

Public Property Let OrderNumber(ByVal sNewValue As String)
10    m_OrderNumber = sNewValue
End Property
Public Property Let Screen(ByVal sNewValue As String)
10    m_Screen = sNewValue
End Property

Public Property Let Reason(ByVal sNewValue As String)
10    m_Reason = sNewValue
End Property


Public Property Let crt(ByVal bNewValue As Boolean)
10    m_crt = bNewValue
End Property
Public Property Let cco(ByVal bNewValue As Boolean)
10    m_cco = bNewValue
End Property
Public Property Let cen(ByVal bNewValue As Boolean)
10    m_cen = bNewValue
End Property
Public Property Let crtr(ByVal bNewValue As Boolean)
10    m_crtr = bNewValue
End Property
Public Property Let ccor(ByVal bNewValue As Boolean)
10    m_ccor = bNewValue
End Property
Public Property Let cenr(ByVal bNewValue As Boolean)
10    m_cenr = bNewValue
End Property
Public Property Let Checked(ByVal bNewValue As Boolean)
10    m_Checked = bNewValue
End Property

Public Property Let Supplier(ByVal sNewValue As String)
10    m_Supplier = sNewValue
End Property


Public Property Let PackEvent(ByVal sNewValue As String)
10    m_PackEvent = sNewValue
End Property



Public Property Let SampleID(ByVal sNewValue As String)
10    m_SampleID = sNewValue
End Property


Public Property Let Notes(ByVal sNewValue As String)
10    m_Notes = sNewValue
End Property

