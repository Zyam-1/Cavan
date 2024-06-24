VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Text As String
Private m_Code As String
Private m_Addr0  As String
Private m_Addr1  As String
Private m_InUse As Boolean
Private m_Title  As String
Private m_ForeName  As String
Private m_SurName  As String
Private m_Phone  As String
Private m_FAX  As String
Private m_Practice   As String
Private m_Compiled  As Boolean
Private m_HospitalCode   As String
Private m_ListOrder As Integer
Private m_Healthlink As Boolean
Private m_MCNumber  As String
Private m_PrintReport As Boolean
Private m_EGFR As Boolean
Private m_CC As Boolean
Private m_Interim As Boolean
Private m_PracticeNumber As String

Private Const GPsql As String = "SELECT Text, Code, Addr0, Addr1, " & _
                                "Title, ForeName, SurName, Phone, FAX, Practice, " & _
                                "HospitalCode, COALESCE(ListOrder, 999) ListOrder, MCNumber, " & _
                                "COALESCE(InUse, 1) InUse, " & _
                                "COALESCE(Compiled, 1) Compiled, " & _
                                "COALESCE(HealthLink, 1) Healthlink, " & _
                                "COALESCE(PrintReport, 1) PrintReport, " & _
                                "COALESCE(AutoCC, 0) AutoCC, " & _
                                "COALESCE(Interim, 0) Interim, " & _
                                "COALESCE(PracticeNumber, '') PracticeNumber " & _
                                "FROM GPs WHERE "

Public Sub LoadName(ByVal GPName As String)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo LoadName_Error

20    sql = GPsql & "Text = '" & AddTicks(GPName) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60      Populate tb
70    End If

80    Exit Sub

LoadName_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "GP", "LoadName", intEL, strES, sql

End Sub

Public Sub LoadCode(ByVal GPCode As String)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo LoadCode_Error

20    sql = GPsql & "Code = '" & AddTicks(GPCode) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60      Populate tb
70    End If

80    Exit Sub

LoadCode_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "GP", "LoadCode", intEL, strES, sql

End Sub

Public Sub LoadCodeOrText(ByVal CodeOrText As String)

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo LoadCodeOrText_Error

20    s = AddTicks(CodeOrText)

30    sql = GPsql & " ISNULL(InUse,0) = 1 AND (Code = '" & s & "' " & _
                    "OR Text = '" & s & "')"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      Populate tb
80    End If

90    Exit Sub

LoadCodeOrText_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "GP", "LoadCodeOrText", intEL, strES, sql

End Sub

Private Sub Populate(ByVal tb As Recordset)

10    With tb
20        m_Text = !Text & ""
30        m_Code = !Code & ""
40        m_Addr0 = !Addr0 & ""
50        m_Addr1 = !Addr1 & ""
60        m_InUse = !InUse
70        m_Title = !Title & ""
80        m_ForeName = !ForeName & ""
90        m_SurName = !SurName & ""
100       m_Phone = !Phone & ""
110       m_FAX = !FAX & ""
120       m_Practice = !Practice & ""
130       m_Compiled = !Compiled
140       m_HospitalCode = !HospitalCode & ""
150       m_ListOrder = !ListOrder
160       m_Healthlink = !Healthlink
170       m_MCNumber = !McNumber & ""
180       m_PracticeNumber = !PracticeNumber & ""
190       m_PrintReport = !PrintReport
200       m_CC = !AutoCC
210       m_Interim = !Interim
220   End With

End Sub


Public Property Get Text() As String
10    Text = m_Text
End Property

Public Property Get Code() As String
10    Code = m_Code
End Property


Public Property Get Addr0() As String
10    Addr0 = m_Addr0
End Property

Public Property Get McNumber() As String
10    McNumber = m_MCNumber
End Property

Public Property Get PracticeNumber() As String
10    PracticeNumber = m_PracticeNumber
End Property

Public Property Get ListOrder() As Integer
10    ListOrder = m_ListOrder
End Property


Public Property Get HospitalCode() As String
10    HospitalCode = m_HospitalCode
End Property


Public Property Get Practice() As String
10    Practice = m_Practice
End Property


Public Property Get Phone() As String
10    Phone = m_Phone
End Property


Public Property Get FAX() As String
10    FAX = m_FAX
End Property


Public Property Get ForeName() As String
10    ForeName = m_ForeName
End Property


Public Property Get SurName() As String
10    SurName = m_SurName
End Property



Public Property Get Title() As String
10    Title = m_Title
End Property


Public Property Get InUse() As Boolean
10    InUse = m_InUse
End Property


Public Property Get Healthlink() As Boolean
10    Healthlink = m_Healthlink
End Property
Public Property Get Interim() As Boolean
10    Interim = m_Interim
End Property


Public Property Get Compiled() As Boolean
10    Compiled = m_Compiled
End Property


Public Property Get EGFR() As Boolean
10    EGFR = m_EGFR
End Property

Public Property Get AutoCC() As Boolean
10    AutoCC = m_CC
End Property

Public Property Get PrintReport() As Boolean
10    PrintReport = m_PrintReport
End Property
Public Property Get Addr1() As String
10    Addr1 = m_Addr1
End Property


Public Property Let ForeName(ByVal sNewValue As String)
10    m_ForeName = sNewValue
End Property
Public Property Let FAX(ByVal sNewValue As String)
10    m_FAX = sNewValue
End Property

Public Property Let Phone(ByVal sNewValue As String)
10    m_Phone = sNewValue
End Property


Public Property Let McNumber(ByVal sNewValue As String)
10    m_MCNumber = sNewValue
End Property

Public Property Let PracticeNumber(ByVal sNewValue As String)
10    m_PracticeNumber = sNewValue
End Property

Public Property Let ListOrder(ByVal iNewValue As Integer)
10    m_ListOrder = iNewValue
End Property



Public Property Let HospitalCode(ByVal sNewValue As String)
10    m_HospitalCode = sNewValue
End Property



Public Property Let SurName(ByVal sNewValue As String)
10    m_SurName = sNewValue
End Property

Public Property Let Practice(ByVal sNewValue As String)
10    m_Practice = sNewValue
End Property


Public Property Let Title(ByVal sNewValue As String)
10    m_Title = sNewValue
End Property

Public Property Let Text(ByVal sNewValue As String)
10    m_Text = sNewValue
End Property


Public Property Let Code(ByVal sNewValue As String)
10    m_Code = sNewValue
End Property

Public Property Let InUse(ByVal bNewValue As Boolean)
10    m_InUse = bNewValue
End Property

Public Property Let PrintReport(ByVal bNewValue As Boolean)
10    m_PrintReport = bNewValue
End Property

Public Property Let Healthlink(ByVal bNewValue As Boolean)
10    m_Healthlink = bNewValue
End Property
Public Property Let Interim(ByVal bNewValue As Boolean)
10    m_Interim = bNewValue
End Property


Public Property Let Compiled(ByVal bNewValue As Boolean)
10    m_Compiled = bNewValue
End Property


Public Property Let EGFR(ByVal bNewValue As Boolean)
10    m_EGFR = bNewValue
End Property

Public Property Let AutoCC(ByVal bNewValue As Boolean)
10    m_CC = bNewValue
End Property

Public Property Let Addr0(ByVal sNewValue As String)
10    m_Addr0 = sNewValue
End Property


Public Property Let Addr1(ByVal sNewValue As String)
10    m_Addr1 = sNewValue
End Property



Private Sub Class_Initialize()

10    m_Text = ""
20    m_Code = ""
30    m_Addr0 = ""
40    m_Addr1 = ""
50    m_InUse = False
60    m_Title = ""
70    m_ForeName = ""
80    m_SurName = ""
90    m_Phone = ""
100   m_FAX = ""
110   m_Practice = ""
120   m_Compiled = False
130   m_HospitalCode = ""
140   m_ListOrder = 999
150   m_Healthlink = False
160   m_MCNumber = ""
170   m_PracticeNumber = ""
180   m_PrintReport = True
190   m_Interim = False
End Sub


