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

15150     On Error GoTo LoadName_Error

15160     sql = GPsql & "Text = '" & AddTicks(GPName) & "'"
15170     Set tb = New Recordset
15180     RecOpenServer 0, tb, sql
15190     If Not tb.EOF Then
15200         Populate tb
15210     End If

15220     Exit Sub

LoadName_Error:

          Dim strES As String
          Dim intEL As Integer

15230     intEL = Erl
15240     strES = Err.Description
15250     LogError "GP", "LoadName", intEL, strES, sql

End Sub

Public Sub LoadCode(ByVal GPCode As String)

          Dim sql As String
          Dim tb As Recordset

15260     On Error GoTo LoadCode_Error

15270     sql = GPsql & "Code = '" & AddTicks(GPCode) & "'"
15280     Set tb = New Recordset
15290     RecOpenServer 0, tb, sql
15300     If Not tb.EOF Then
15310         Populate tb
15320     End If

15330     Exit Sub

LoadCode_Error:

          Dim strES As String
          Dim intEL As Integer

15340     intEL = Erl
15350     strES = Err.Description
15360     LogError "GP", "LoadCode", intEL, strES, sql

End Sub

Public Sub LoadCodeOrText(ByVal CodeOrText As String)

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

15370     On Error GoTo LoadCodeOrText_Error

15380     s = AddTicks(CodeOrText)

15390     sql = GPsql & " ISNULL(InUse,0) = 1 AND (Code = '" & s & "' " & _
              "OR Text = '" & s & "')"
15400     Set tb = New Recordset
15410     RecOpenServer 0, tb, sql
15420     If Not tb.EOF Then
15430         Populate tb
15440     End If

15450     Exit Sub

LoadCodeOrText_Error:

          Dim strES As String
          Dim intEL As Integer

15460     intEL = Erl
15470     strES = Err.Description
15480     LogError "GP", "LoadCodeOrText", intEL, strES, sql

End Sub

Private Sub Populate(ByVal tb As Recordset)

15490     With tb
15500         m_Text = !Text & ""
15510         m_Code = !Code & ""
15520         m_Addr0 = !Addr0 & ""
15530         m_Addr1 = !Addr1 & ""
15540         m_InUse = !InUse
15550         m_Title = !Title & ""
15560         m_ForeName = !ForeName & ""
15570         m_SurName = !SurName & ""
15580         m_Phone = !Phone & ""
15590         m_FAX = !FAX & ""
15600         m_Practice = !Practice & ""
15610         m_Compiled = !Compiled
15620         m_HospitalCode = !HospitalCode & ""
15630         m_ListOrder = !ListOrder
15640         m_Healthlink = !HealthLink
15650         m_MCNumber = !McNumber & ""
15660         m_PracticeNumber = !PracticeNumber & ""
15670         m_PrintReport = !PrintReport
15680         m_CC = !AutoCC
15690         m_Interim = !Interim
15700     End With

End Sub


Public Property Get Text() As String
15710     Text = m_Text
End Property

Public Property Get Code() As String
15720     Code = m_Code
End Property


Public Property Get Addr0() As String
15730     Addr0 = m_Addr0
End Property

Public Property Get McNumber() As String
15740     McNumber = m_MCNumber
End Property

Public Property Get PracticeNumber() As String
15750     PracticeNumber = m_PracticeNumber
End Property

Public Property Get ListOrder() As Integer
15760     ListOrder = m_ListOrder
End Property


Public Property Get HospitalCode() As String
15770     HospitalCode = m_HospitalCode
End Property


Public Property Get Practice() As String
15780     Practice = m_Practice
End Property


Public Property Get Phone() As String
15790     Phone = m_Phone
End Property


Public Property Get FAX() As String
15800     FAX = m_FAX
End Property


Public Property Get ForeName() As String
15810     ForeName = m_ForeName
End Property


Public Property Get SurName() As String
15820     SurName = m_SurName
End Property



Public Property Get Title() As String
15830     Title = m_Title
End Property


Public Property Get InUse() As Boolean
15840     InUse = m_InUse
End Property


Public Property Get HealthLink() As Boolean
15850     HealthLink = m_Healthlink
End Property
Public Property Get Interim() As Boolean
15860     Interim = m_Interim
End Property


Public Property Get Compiled() As Boolean
15870     Compiled = m_Compiled
End Property


Public Property Get EGFR() As Boolean
15880     EGFR = m_EGFR
End Property

Public Property Get AutoCC() As Boolean
15890     AutoCC = m_CC
End Property

Public Property Get PrintReport() As Boolean
15900     PrintReport = m_PrintReport
End Property
Public Property Get Addr1() As String
15910     Addr1 = m_Addr1
End Property


Public Property Let ForeName(ByVal sNewValue As String)
15920     m_ForeName = sNewValue
End Property
Public Property Let FAX(ByVal sNewValue As String)
15930     m_FAX = sNewValue
End Property

Public Property Let Phone(ByVal sNewValue As String)
15940     m_Phone = sNewValue
End Property


Public Property Let McNumber(ByVal sNewValue As String)
15950     m_MCNumber = sNewValue
End Property

Public Property Let PracticeNumber(ByVal sNewValue As String)
15960     m_PracticeNumber = sNewValue
End Property

Public Property Let ListOrder(ByVal iNewValue As Integer)
15970     m_ListOrder = iNewValue
End Property



Public Property Let HospitalCode(ByVal sNewValue As String)
15980     m_HospitalCode = sNewValue
End Property



Public Property Let SurName(ByVal sNewValue As String)
15990     m_SurName = sNewValue
End Property

Public Property Let Practice(ByVal sNewValue As String)
16000     m_Practice = sNewValue
End Property


Public Property Let Title(ByVal sNewValue As String)
16010     m_Title = sNewValue
End Property

Public Property Let Text(ByVal sNewValue As String)
16020     m_Text = sNewValue
End Property


Public Property Let Code(ByVal sNewValue As String)
16030     m_Code = sNewValue
End Property

Public Property Let InUse(ByVal bNewValue As Boolean)
16040     m_InUse = bNewValue
End Property

Public Property Let PrintReport(ByVal bNewValue As Boolean)
16050     m_PrintReport = bNewValue
End Property

Public Property Let HealthLink(ByVal bNewValue As Boolean)
16060     m_Healthlink = bNewValue
End Property
Public Property Let Interim(ByVal bNewValue As Boolean)
16070     m_Interim = bNewValue
End Property


Public Property Let Compiled(ByVal bNewValue As Boolean)
16080     m_Compiled = bNewValue
End Property


Public Property Let EGFR(ByVal bNewValue As Boolean)
16090     m_EGFR = bNewValue
End Property

Public Property Let AutoCC(ByVal bNewValue As Boolean)
16100     m_CC = bNewValue
End Property

Public Property Let Addr0(ByVal sNewValue As String)
16110     m_Addr0 = sNewValue
End Property


Public Property Let Addr1(ByVal sNewValue As String)
16120     m_Addr1 = sNewValue
End Property



Private Sub Class_Initialize()

16130     m_Text = ""
16140     m_Code = ""
16150     m_Addr0 = ""
16160     m_Addr1 = ""
16170     m_InUse = False
16180     m_Title = ""
16190     m_ForeName = ""
16200     m_SurName = ""
16210     m_Phone = ""
16220     m_FAX = ""
16230     m_Practice = ""
16240     m_Compiled = False
16250     m_HospitalCode = ""
16260     m_ListOrder = 999
16270     m_Healthlink = False
16280     m_MCNumber = ""
16290     m_PracticeNumber = ""
16300     m_PrintReport = True
16310     m_Interim = False
End Sub


