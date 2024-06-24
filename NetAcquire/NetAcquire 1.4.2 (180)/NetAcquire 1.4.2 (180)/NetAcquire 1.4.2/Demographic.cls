VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Demographic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SampleID As String
Private m_Chart As String
Private m_PatName As String
Private m_Age As String
Private m_Sex As String
Private m_DoB As String ' [datetime]
Private m_Addr0 As String
Private m_Addr1 As String
Private m_Ward As String
Private m_Clinician As String
Private m_GP As String
Private m_ClDetails As String
Private m_Hospital As String
Private m_RooH As Boolean
Private m_FAXed As Boolean
Private m_Fasting As Boolean
Private m_Pregnant As Boolean
Private m_sOperator As String
Private m_Category As String
Private m_UserName As String
Private m_Urgent As Integer
Private m_Valid As Boolean
Private m_ForMicro As Boolean
Private m_SentToEMedRenal As Integer
Private m_AssID As String
Private m_SurName As String
Private m_ForeName As String
Private m_ExtSampleID As String
Private m_Healthlink As Integer
Private m_AandE As String

Private m_RunDate As String
Private m_SampleDate As String
Private m_RecDate As String
Private m_DateTimeDemographics As String
Private m_DateTimeHaemPrinted As String
Private m_DateTimeBioPrinted As String
Private m_DateTimeCoagPrinted As String
Private m_RecordDateTime As String
Private m_LabNo As Double

Public Property Get SampleID() As String

9260      SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal Value As String)

9270      m_SampleID = Value

End Property

Public Property Get Chart() As String

9280      Chart = m_Chart

End Property

Public Property Let Chart(ByVal Value As String)

9290      m_Chart = Value

End Property

Public Property Get PatName() As String

9300      PatName = m_PatName

End Property

Public Property Let PatName(ByVal Value As String)

9310      m_PatName = Value

End Property

Public Property Get Age() As String

9320      Age = m_Age

End Property

Public Property Let Age(ByVal Value As String)

9330      m_Age = Value

End Property

Public Property Get Sex() As String

9340      Sex = m_Sex

End Property

Public Property Let Sex(ByVal Value As String)

9350      m_Sex = Value

End Property


Public Property Get Rundate() As String

9360      Rundate = m_RunDate

End Property

Public Property Let Rundate(ByVal Value As String)

9370      If Value = "01/01/1900" Then
9380          m_RunDate = ""
9390      Else
9400          m_RunDate = Value
9410      End If

End Property

Public Property Get DoB() As String

9420      DoB = m_DoB

End Property

Public Property Let DoB(ByVal Value As String)

9430      If Value = "01/01/1900" Then
9440          m_DoB = ""
9450      Else
9460          m_DoB = Value
9470      End If

End Property

Public Property Get Addr0() As String

9480      Addr0 = m_Addr0

End Property

Public Property Let Addr0(ByVal Value As String)

9490      m_Addr0 = Value

End Property

Public Property Get Addr1() As String

9500      Addr1 = m_Addr1

End Property

Public Property Let Addr1(ByVal Value As String)

9510      m_Addr1 = Value

End Property

Public Property Get Ward() As String

9520      Ward = m_Ward

End Property

Public Property Let Ward(ByVal Value As String)

9530      m_Ward = Value

End Property

Public Property Get Clinician() As String

9540      Clinician = m_Clinician

End Property

Public Property Let Clinician(ByVal Value As String)

9550      m_Clinician = Value

End Property

Public Property Get GP() As String

9560      GP = m_GP

End Property

Public Property Let GP(ByVal Value As String)

9570      m_GP = Value

End Property

Public Property Get SampleDate() As String

9580      SampleDate = m_SampleDate

End Property

Public Property Let SampleDate(ByVal Value As String)

9590      If Value = "01/01/1900" Then
9600          m_SampleDate = ""
9610      Else
9620          m_SampleDate = Value
9630      End If

End Property

Public Property Get ClDetails() As String

9640      ClDetails = m_ClDetails

End Property

Public Property Let ClDetails(ByVal Value As String)

9650      m_ClDetails = Value

End Property

Public Property Get Hospital() As String

9660      Hospital = m_Hospital

End Property

Public Property Let Hospital(ByVal Value As String)

9670      m_Hospital = Value

End Property

Public Property Get RooH() As Boolean

9680      RooH = m_RooH

End Property

Public Property Let RooH(ByVal Value As Boolean)

9690      m_RooH = Value

End Property

Public Property Get FAXed() As Boolean

9700      FAXed = m_FAXed

End Property

Public Property Let FAXed(ByVal Value As Boolean)

9710      m_FAXed = Value

End Property

Public Property Get Fasting() As Boolean

9720      Fasting = m_Fasting

End Property

Public Property Let Fasting(ByVal Value As Boolean)

9730      m_Fasting = Value

End Property

Public Property Get DateTimeDemographics() As String

9740      DateTimeDemographics = m_DateTimeDemographics

End Property

Public Property Let DateTimeDemographics(ByVal Value As String)

9750      If Value = "01/01/1900" Then
9760          m_DateTimeDemographics = ""
9770      Else
9780          m_DateTimeDemographics = Value
9790      End If

End Property

Public Property Get DateTimeHaemPrinted() As String

9800      DateTimeHaemPrinted = m_DateTimeHaemPrinted

End Property

Public Property Let DateTimeHaemPrinted(ByVal Value As String)

9810      If Value = "01/01/1900" Then
9820          m_DateTimeHaemPrinted = ""
9830      Else
9840          m_DateTimeHaemPrinted = Value
9850      End If

End Property

Public Property Get DateTimeBioPrinted() As String

9860      DateTimeBioPrinted = m_DateTimeBioPrinted

End Property

Public Property Let DateTimeBioPrinted(ByVal Value As String)

9870      If Value = "01/01/1900" Then
9880          m_DateTimeBioPrinted = ""
9890      Else
9900          m_DateTimeBioPrinted = Value
9910      End If

End Property

Public Property Get DateTimeCoagPrinted() As String

9920      DateTimeCoagPrinted = m_DateTimeCoagPrinted

End Property

Public Property Let DateTimeCoagPrinted(ByVal Value As String)

9930      If Value = "01/01/1900" Then
9940          m_DateTimeCoagPrinted = ""
9950      Else
9960          m_DateTimeCoagPrinted = Value
9970      End If

End Property

Public Property Get Pregnant() As Boolean

9980      Pregnant = m_Pregnant

End Property

Public Property Let Pregnant(ByVal Value As Boolean)

9990      m_Pregnant = Value

End Property

Public Property Let AandE(ByVal Value As String)

10000     m_AandE = Value

End Property

Public Property Get AandE() As String

10010     AandE = m_AandE

End Property

Public Property Get RecDate() As String

10020     RecDate = m_RecDate

End Property

Public Property Let RecDate(ByVal Value As String)

10030     If Value = "01/01/1900" Then
10040         m_RecDate = ""
10050     Else
10060         m_RecDate = Value
10070     End If

End Property

Public Property Get RecordDateTime() As String

10080     RecordDateTime = m_RecordDateTime

End Property

Public Property Let RecordDateTime(ByVal Value As String)

10090     If Value = "01/01/1900" Then
10100         m_RecordDateTime = ""
10110     Else
10120         m_RecordDateTime = Value
10130     End If

End Property

Public Property Get Operator() As String

10140     Operator = m_sOperator

End Property

Public Property Let Operator(ByVal Value As String)

10150     m_sOperator = Value

End Property

Public Property Get Category() As String

10160     Category = m_Category

End Property

Public Property Let Category(ByVal Value As String)

10170     m_Category = Value

End Property

Public Property Get UserName() As String

10180     UserName = m_UserName

End Property

Public Property Let UserName(ByVal Value As String)

10190     m_UserName = Value

End Property

Public Property Get Urgent() As Integer

10200     Urgent = m_Urgent

End Property

Public Property Let Urgent(ByVal Value As Integer)

10210     m_Urgent = Value

End Property

Public Property Get Valid() As Boolean

10220     Valid = m_Valid

End Property

Public Property Let Valid(ByVal Value As Boolean)

10230     m_Valid = Value

End Property

Public Property Get ForMicro() As Boolean

10240     ForMicro = m_ForMicro

End Property

Public Property Let ForMicro(ByVal Value As Boolean)

10250     m_ForMicro = Value

End Property

Public Property Get SentToEMedRenal() As Integer

10260     SentToEMedRenal = m_SentToEMedRenal

End Property

Public Property Let SentToEMedRenal(ByVal Value As Integer)

10270     m_SentToEMedRenal = Value

End Property

Public Property Get AssID() As String

10280     AssID = m_AssID

End Property

Public Property Let AssID(ByVal Value As String)

10290     m_AssID = Value

End Property

Public Property Get SurName() As String

10300     SurName = m_SurName

End Property

Public Property Let SurName(ByVal Value As String)

10310     m_SurName = Value

End Property

Public Property Get ForeName() As String

10320     ForeName = m_ForeName

End Property

Public Property Let ForeName(ByVal Value As String)

10330     m_ForeName = Value

End Property

Public Property Get ExtSampleID() As String

10340     ExtSampleID = m_ExtSampleID

End Property

Public Property Let ExtSampleID(ByVal Value As String)

10350     m_ExtSampleID = Value

End Property

Public Property Get HealthLink() As Integer

10360     HealthLink = m_Healthlink

End Property

Public Property Let HealthLink(ByVal Value As Integer)

10370     m_Healthlink = Value

End Property

Public Property Get LabNo() As Double

10380     LabNo = m_LabNo

End Property

Public Property Let LabNo(ByVal Value As Double)

10390     m_LabNo = Value

End Property

