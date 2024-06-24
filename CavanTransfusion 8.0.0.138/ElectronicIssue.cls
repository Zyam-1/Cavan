VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElectronicIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PreviousSample As Integer
Private m_PreviousGroupAgreement As Integer
Private m_CurrentNegativeAB As Integer
Private m_PreviousNegativeAB As Integer
Private m_AdverseReactions As Integer
Private m_PreviousSampleEligible As Integer
Private m_Chart As String
Private m_SampleDate As String
Private m_SampleID As String
Private m_ForcedEligible As Integer
Private m_ForcedNotEligible As Integer
Private m_ResultAbnormalFlags As Integer

Public Sub Load()

10    If Trim$(m_Chart) <> "" And IsDate(m_SampleDate) And Trim$(m_SampleID) <> "" Then
20      GetHasPreviousSample
30      GetGroupAgreement
'40      GetCurrentNegativeAB
40      GetCurrentNegativeAB_From_Vision
50      GetPreviousNegativeAB
60      GetAdverseReactions
70      GetPreviousSampleEligible
80      GetForced
90      GetVisionResultAbnormalFlags
100   End If

End Sub

Public Sub GetVisionResultAbnormalFlags()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetVisionResultAbnormalFlags_Error

20    m_ResultAbnormalFlags = 2

30    sql = "SELECT ResultAbnormalFlags FROM AutoVue " & _
            "WHERE SampleId = '" & m_SampleID & "' " & _
            "ORDER BY MessageDateTime desc"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF

70      If UCase(tb!ResultAbnormalFlags) = "M" Then
80          m_ResultAbnormalFlags = 0
90          Exit Do
100     Else
110         m_ResultAbnormalFlags = 1
120     End If

130     tb.MoveNext
140   Loop

150   Exit Sub

GetVisionResultAbnormalFlags_Error:

 Dim strES As String
 Dim intEL As Integer

160    intEL = Erl
170    strES = Err.Description
180    LogError "ElectronicIssue", "GetVisionResultAbnormalFlags", intEL, strES, sql

End Sub

Public Sub GetForced()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetForced_Error

20    m_ForcedEligible = 2
30    m_ForcedNotEligible = 2

      'Is sample Eligible for EI  (either Automatically Eligible or Forced Eligible)
40    sql = "Select Eligible4EI from PatientDetails where " & _
            "labnumber = '" & m_SampleID & "' and (Eligible4EI = 'F_E' or Eligible4EI = 'FNE')"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    If Not tb.EOF Then
80      If Trim$(tb!Eligible4EI) = "F_E" Then
90        m_ForcedEligible = 1
100       m_ForcedNotEligible = 0
110     ElseIf Trim$(tb!Eligible4EI) = "FNE" Then
120       m_ForcedEligible = 0
130       m_ForcedNotEligible = 1
140     End If
150   End If

160   Exit Sub

GetForced_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "ElectronicIssue", "GetForced", intEL, strES, sql


End Sub

Private Sub GetPreviousSampleEligible()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetPreviousSampleEligible_Error

20    m_PreviousSampleEligible = 2

30    sql = "SELECT TOP 2 Eligible4EI Flag FROM PatientDetails " & _
            "WHERE PatNum = '" & AddTicks(m_Chart) & "' " & _
            "AND DATEDIFF(minute, SampleDate, '" & m_SampleDate & "') >= 0 " & _
            "and fGroup <> '' " & _
            "ORDER BY SampleDate desc"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      tb.MoveNext 'move to 2nd sample
80      If Not tb.EOF Then
90        If Trim$(tb!Flag & "") = "A_E" Or Trim$(tb!Flag & "") = "F_E" Or Trim$(tb!Flag & "") = "" Then  'either Automatically Eligible or Forced Eligible
100         m_PreviousSampleEligible = 1
110       Else
120         m_PreviousSampleEligible = 0
130       End If
140     End If
150   End If

160   Exit Sub

GetPreviousSampleEligible_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "ElectronicIssue", "GetPreviousSampleEligible", intEL, strES, sql


End Sub

Private Sub GetAdverseReactions()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetAdverseReactions_Error

20    m_AdverseReactions = 2

30    sql = "SELECT COUNT(*) Tot FROM BadReact " & _
            "WHERE PatNo = '" & AddTicks(m_Chart) & "'"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If tb!Tot = 0 Then
70      sql = "SELECT COUNT(*) Tot FROM AdverseReactions " & _
              "WHERE Chart = '" & AddTicks(m_Chart) & "'"
80      Set tb = New Recordset
90      RecOpenServerBB 0, tb, sql
100     If tb!Tot = 0 Then
110       m_AdverseReactions = 0
120     Else
130       m_AdverseReactions = 1
140     End If
150   End If

160   Exit Sub

GetAdverseReactions_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "ElectronicIssue", "GetAdverseReactions", intEL, strES, sql

End Sub

Private Sub GetPreviousNegativeAB()

Dim sql As String
Dim rs As Recordset
Dim tb As Recordset
Dim ch As String

10    On Error GoTo GetPreviousNegativeAB_Error

20    ch = AddTicks(m_Chart)

30    m_PreviousNegativeAB = 2

32    sql = "SELECT * FROM PatientDetails WHERE PatNum = '" & ch & "'  AND DATEDIFF(minute, SampleDate, '" & m_SampleDate & "') >= 0   AND (AIDR is not null and AIDR <> '') and fGroup <> ''"
    
35    Set rs = New Recordset
37    RecOpenServerBB 0, rs, sql
    
38    If Not rs.EOF Then
40      sql = "SELECT CASE (SELECT COUNT(*) FROM PatientDetails " & _
        "             WHERE PatNum = '" & ch & "' " & _
        "             AND DATEDIFF(minute, SampleDate, '" & m_SampleDate & "') >= 0 ) " & _
        "  WHEN 0 THEN 2 " & _
        "  WHEN 1 THEN 2 " & _
        "  ELSE (SELECT CASE " & _
        "           (SELECT COUNT(*) FROM PatientDetails " & _
        "           WHERE PatNum = '" & ch & "' " & _
        "           AND DATEDIFF(minute, SampleDate, '" & m_SampleDate & "') >= 0 " & _
        "           AND AIDR LIKE '%anti%' ) " & _
        "        WHEN 0 THEN 1 " & _
        "        ELSE 0 " & _
        "        END) " & _
        "  END Tot"
50      Set tb = New Recordset
60      RecOpenServerBB 0, tb, sql
70      m_PreviousNegativeAB = tb!Tot
75    End If
    
80    Exit Sub

GetPreviousNegativeAB_Error:

Dim strES As String
Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "ElectronicIssue", "GetPreviousNegativeAB", intEL, strES, sql

End Sub

Private Sub GetCurrentNegativeAB()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetCurrentNegativeAB_Error

20    CurrentNegativeAB = 2

30    sql = "SELECT TOP 1 AIDR FROM PatientDetails " & _
            "WHERE LabNumber = '" & AddTicks(m_SampleID) & "' " & _
            "ORDER BY DateTime desc"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
70      If InStr(UCase$(tb!AIDR & ""), "NEG") > 0 Then
80        m_CurrentNegativeAB = 1
90      Else
100       m_CurrentNegativeAB = 0
110     End If
120   End If

130   Exit Sub

GetCurrentNegativeAB_Error:

Dim strES As String
Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "ElectronicIssue", "GetCurrentNegativeAB", intEL, strES, sql


End Sub

Private Sub GetCurrentNegativeAB_From_Vision()
'Get Latest AB screen result sent from Vision un-modified
'Modified results prevent Electronic Issue

      Dim sql As String
      Dim tb As Recordset

   On Error GoTo GetCurrentNegativeAB_From_Vision_Error

20    CurrentNegativeAB = 2

30    sql = "SELECT ABSCR_Result FROM VisionABscreenResults " & _
            "WHERE SampleID = '" & AddTicks(m_SampleID) & "' and Result_Flag <> 'M' and ABSCR_Result like '%NEG%' "
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then
80        m_CurrentNegativeAB = 1
90    Else
100       m_CurrentNegativeAB = 0
120   End If

   Exit Sub

GetCurrentNegativeAB_From_Vision_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "ElectronicIssue", "GetCurrentNegativeAB_From_Vision", intEL, strES, sql


End Sub


Private Sub GetGroupAgreement()

      Dim sql As String
      Dim tb As Recordset
      Dim CurrentGroup As String

10    On Error GoTo GetGroupAgreement_Error

20    m_PreviousGroupAgreement = 2

30    sql = "SELECT top 1 fGroup FROM PatientDetails " & _
            "WHERE LabNumber = '" & m_SampleID & "' " & _
            "and fGroup <> '' " & _
            "ORDER BY DateTime desc"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    If Not tb.EOF Then 'Current sample has a Group saved. Therefore can be compared to previous sample
70        sql = "SELECT TOP 2 fGroup, DateTime FROM PatientDetails " & _
                "WHERE PatNum = '" & AddTicks(m_Chart) & "' " & _
                "and fGroup <> '' " & _
                "AND DATEDIFF(minute, SampleDate, '" & m_SampleDate & "') >= 0 ORDER BY DateTime desc"
80        Set tb = New Recordset
90        RecOpenServerBB 0, tb, sql
100       If Not tb.EOF Then
110         CurrentGroup = UCase$(tb!fGroup & "")
120         If Trim$(CurrentGroup) <> "" Then
130           tb.MoveNext
140           If Not tb.EOF Then
150             If CurrentGroup = UCase$(tb!fGroup & "") Then
160               m_PreviousGroupAgreement = 1
170             Else
180               m_PreviousGroupAgreement = 0
190             End If
200           End If
210         End If
220       End If
230   End If

240   Exit Sub

GetGroupAgreement_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "ElectronicIssue", "GetGroupAgreement", intEL, strES, sql

End Sub

Private Sub GetHasPreviousSample()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetHasPreviousSample_Error

20    m_PreviousSample = 2

30    sql = "SELECT COUNT(*) Tot FROM PatientDetails " & _
            "WHERE PatNum = '" & AddTicks(m_Chart) & "' " & _
            "AND DATEDIFF(minute, SampleDate, '" & Format$(m_SampleDate, "dd/MMM/yyyy HH:nn") & "') > 0 "
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql

60    m_PreviousSample = IIf(tb!Tot > 0, 1, 0)

70    Exit Sub

GetHasPreviousSample_Error:

Dim strES As String
Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "ElectronicIssue", "GetHasPreviousSample", intEL, strES, sql


End Sub

Public Property Get PreviousSample() As Integer

10      PreviousSample = m_PreviousSample

End Property

Public Property Let PreviousSample(ByVal PreviousSample As Integer)

10      m_PreviousSample = PreviousSample

End Property

Public Property Get PreviousGroupAgreement() As Integer

10      PreviousGroupAgreement = m_PreviousGroupAgreement

End Property

Public Property Let PreviousGroupAgreement(ByVal PreviousGroupAgreement As Integer)

10      m_PreviousGroupAgreement = PreviousGroupAgreement

End Property

Public Property Get CurrentNegativeAB() As Integer

10      CurrentNegativeAB = m_CurrentNegativeAB

End Property

Public Property Let CurrentNegativeAB(ByVal CurrentNegativeAB As Integer)

10      m_CurrentNegativeAB = CurrentNegativeAB

End Property

Public Property Get PreviousNegativeAB() As Integer

10      PreviousNegativeAB = m_PreviousNegativeAB

End Property

Public Property Let PreviousNegativeAB(ByVal PreviousNegativeAB As Integer)

10      m_PreviousNegativeAB = PreviousNegativeAB

End Property

Public Property Get AdverseReactions() As Integer

10      AdverseReactions = m_AdverseReactions

End Property

Public Property Let AdverseReactions(ByVal AdverseReactions As Integer)

10      m_AdverseReactions = AdverseReactions

End Property

Public Property Get PreviousSampleEligible() As Integer

10      PreviousSampleEligible = m_PreviousSampleEligible

End Property

Public Property Let PreviousSampleEligible(ByVal PreviousSampleEligible As Integer)

10      m_PreviousSampleEligible = PreviousSampleEligible

End Property

Public Property Let Chart(ByVal Chart As String)

10      m_Chart = Chart

End Property

Public Property Let SampleDate(ByVal SampleDate As String)

10      m_SampleDate = Format$(SampleDate, "dd/MMM/yyyy HH:nn")

End Property

Public Property Let SampleID(ByVal SampleID As String)

10      m_SampleID = SampleID

End Property


Public Property Get ForcedEligible() As Integer

10      ForcedEligible = m_ForcedEligible

End Property

Public Property Let ForcedEligible(ByVal ForcedEligible As Integer)

10      m_ForcedEligible = ForcedEligible

End Property

Public Property Get ForcedNotEligible() As Integer

10      ForcedNotEligible = m_ForcedNotEligible

End Property

Public Property Let ForcedNotEligible(ByVal ForcedNotEligible As Integer)

10      m_ForcedNotEligible = ForcedNotEligible

End Property


Public Property Get ResultAbnormalFlags() As Integer

10      ResultAbnormalFlags = m_ResultAbnormalFlags

End Property

Public Property Let ResultAbnormalFlags(ByVal ResultAbnormalFlags As Integer)

10      m_ResultAbnormalFlags = ResultAbnormalFlags

End Property

Public Sub Save()

Dim sql As String
Dim Flag As Integer
Dim strFlag As String


10    On Error GoTo Save_Error

20    strFlag = ""

30    If m_ForcedEligible = 1 Then
40      Flag = 64
50      strFlag = "F_E"
60    ElseIf m_ForcedNotEligible = 1 Then
70      Flag = 128
80      strFlag = "FNE"
90    End If

100   If Len(strFlag) > 0 Then
        'F_E = "Forced Eligible for Electronic Issue"
        'FNE = "Forced NOT Eligible for Electronic Issue"
110     sql = "UPDATE PatientDetails SET Eligible4EI = '" & strFlag & "' WHERE (labnumber = '" & m_SampleID & "')"
120   CnxnBB(0).Execute sql
130   End If

140   Exit Sub

Save_Error:

 Dim strES As String
 Dim intEL As Integer

150    intEL = Erl
160    strES = Err.Description
170    LogError "ElectronicIssue", "Save", intEL, strES, sql

End Sub


