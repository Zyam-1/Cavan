VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_LIorH As String
Private m_Code As String
Private m_CutOff As Single
Private m_NoPrintOrWarning As String
Private m_UserName As String
Private m_dtDateTimeOfRecord As Date

Public Sub Save()

      Dim sql As String

10    On Error GoTo Save_Error

20    sql = "IF EXISTS(SELECT * FROM LIHValues " & _
            "          WHERE LIH = '" & m_LIorH & "' " & _
            "          AND Code = '" & m_Code & "' " & _
            "          AND NoPrintOrWarning = '" & m_NoPrintOrWarning & "' ) " & _
            "  UPDATE LIHValues " & _
            "  SET CutOff = '" & m_CutOff & "', " & _
            "  UserName = '" & AddTicks(UserName) & "', " & _
            "  DateTimeOfRecord = getdate() " & _
            "  WHERE LIH = '" & m_LIorH & "' " & _
            "  AND Code = '" & m_Code & "' AND NoPrintOrWarning = '" & m_NoPrintOrWarning & "'  " & _
            "ELSE " & _
            "  INSERT INTO LIHValues " & _
            "  (LIH, Code, CutOff, NoPrintOrWarning, UserName) " & _
            "  VALUES (" & _
            "  '" & m_LIorH & "', " & _
            "  '" & m_Code & "', " & _
            "  '" & m_CutOff & "', " & _
            "  '" & m_NoPrintOrWarning & "', " & _
            "  '" & AddTicks(UserName) & "')"

30    Cnxn(0).Execute sql

40    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "LIHs", "Save", intEL, strES, sql

End Sub

Public Property Get LIorH() As String

10      LIorH = m_LIorH

End Property

Public Property Let LIorH(ByVal LIorH As String)

10      m_LIorH = LIorH

End Property

Public Property Get Code() As String

  Code = m_Code

End Property

Public Property Let Code(ByVal Code As String)

  m_Code = Code

End Property

Public Property Get CutOff() As Single

  CutOff = m_CutOff

End Property

Public Property Let CutOff(ByVal CutOff As Single)

  m_CutOff = CutOff

End Property

Public Property Get UserName() As String

  UserName = m_UserName

End Property

Public Property Let UserName(ByVal UserName As String)

  m_UserName = UserName

End Property

Public Property Get DateTimeOfRecord() As Date

  DateTimeOfRecord = m_dtDateTimeOfRecord

End Property

Public Property Let DateTimeOfRecord(ByVal dtDateTimeOfRecord As Date)

  m_dtDateTimeOfRecord = dtDateTimeOfRecord

End Property

Public Property Get NoPrintOrWarning() As String

  NoPrintOrWarning = m_NoPrintOrWarning

End Property

Public Property Let NoPrintOrWarning(ByVal NoPrintOrWarning As String)

  m_NoPrintOrWarning = NoPrintOrWarning

End Property
