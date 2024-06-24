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

20750     On Error GoTo Save_Error

20760     sql = "IF EXISTS(SELECT * FROM LIHValues " & _
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

20770     Cnxn(0).Execute sql

20780     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

20790     intEL = Erl
20800     strES = Err.Description
20810     LogError "LIHs", "Save", intEL, strES, sql

End Sub

Public Property Get LIorH() As String

20820     LIorH = m_LIorH

End Property

Public Property Let LIorH(ByVal LIorH As String)

20830     m_LIorH = LIorH

End Property

Public Property Get Code() As String

20840     Code = m_Code

End Property

Public Property Let Code(ByVal Code As String)

20850     m_Code = Code

End Property

Public Property Get CutOff() As Single

20860     CutOff = m_CutOff

End Property

Public Property Let CutOff(ByVal CutOff As Single)

20870     m_CutOff = CutOff

End Property

Public Property Get UserName() As String

20880     UserName = m_UserName

End Property

Public Property Let UserName(ByVal UserName As String)

20890     m_UserName = UserName

End Property

Public Property Get DateTimeOfRecord() As Date

20900     DateTimeOfRecord = m_dtDateTimeOfRecord

End Property

Public Property Let DateTimeOfRecord(ByVal dtDateTimeOfRecord As Date)

20910     m_dtDateTimeOfRecord = dtDateTimeOfRecord

End Property

Public Property Get NoPrintOrWarning() As String

20920     NoPrintOrWarning = m_NoPrintOrWarning

End Property

Public Property Let NoPrintOrWarning(ByVal NoPrintOrWarning As String)

20930     m_NoPrintOrWarning = NoPrintOrWarning

End Property
