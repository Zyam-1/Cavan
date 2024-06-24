VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Isolate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SampleID As String
Private m_IsolateNumber As Integer
Private m_OrganismGroup As String
Private m_OrganismName As String
Private m_Qualifier As String
Private m_UserName As String
Private m_RecordDateTime As String
Private m_Valid As Integer
Private m_Printed As Integer

Public Sub Save()

          Dim sql As String

20000     On Error GoTo Save_Error

20010     sql = "IF EXISTS (SELECT * FROM Isolates WHERE " & _
              "           SampleID = '" & m_SampleID & "' " & _
              "           AND IsolateNumber = '" & m_IsolateNumber & "') " & _
              "  UPDATE Isolates " & _
              "  SET OrganismGroup = '" & m_OrganismGroup & "', " & _
              "  OrganismName = '" & m_OrganismName & "', " & _
              "  Qualifier = '" & m_Qualifier & "', " & _
              "  UserName = '" & AddTicks(m_UserName) & "', RecordDateTime = GetDate() " & _
              "  WHERE SampleID = '" & m_SampleID & "' " & _
              "  AND IsolateNumber = '" & m_IsolateNumber & "' " & _
              "ELSE " & _
              "  INSERT INTO Isolates (SampleID, IsolateNumber, OrganismGroup, OrganismName, Qualifier, UserName, RecordDateTime) " & _
              "  VALUES " & _
              "  ('" & m_SampleID & "', " & _
              "  '" & m_IsolateNumber & "', " & _
              "  '" & m_OrganismGroup & "', " & _
              "  '" & m_OrganismName & "', " & _
              "  '" & m_Qualifier & "', " & _
              "  '" & AddTicks(m_UserName) & "', GetDate())"
20020     Cnxn(0).Execute sql

20030     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

20040     intEL = Erl
20050     strES = Err.Description
20060     LogError "Isolate", "Save", intEL, strES, sql

End Sub

Public Property Get SampleID() As String

20070     SampleID = m_SampleID

End Property

Public Property Get RecordDateTime() As String

20080     RecordDateTime = m_RecordDateTime

End Property
Public Property Let SampleID(ByVal sNewValue As String)

20090     m_SampleID = sNewValue

End Property

Public Property Get IsolateNumber() As Integer

20100     IsolateNumber = m_IsolateNumber

End Property

Public Property Get Valid() As Integer

20110     Valid = m_Valid

End Property


Public Property Get Printed() As Integer

20120     Printed = m_Printed

End Property



Public Property Let IsolateNumber(ByVal iNewValue As Integer)

20130     m_IsolateNumber = iNewValue

End Property

Public Property Let Valid(ByVal iNewValue As Integer)

20140     m_Valid = iNewValue

End Property


Public Property Let Printed(ByVal iNewValue As Integer)

20150     m_Printed = iNewValue

End Property



Public Property Get OrganismGroup() As String

20160     OrganismGroup = m_OrganismGroup

End Property

Public Property Get OrganismName() As String

20170     OrganismName = m_OrganismName

End Property
Public Property Get Qualifier() As String

20180     Qualifier = m_Qualifier

End Property

Public Property Get UserName() As String

20190     UserName = m_UserName

End Property

Public Property Let OrganismGroup(ByVal sNewValue As String)

20200     m_OrganismGroup = sNewValue

End Property
Public Property Let OrganismName(ByVal sNewValue As String)

20210     m_OrganismName = sNewValue

End Property

Public Property Let Qualifier(ByVal sNewValue As String)

20220     m_Qualifier = sNewValue

End Property

Public Property Let RecordDateTime(ByVal sNewValue As String)

20230     m_RecordDateTime = sNewValue

End Property

Public Property Let UserName(ByVal sNewValue As String)

20240     m_UserName = sNewValue

End Property


