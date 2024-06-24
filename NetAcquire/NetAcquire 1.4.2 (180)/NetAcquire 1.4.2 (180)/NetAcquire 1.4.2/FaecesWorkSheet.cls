VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaecesWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SampleID As String
Private m_DayIndex As String
Private m_Result As String
Private m_UserName As String
Private m_DateTimeOfRecord As Date

Public Sub Save()

          Dim sql As String

13100     On Error GoTo Save_Error

13110     If Trim$(m_Result) = "" Then
              '30      Sql = "DELETE FROM FaecesWorkSheets50 " & _
              '              "WHERE SampleID = '" & m_SampleID & "' " & _
              '              "AND DayIndex = '" & m_DayIndex & "'"
              '        Exit Sub
13120     Else
13130         sql = "IF EXISTS (SELECT * FROM FaecesWorkSheets50 WHERE " & _
                  "           SampleID = '" & m_SampleID & "' " & _
                  "           AND DayIndex = '" & m_DayIndex & "') " & _
                  "  UPDATE FaecesWorkSheets50 " & _
                  "  SET Result = '" & m_Result & "', " & _
                  "  UserName = '" & AddTicks(m_UserName) & "', DateTimeOfRecord = '" & Format$(Now, "dd/MMM/yyyy HH:mm:ss") & "' " & _
                  "  WHERE SampleID = '" & m_SampleID & "' " & _
                  "  AND DayIndex = '" & m_DayIndex & "' "
13140         sql = sql & "ELSE " & _
                  "  INSERT INTO FaecesWorkSheets50 (SampleID, DayIndex, Result, UserName) " & _
                  "  VALUES " & _
                  "  ('" & m_SampleID & "', " & _
                  "  '" & m_DayIndex & "', " & _
                  "  '" & m_Result & "', " & _
                  "  '" & AddTicks(m_UserName) & "') "
13150         Cnxn(0).Execute sql
13160     End If


13170     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

13180     intEL = Erl
13190     strES = Err.Description
13200     LogError "FaecesWorkSheet", "Save", intEL, strES, sql

End Sub
Public Property Get SampleID() As String

13210     SampleID = m_SampleID

End Property

Public Property Get UserName() As String

13220     UserName = m_UserName

End Property
Public Property Get DateTimeOfRecord() As Date

13230     DateTimeOfRecord = m_DateTimeOfRecord

End Property
Public Property Get Result() As String

13240     Result = m_Result

End Property

Public Property Get Dayindex() As String

13250     Dayindex = m_DayIndex

End Property

Public Property Let Result(ByVal Value As String)

13260     m_Result = Value

End Property


Public Property Let DateTimeOfRecord(ByVal Value As Date)

13270     m_DateTimeOfRecord = Value

End Property
Public Property Let SampleID(ByVal Value As String)

13280     m_SampleID = Value

End Property
Public Property Let UserName(ByVal Value As String)

13290     m_UserName = Value

End Property
Public Property Let Dayindex(ByVal Value As String)

13300     m_DayIndex = Value

End Property
