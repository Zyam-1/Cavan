VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarUserRoleCollection As UserRoleCollection
Private m_MemberOf As String
Private m_SystemRole As String
Private m_Enabled As Byte
Private m_UserName As String
Private m_DateTimeOfRecord As Date
Private m_Description As String

Public Property Get UserRoleCollection() As UserRoleCollection
30300     If mvarUserRoleCollection Is Nothing Then
30310         Set mvarUserRoleCollection = New UserRoleCollection
30320     End If

30330     Set UserRoleCollection = mvarUserRoleCollection
End Property
Public Property Set UserRoleCollection(vData As UserRoleCollection)
30340     Set mvarUserRoleCollection = vData
End Property

Public Property Let Description(ByVal vData As String)
30350     m_Description = vData
End Property
Public Property Get Description() As String
30360     Description = m_Description
End Property
Public Property Let MemberOf(ByVal vData As String)
30370     m_MemberOf = vData
End Property
Public Property Get MemberOf() As String
30380     MemberOf = m_MemberOf
End Property
Public Property Let SystemRole(ByVal vData As String)
30390     m_SystemRole = vData
End Property
Public Property Get SystemRole() As String
30400     SystemRole = m_SystemRole
End Property
Public Property Let Enabled(ByVal vData As Byte)
30410     m_Enabled = vData
End Property
Public Property Get Enabled() As Byte
30420     Enabled = m_Enabled
End Property
Public Property Let UserName(ByVal vData As String)
30430     m_UserName = vData
End Property
Public Property Get UserName() As String
30440     UserName = m_UserName
End Property
Public Property Let DateTimeOfRecord(ByVal vData As Date)
30450     m_DateTimeOfRecord = vData
End Property
Public Property Get DateTimeOfRecord() As Date
30460     DateTimeOfRecord = m_DateTimeOfRecord
End Property

Public Function PopulateUserRole(ByVal UserRoleInstance As UserRole, ByVal tb As Recordset, ByVal CurrentUser As String) As Boolean
30470     On Error GoTo PopulateUserRole_Error

30480     With UserRoleInstance
30490         If Not IsNull(tb("MemberOf")) Then
30500             .MemberOf = tb("MemberOf")
30510         End If
30520         If Not IsNull(tb("SystemRole")) Then
30530             .SystemRole = tb("SystemRole")
30540         End If
30550         If Not IsNull(tb("Description")) Then
30560             .Description = tb("Description")
30570         End If
30580         If Not IsNull(tb("Enabled")) Then
30590             .Enabled = tb("Enabled")
30600         End If
30610         If Not IsNull(tb("Username")) Then
30620             .UserName = tb("Username")
30630         End If
30640         If Not IsNull(tb("DateTimeOfRecord")) Then
30650             .DateTimeOfRecord = tb("DateTimeOfRecord")
30660         End If
          
30670     End With
30680     PopulateUserRole = True

30690     Exit Function

PopulateUserRole_Error:

          Dim strES As String
          Dim intEL As Integer

30700     intEL = Erl
30710     strES = Err.Description
30720     LogError "UserRole", "PopulateUserRole", intEL, strES

End Function
Public Function PopulateUserRoleList(ByVal tb As Recordset, ByVal CurrentUser As String) As UserRoleCollection
          Dim UserRoleList As New UserRoleCollection
          Dim NewUserRole As UserRole
30730     On Error GoTo PopulateUserRoleList_Error

30740     If tb.EOF Then
30750         Set UserRoleList = Nothing
30760         Set PopulateUserRoleList = Nothing
30770         Exit Function
30780     End If
30790     While Not tb.EOF
30800         Set NewUserRole = New UserRole
30810         NewUserRole.PopulateUserRole NewUserRole, tb, CurrentUser
30820         UserRoleList.Add NewUserRole
30830         tb.MoveNext
30840     Wend
30850     tb.Close
30860     Set tb = Nothing
30870     Set PopulateUserRoleList = UserRoleList

30880     Exit Function

PopulateUserRoleList_Error:

          Dim strES As String
          Dim intEL As Integer

30890     intEL = Erl
30900     strES = Err.Description
30910     LogError "UserRole", "PopulateUserRoleList", intEL, strES

End Function
Public Function GetUserRoleList(ByVal MemberOf As String, ByVal CurrentUser As String) As UserRoleCollection
          Dim sql As String
          Dim tb As Recordset
30920     On Error GoTo GetUserRoleList_Error

30930     sql = "SELECT * FROM UserRole WHERE MemberOf = '" & MemberOf & "' ORDER BY SystemRole"
30940     Set tb = New Recordset
30950     RecOpenServer 0, tb, sql
30960     Set GetUserRoleList = PopulateUserRoleList(tb, CurrentUser)

30970     Exit Function

GetUserRoleList_Error:

          Dim strES As String
          Dim intEL As Integer

30980     intEL = Erl
30990     strES = Err.Description
31000     LogError "UserRole", "GetUserRoleList", intEL, strES, sql

End Function
Public Function GetUserRole(ByVal MemberOf As String, ByVal SystemRole As String, CurrentUser As String) As Boolean
          Dim sql As String
          Dim tb As Recordset
31010     On Error GoTo GetUserRole_Error

31020     sql = "SELECT * FROM UserRole WHERE MemberOf = '" & MemberOf & "' AND SystemRole = '" & SystemRole & "'"
31030     Set tb = New Recordset
31040     RecOpenServer 0, tb, sql
31050     If tb.EOF Then
31060         GetUserRole = False
31070         Exit Function
31080     End If
31090     PopulateUserRole Me, tb, CurrentUser
31100     tb.Close
31110     Set tb = Nothing
31120     GetUserRole = True

31130     Exit Function

GetUserRole_Error:

          Dim strES As String
          Dim intEL As Integer

31140     intEL = Erl
31150     strES = Err.Description
31160     LogError "UserRole", "GetUserRole", intEL, strES, sql

End Function

Public Function Update(ByVal MemberOf As String, ByVal SystemRole As String, ByVal Value As Byte, ByVal CurrentUser As String) As Boolean

          Dim sql As String

31170     On Error GoTo Save_Error

31180     sql = "UPDATE UserRole SET Username = '" & UserName & "', Enabled = " & Value & _
              " WHERE SystemRole = '" & SystemRole & "' AND MemberOf = '" & MemberOf & "'"
31190     Cnxn(0).Execute sql
31200     Update = True
       
31210     Exit Function

Save_Error:

          Dim strES As String
          Dim intEL As Integer

31220     intEL = Erl
31230     strES = Err.Description
31240     LogError "UserRole", "Save", intEL, strES, sql

End Function
Public Function Add(ByVal CurrentUser As String) As Boolean

          Dim sql As String

31250     On Error GoTo Add_Error

31260     sql = "INSERT INTO [dbo].[UserRole] ([MemberOf],[SystemRole],[Description],[Enabled],[Username],[DateTimeOfRecord]) " & _
              "VALUES " & _
              "('" & MemberOf & "','" & SystemRole & "','" & Description & "'," & Enabled & ",'" & CurrentUser & "',GETDATE())"

31270     Cnxn(0).Execute sql

31280     Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer

31290     intEL = Erl
31300     strES = Err.Description
31310     LogError "UserRole", "Add", intEL, strES, sql

End Function


