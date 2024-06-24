VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ABDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolABDefinitions As New Collection

Private mABDefinitions As New ABDefinitions

Public Function Count() As Long
  
10    Count = mcolABDefinitions.Count

End Function




Public Sub Delete(ByVal AB As ABDefinition)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM ABDefinitions WHERE " & _
            "AntibioticName = '" & AB.AntibioticName & "' " & _
            "AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
            "AND Site = '" & AB.Site & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "ABDefinitions", "Delete", intEL, strES, sql

End Sub

Public Property Get ABDefinitions() As ABDefinitions
  
10    Set ABDefinitions = mABDefinitions

End Property


Public Function Item(ByVal Site As String, ByVal OrganismGroup As String, ByVal AntibioticName As String) As ABDefinition
  

Dim I As Integer

On Error GoTo Item_Error

Set Item = Nothing
For I = 1 To mcolABDefinitions.Count
    If mcolABDefinitions.Item(I).Site = Site And mcolABDefinitions.Item(I).OrganismGroup = OrganismGroup _
        And mcolABDefinitions.Item(I).AntibioticName = AntibioticName Then
        Set Item = mcolABDefinitions(I)
    End If
Next I


Exit Function

Item_Error:

 Dim strES As String
 Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "ABDefinitions", "Item", intEL, strES
    
End Function


Public Sub Load(ByVal Site As String, ByVal OrganismGroup As String)

      Dim tb As Recordset
      Dim sql As String
      Dim AB As ABDefinition
      Dim sx As Sensitivity

10    On Error GoTo Load_Error

20    sql = "SELECT A.Code, D.AntibioticName, D.OrganismGroup, D.Site, D.ListOrder, D.PriSec, " & _
            "COALESCE(A.AllowIfPregnant, 0) AllowIfPregnant, " & _
            "COALESCE(A.AllowIfOutPatient, 0) AllowIfOutPatient, " & _
            "COALESCE(A.AllowIfChild, 0) AllowIfChild, " & _
            "COALESCE(A.AllowIfPenAll, 0) AllowIfPenAll, " & _
            "COALESCE(A.ViewInGrid, 0) ViewInGrid " & _
            "FROM ABDefinitions D JOIN Antibiotics A " & _
            "ON D.AntibioticName = A.AntibioticName " & _
            "WHERE D.OrganismGroup = '" & OrganismGroup & "' " & _
            "AND D.Site = '" & Site & "' " & _
            "ORDER BY D.ListOrder"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60    sql = "SELECT A.Code, D.AntibioticName, D.OrganismGroup, D.Site, D.ListOrder, D.PriSec, " & _
            "COALESCE(A.AllowIfPregnant, 0) AllowIfPregnant, " & _
            "COALESCE(A.AllowIfOutPatient, 0) AllowIfOutPatient, " & _
            "COALESCE(A.AllowIfChild, 0) AllowIfChild, " & _
            "COALESCE(A.AllowIfPenAll, 0) AllowIfPenAll, " & _
            "COALESCE(A.ViewInGrid, 0) ViewInGrid " & _
            "FROM ABDefinitions D JOIN Antibiotics A " & _
            "ON D.AntibioticName = A.AntibioticName " & _
            "WHERE D.OrganismGroup = '" & OrganismGroup & "' " & _
            "AND D.Site = 'Generic' " & _
            "ORDER BY D.ListOrder"
70      Set tb = New Recordset
80      RecOpenClient 0, tb, sql
90    End If
100   Do While Not tb.EOF
110     Set AB = New ABDefinition
120     With AB
130       .Code = tb!Code & ""
140       .AntibioticName = tb!AntibioticName & ""
150       .OrganismGroup = tb!OrganismGroup & ""
160       .Site = tb!Site & ""

170       .ListOrder = tb!ListOrder
180       .PriSec = tb!PriSec & ""
190       .AllowIfPregnant = tb!AllowIfPregnant
200       .AllowIfOutPatient = tb!AllowIfOutPatient
210       .AllowIfChild = tb!AllowIfChild
220       .AllowIfPenicillinAllergy = tb!AllowIfPenAll
230       .ViewInGrid = tb!ViewInGrid
240       mcolABDefinitions.Add AB, .Site & .OrganismGroup & .AntibioticName
250     End With
260     tb.MoveNext
270   Loop

280   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "ABDefinitions", "Load", intEL, strES, sql

End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolABDefinitions.[_NewEnum]

End Function

Public Sub Save(ByVal AB As ABDefinition)

      Dim sql As String

10    On Error GoTo Save_Error

20    sql = "IF EXISTS (SELECT * FROM ABDefinitions WHERE " & _
            "           AntibioticName = '" & AB.AntibioticName & "' " & _
            "           AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
            "           AND Site = '" & AB.Site & "') " & _
            "  UPDATE ABDefinitions " & _
            "  SET ListOrder = '" & AB.ListOrder & "', " & _
            "  PriSec = '" & AB.PriSec & "' " & _
            "  WHERE AntibioticName = '" & AB.AntibioticName & "' " & _
            "  AND OrganismGroup = '" & AB.OrganismGroup & "' " & _
            "  AND Site = '" & AB.Site & "' "
30    sql = sql & "ELSE " & _
            "  INSERT INTO ABDefinitions (AntibioticName, OrganismGroup, Site, ListOrder, PriSec ) " & _
            "  VALUES " & _
            "  ('" & AB.AntibioticName & "', " & _
            "  '" & AB.OrganismGroup & "', " & _
            "  '" & AB.Site & "', " & _
            "  '" & AB.ListOrder & "', " & _
            "  '" & AB.PriSec & "')"
40    Cnxn(0).Execute sql

50    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "ABDefinitions", "Save", intEL, strES, sql

End Sub


Public Sub SaveAllowable(ByVal AB As ABDefinition)

      Dim sql As String

10    On Error GoTo SaveAllowable_Error

20    sql = "IF EXISTS (SELECT * FROM Antibiotics WHERE " & _
            "           AntibioticName = '" & AB.AntibioticName & "') " & _
            "  UPDATE Antibiotics " & _
            "  SET ListOrder = '" & AB.ListOrder & "', " & _
            "  AllowIfPregnant = '" & AB.AllowIfPregnant & "', " & _
            "  AllowIfOutPatient = '" & AB.AllowIfOutPatient & "', " & _
            "  AllowIfChild = '" & AB.AllowIfChild & "', " & _
            "  AllowIfPenAll = '" & AB.AllowIfPenicillinAllergy & "' " & _
            "  WHERE AntibioticName = '" & AB.AntibioticName & "' "
30    sql = sql & "ELSE " & _
            "  INSERT INTO Antibiotics (Code, AntibioticName, ListOrder, AllowIfPregnant, AllowIfOutPatient, AllowIfChild, AllowIfPenAll)" & _
            "  VALUES " & _
            "  ('" & AB.Code & "', " & _
            "  '" & AB.AntibioticName & "', " & _
            "  '" & AB.ListOrder & "', " & _
            "  '" & AB.AllowIfPregnant & "', " & _
            "  '" & AB.AllowIfOutPatient & "', " & _
            "  '" & AB.AllowIfChild & "', " & _
            "  '" & AB.AllowIfPenicillinAllergy & "')"
40    Cnxn(0).Execute sql

50    Exit Sub

SaveAllowable_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "ABDefinitions", "SaveAllowable", intEL, strES, sql

End Sub


