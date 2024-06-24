VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiteDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolSiteDetails As New Collection

Private mSiteDetails As New SiteDetails
Public Function Count() As Long
  
10    Count = mcolSiteDetails.Count

End Function

Public Sub Delete(ByVal SampleID As String)
  
      Dim sql As String

10    On Error GoTo Delete_Error

20    sql = "DELETE FROM SiteDetails50 WHERE " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "SiteDetails", "Delete", intEL, strES, sql

End Sub

Public Sub Load(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim Ix As SiteDetail

10    On Error GoTo Load_Error

20    sql = "SELECT SampleID, Site, SiteDetails, " & _
            "UserName, DateTimeOfRecord " & _
            "FROM SiteDetails50 " & _
            "WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Ix = New SiteDetail
70      With Ix
80        .SampleID = tb!SampleID & ""
90        .Site = tb!Site & ""
100       .SiteDetails = tb!SiteDetails & ""
110       .UserName = tb!UserName & ""
120       .DateTimeOfRecord = tb!DateTimeOfRecord & ""
130       mcolSiteDetails.Add Ix
140     End With
150     tb.MoveNext
160   Loop

170   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "SiteDetails", "Load", intEL, strES, sql

End Sub
Public Function Item(ByVal Index As Integer) As SiteDetail
Attribute Item.VB_UserMemId = 0
  
10    On Error Resume Next

20    Set Item = mcolSiteDetails.Item(Index)

End Function




Public Sub Save(ByVal Ix As SiteDetail)

      Dim sql As String
      Dim UN As String
      Dim SD As String
      
10    On Error GoTo Save_Error

20    UN = AddTicks(Ix.UserName)

30    If Trim$(Ix.Site) = "" Then
40      sql = "DELETE FROM SiteDetails50 WHERE " & _
              "SampleID = '" & Ix.SampleID & "'"
50    Else
        SD = AddTicks(Ix.SiteDetails)
60      sql = "IF EXISTS (SELECT * FROM SiteDetails50 WHERE " & _
              "           SampleID = '" & Ix.SampleID & "') " & _
              "  UPDATE SiteDetails50 " & _
              "  SET Site = '" & Ix.Site & "', " & _
              "  SiteDetails = '" & SD & "', " & _
              "  UserName = '" & UN & "' " & _
              "  WHERE SampleID = '" & Ix.SampleID & "' " & _
              "ELSE " & _
              "  INSERT INTO SiteDetails50 (SampleID, Site, SiteDetails, UserName) " & _
              "  VALUES " & _
              "  ('" & Ix.SampleID & "', " & _
              "  '" & Ix.Site & "', " & _
              "  '" & SD & "', " & _
              "  '" & UN & "')"
70    End If
80    Cnxn(0).Execute sql

90    Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "SiteDetails", "Save", intEL, strES, sql

End Sub

Public Property Get SiteDetails() As SiteDetails
  
10    Set SiteDetails = mSiteDetails

End Property



