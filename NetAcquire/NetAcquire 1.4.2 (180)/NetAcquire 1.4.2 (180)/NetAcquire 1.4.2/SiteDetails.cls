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
        
26230     Count = mcolSiteDetails.Count

End Function

Public Sub Delete(ByVal SampleID As String)
        
          Dim sql As String

26240     On Error GoTo Delete_Error

26250     sql = "DELETE FROM SiteDetails50 WHERE " & _
              "SampleID = '" & SampleID & "'"
26260     Cnxn(0).Execute sql

26270     Exit Sub

Delete_Error:

          Dim strES As String
          Dim intEL As Integer

26280     intEL = Erl
26290     strES = Err.Description
26300     LogError "SiteDetails", "Delete", intEL, strES, sql

End Sub

Public Sub Load(ByVal SampleID As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Ix As SiteDetail

26310     On Error GoTo Load_Error

26320     sql = "SELECT SampleID, Site, SiteDetails, " & _
              "UserName, DateTimeOfRecord " & _
              "FROM SiteDetails50 " & _
              "WHERE SampleID = '" & SampleID & "'"
26330     Set tb = New Recordset
26340     RecOpenClient 0, tb, sql
26350     Do While Not tb.EOF
26360         Set Ix = New SiteDetail
26370         With Ix
26380             .SampleID = tb!SampleID & ""
26390             .Site = tb!Site & ""
26400             .SiteDetails = tb!SiteDetails & ""
26410             .UserName = tb!UserName & ""
26420             .DateTimeOfRecord = tb!DateTimeOfRecord & ""
26430             mcolSiteDetails.Add Ix
26440         End With
26450         tb.MoveNext
26460     Loop

26470     Exit Sub

Load_Error:

          Dim strES As String
          Dim intEL As Integer

26480     intEL = Erl
26490     strES = Err.Description
26500     LogError "SiteDetails", "Load", intEL, strES, sql

End Sub
Public Function Item(ByVal Index As Integer) As SiteDetail
Attribute Item.VB_UserMemId = 0
        
26510     On Error Resume Next

26520     Set Item = mcolSiteDetails.Item(Index)

End Function




Public Sub Save(ByVal Ix As SiteDetail)

          Dim sql As String
          Dim UN As String
          Dim SD As String

26530     On Error GoTo Save_Error

26540     UN = AddTicks(Ix.UserName)

26550     If Trim$(Ix.Site) = "" Then
26560         sql = "DELETE FROM SiteDetails50 WHERE " & _
                  "SampleID = '" & Ix.SampleID & "'"
26570     Else
26580         SD = AddTicks(Ix.SiteDetails)
26590         sql = "IF EXISTS (SELECT * FROM SiteDetails50 WHERE " & _
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
26600     End If
26610     Cnxn(0).Execute sql

26620     Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

26630     intEL = Erl
26640     strES = Err.Description
26650     LogError "SiteDetails", "Save", intEL, strES, sql

End Sub

Public Property Get SiteDetails() As SiteDetails
        
26660     Set SiteDetails = mSiteDetails

End Property



