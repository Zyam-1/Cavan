VERSION 5.00
Begin VB.Form frmDuplicates 
   Caption         =   "NetAcquire"
   ClientHeight    =   2400
   ClientLeft      =   3120
   ClientTop       =   3150
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   405
      Left            =   1110
      TabIndex        =   0
      Top             =   1620
      Width           =   1725
   End
   Begin VB.Label lblDuplicates 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   990
      Width           =   3405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Duplicates"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   750
   End
   Begin VB.Label lblTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   540
      Width           =   3405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Table"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   570
      Width           =   405
   End
End
Attribute VB_Name = "frmDuplicates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoBio()

          Dim sql As String
          Dim tb As Recordset
          Dim tbD As Recordset
          Dim SampleID As String
          Dim Code As String
          Dim f As Long
35340     On Error GoTo DoBio_Error

35350     ReDim FieldSave(0 To 0) As Variant

35360     lblTable = "Opening BioResults"
35370     lblTable.Refresh
35380     sql = "Select a.* from BioResults as a " & _
              "Join " & _
              "  (select b.sampleid, b.code " & _
              "   from bioresults as b " & _
              "   group by b.sampleid, b.code " & _
              "   having count(*) > 1) as b " & _
              "on a.sampleid = b.sampleid " & _
              "and a.code = b.code " & _
              "order by a.sampleid,a.code"
35390     Set tb = New Recordset
35400     RecOpenServer 0, tb, sql
35410     lblTable = "BioResults"
35420     lblTable.Refresh
35430     Do While Not tb.EOF

35440         SampleID = tb!SampleID & ""
35450         Code = tb!Code & ""
        
35460         lblDuplicates = SampleID & " " & Code
35470         lblDuplicates.Refresh
        
35480         sql = "Select top 1 * from BioResults where " & _
                  "SampleID = '" & SampleID & "' " & _
                  "and Code = '" & Code & "'"
35490         Set tbD = New Recordset
35500         RecOpenClient 0, tbD, sql
35510         ReDim FieldSave(0 To tbD.Fields.Count - 1)
35520         For f = 0 To tbD.Fields.Count - 1
35530             FieldSave(f) = tbD(f)
35540         Next
35550         tbD.Close
        
35560         sql = "Delete from BioResults where SampleID = '" & SampleID & "' and Code = '" & Code & "'"
35570         Cnxn(0).Execute sql
        
35580         sql = "Select * from BioResults where SampleID = '-1'"
35590         Set tbD = New Recordset
35600         RecOpenServer 0, tbD, sql
35610         tbD.AddNew
35620         For f = 0 To tbD.Fields.Count - 1
35630             tbD(f) = FieldSave(f)
35640         Next
35650         tbD.Update
        
35660         tb.MoveNext
        
35670     Loop

35680     lblTable = "Finished"
35690     lblTable.Refresh
35700     lblDuplicates = ""
35710     lblDuplicates.Refresh

35720     Exit Sub

DoBio_Error:

          Dim strES As String
          Dim intEL As Integer

35730     intEL = Erl
35740     strES = Err.Description
35750     LogError "frmDuplicates", "DoBio", intEL, strES, sql


End Sub
Private Sub DoCoag()

          Dim sql As String
          Dim tb As Recordset
          Dim tbD As Recordset
          Dim SampleID As String
          Dim Code As String
          Dim f As Long
35760     On Error GoTo DoCoag_Error

35770     ReDim FieldSave(0 To 0) As Variant

35780     lblTable = "Opening CoagResults"
35790     lblTable.Refresh
35800     sql = "Select a.* from CoagResults as a " & _
              "Join " & _
              "  (select b.SampleID, b.Code " & _
              "   from CoagResults as b " & _
              "   group by b.SampleID, b.Code " & _
              "   having Count(*) > 1) as b " & _
              "on a.SampleID = b.SampleID " & _
              "and a.Code = b.Code " & _
              "order by a.SampleID, a.Code"
35810     Set tb = New Recordset
35820     RecOpenServer 0, tb, sql
35830     lblTable = "CoagResults"
35840     lblTable.Refresh
35850     Do While Not tb.EOF

35860         SampleID = tb!SampleID & ""
35870         Code = tb!Code & ""
        
35880         lblDuplicates = SampleID & " " & Code
35890         lblDuplicates.Refresh
        
35900         sql = "Select top 1 * from CoagResults where " & _
                  "SampleID = '" & SampleID & "' " & _
                  "and Code = '" & Code & "'"
35910         Set tbD = New Recordset
35920         RecOpenClient 0, tbD, sql
35930         ReDim FieldSave(0 To tbD.Fields.Count - 1)
35940         For f = 0 To tbD.Fields.Count - 1
35950             FieldSave(f) = tbD(f)
35960         Next
35970         tbD.Close
        
35980         sql = "Delete from CoagResults where " & _
                  "SampleID = '" & SampleID & "' " & _
                  "and Code = '" & Code & "'"
35990         Cnxn(0).Execute sql
        
36000         sql = "Select * from CoagResults where SampleID = '-1'"
36010         Set tbD = New Recordset
36020         RecOpenServer 0, tbD, sql
36030         tbD.AddNew
36040         For f = 0 To tbD.Fields.Count - 1
36050             tbD(f) = FieldSave(f)
36060         Next
36070         tbD.Update
        
36080         tb.MoveNext
        
36090     Loop

36100     lblTable = "Finished"
36110     lblTable.Refresh
36120     lblDuplicates = ""
36130     lblDuplicates.Refresh

36140     Exit Sub

DoCoag_Error:

          Dim strES As String
          Dim intEL As Integer

36150     intEL = Erl
36160     strES = Err.Description
36170     LogError "frmDuplicates", "DoCoag", intEL, strES, sql


End Sub

Private Sub DoHaem()

          Dim sql As String
          Dim tb As Recordset
          Dim tbD As Recordset
          Dim SampleID As String
          Dim f As Long
36180     On Error GoTo DoHaem_Error

36190     ReDim FieldSave(0 To 0) As Variant

36200     lblTable = "Opening HaemResults"
36210     lblTable.Refresh
36220     sql = "Select a.* from HaemResults as a " & _
              "Join " & _
              "  (select b.SampleID " & _
              "   from HaemResults as b " & _
              "   group by b.SampleID " & _
              "   having count(*) > 1) as b " & _
              "on a.SampleID = b.SampleID " & _
              "order by a.SampleID"
36230     Set tb = New Recordset
36240     RecOpenServer 0, tb, sql
36250     lblTable = "HaemResults"
36260     lblTable.Refresh
36270     Do While Not tb.EOF

36280         SampleID = tb!SampleID & ""
        
36290         lblDuplicates = SampleID
36300         lblDuplicates.Refresh
        
36310         sql = "Select top 1 * from HaemResults where " & _
                  "SampleID = '" & SampleID & "'"
36320         Set tbD = New Recordset
36330         RecOpenClient 0, tbD, sql
36340         ReDim FieldSave(0 To tbD.Fields.Count - 1)
36350         For f = 0 To tbD.Fields.Count - 1
36360             FieldSave(f) = tbD(f)
36370         Next
36380         tbD.Close
        
36390         sql = "Delete from HaemResults where SampleID = '" & SampleID & "'"
36400         Cnxn(0).Execute sql
        
36410         sql = "Select * from HaemResults where SampleID = '-1'"
36420         Set tbD = New Recordset
36430         RecOpenServer 0, tbD, sql
36440         tbD.AddNew
36450         For f = 0 To tbD.Fields.Count - 1
36460             tbD(f) = FieldSave(f)
36470         Next
36480         tbD.Update
        
36490         tb.MoveNext
        
36500     Loop

36510     lblTable = "Finished"
36520     lblTable.Refresh
36530     lblDuplicates = ""
36540     lblDuplicates.Refresh

36550     Exit Sub

DoHaem_Error:

          Dim strES As String
          Dim intEL As Integer

36560     intEL = Erl
36570     strES = Err.Description
36580     LogError "frmDuplicates", "DoHaem", intEL, strES, sql


End Sub

Private Sub DoDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim tbD As Recordset
          Dim SampleID As String
          Dim f As Long
36590     On Error GoTo DoDemographics_Error

36600     ReDim FieldSave(0 To 0) As Variant

36610     lblTable = "Opening Demographics"
36620     lblTable.Refresh
36630     sql = "Select a.* from Demographics as a " & _
              "Join " & _
              "  (select b.SampleID " & _
              "   from Demographics as b " & _
              "   group by b.SampleID " & _
              "   having count(*) > 1) as b " & _
              "on a.SampleID = b.SampleID " & _
              "order by a.SampleID"
36640     Set tb = New Recordset
36650     RecOpenServer 0, tb, sql
36660     lblTable = "Demographics"
36670     lblTable.Refresh
36680     Do While Not tb.EOF

36690         SampleID = tb!SampleID & ""
        
36700         lblDuplicates = SampleID
36710         lblDuplicates.Refresh
        
36720         sql = "Select top 1 * from Demographics where " & _
                  "SampleID = '" & SampleID & "'"
36730         Set tbD = New Recordset
36740         RecOpenClient 0, tbD, sql
36750         ReDim FieldSave(0 To tbD.Fields.Count - 1)
36760         For f = 0 To tbD.Fields.Count - 1
36770             FieldSave(f) = tbD(f)
36780         Next
36790         tbD.Close
        
36800         sql = "Delete from Demographics where SampleID = '" & SampleID & "'"
36810         Cnxn(0).Execute sql
        
36820         sql = "Select * from Demographics where SampleID = '-1'"
36830         Set tbD = New Recordset
36840         RecOpenServer 0, tbD, sql
36850         tbD.AddNew
36860         For f = 0 To tbD.Fields.Count - 1
36870             tbD(f) = FieldSave(f)
36880         Next
36890         tbD.Update
        
36900         tb.MoveNext
        
36910     Loop

36920     lblTable = "Finished"
36930     lblTable.Refresh
36940     lblDuplicates = ""
36950     lblDuplicates.Refresh

36960     Exit Sub

DoDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

36970     intEL = Erl
36980     strES = Err.Description
36990     LogError "frmDuplicates", "DoDemographics", intEL, strES, sql


End Sub



Private Sub cmdGo_Click()

37000     DoDemographics
37010     DoBio
37020     DoHaem
37030     DoCoag

End Sub

