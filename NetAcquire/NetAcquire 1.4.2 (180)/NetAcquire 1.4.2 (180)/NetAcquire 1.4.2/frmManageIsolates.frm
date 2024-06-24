VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManageIsolates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Isolates"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   7860
      Picture         =   "frmManageIsolates.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5430
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdCurrent 
      Height          =   1215
      Left            =   1140
      TabIndex        =   2
      Top             =   690
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              "
   End
   Begin MSFlexGridLib.MSFlexGrid grdRepeat 
      Height          =   2175
      Left            =   1140
      TabIndex        =   6
      Top             =   4230
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              "
   End
   Begin MSFlexGridLib.MSFlexGrid grdArc 
      Height          =   2175
      Left            =   1140
      TabIndex        =   7
      Top             =   1980
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              |<Archived By |<Archived Time "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Repeats"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   4290
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "frmManageIsolates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillCurrent()

      Dim Isos As New Isolates
      Dim Iso As Isolate
      Dim s As String

35040 On Error GoTo FillCurrent_Error

35050 Isos.Load Val(lblSampleID) ' + sysOptMicroOffset(0)
35060 With grdCurrent
35070   .Rows = 2
35080   .AddItem ""
35090   .RemoveItem 1

35100   For Each Iso In Isos
35110     s = Iso.IsolateNumber & vbTab & _
              Iso.OrganismGroup & vbTab & _
              Iso.OrganismName & vbTab & _
              Iso.Qualifier
35120     .AddItem s
35130   Next
        
35140   If .Rows > 2 Then
35150     .RemoveItem 1
35160   End If
35170 End With

35180 Exit Sub

FillCurrent_Error:

      Dim strES As String
      Dim intEL As Integer

35190 intEL = Erl
35200 strES = Err.Description
35210 LogError "frmManageIsolates", "FillCurrent", intEL, strES
        
End Sub

Private Sub FillArchive()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

35220 On Error GoTo FillArchive_Error

35230 With grdArc
35240   .Rows = 2
35250   .AddItem ""
35260   .RemoveItem 1

      '60      sql = "SELECT * FROM IsolatesArc WHERE " & _
      '              "SampleID = '" & Val(lblSampleID) + sysOptMicroOffset(0) & "'"
35270   sql = "SELECT * FROM IsolatesArc WHERE " & _
              "SampleID = '" & Val(lblSampleID) & "'"
35280   Set tb = New Recordset
35290   RecOpenServer 0, tb, sql
35300   Do While Not tb.EOF
35310     s = tb!IsolateNumber & vbTab & _
              tb!OrganismGroup & vbTab & _
              tb!OrganismName & vbTab & _
              tb!Qualifier & vbTab & _
              tb!ArchivedBy & vbTab & _
              tb!ArchiveDateTime
35320     .AddItem s
35330     tb.MoveNext
35340   Loop
        
35350   If .Rows > 2 Then
35360     .RemoveItem 1
35370   End If
        
35380 End With

35390 Exit Sub

FillArchive_Error:

      Dim strES As String
      Dim intEL As Integer

35400 intEL = Erl
35410 strES = Err.Description
35420 LogError "frmManageIsolates", "FillArchive", intEL, strES, sql

        
End Sub
Private Sub FillRepeat()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

35430 On Error GoTo FillRepeat_Error

35440 With grdRepeat
35450   .Rows = 2
35460   .AddItem ""
35470   .RemoveItem 1

      '60      sql = "SELECT * FROM IsolatesRepeats WHERE " & _
      '              "SampleID = '" & Val(lblSampleID) + sysOptMicroOffset(0) & "'"
35480   sql = "SELECT * FROM IsolatesRepeats WHERE " & _
              "SampleID = '" & Val(lblSampleID) & "'"
35490   Set tb = New Recordset
35500   RecOpenServer 0, tb, sql
35510   Do While Not tb.EOF
35520     s = tb!IsolateNumber & vbTab & _
              tb!OrganismGroup & vbTab & _
              tb!OrganismName & vbTab & _
              tb!Qualifier & ""
35530     .AddItem s
35540     tb.MoveNext
35550   Loop
        
35560   If .Rows > 2 Then
35570     .RemoveItem 1
35580   End If
35590 End With

35600 Exit Sub

FillRepeat_Error:

      Dim strES As String
      Dim intEL As Integer

35610 intEL = Erl
35620 strES = Err.Description
35630 LogError "frmManageIsolates", "FillRepeat", intEL, strES, sql

        
End Sub
Private Sub cmdCancel_Click()

35640 Unload Me

End Sub


Private Sub Form_Activate()

35650 FillCurrent
35660 FillRepeat
35670 FillArchive

End Sub

