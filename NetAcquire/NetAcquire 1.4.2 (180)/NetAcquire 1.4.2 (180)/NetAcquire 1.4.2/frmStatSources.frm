VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatSources 
   Caption         =   "NetAcquire - Statistics"
   ClientHeight    =   8355
   ClientLeft      =   270
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   7995
   Begin VB.CommandButton cmdGraph 
      Appearance      =   0  'Flat
      Caption         =   "Graph"
      Height          =   800
      Left            =   4890
      Picture         =   "frmStatSources.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   180
      Width           =   900
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   800
      Left            =   3930
      Picture         =   "frmStatSources.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   180
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   7980
      Visible         =   0   'False
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Wards"
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      Top             =   150
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Clinicians"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   7
      Top             =   450
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "GPs"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   6
      Top             =   750
      Width           =   1035
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   2970
      Picture         =   "frmStatSources.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   900
   End
   Begin VB.CommandButton bCalc 
      Caption         =   "Calculate"
      Height          =   800
      Left            =   2010
      Picture         =   "frmStatSources.frx":183E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   885
      Left            =   240
      TabIndex        =   1
      Top             =   90
      Width           =   1695
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   510
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   37606
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   37606
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6645
      Left            =   240
      TabIndex        =   0
      Top             =   1350
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   5
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
      FormatString    =   "<Source                 |<Total Samples |<Coag Samples |<Bio Samples |<Haem Samples    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3720
      TabIndex        =   11
      Top             =   1020
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmStatSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim tb1 As Recordset
      Dim tbD As Recordset
      Dim SourcePanelType As String
      Dim s As String
      Dim total As Long
      Dim TotCoag As Long
      Dim TotHaem As Long
      Dim TotBio As Long

64340 On Error GoTo FillG_Error

64350 If oSource(0) Then
64360   SourcePanelType = "W"
64370 ElseIf oSource(1) Then
64380   SourcePanelType = "C"
64390 ElseIf oSource(2) Then
64400   SourcePanelType = "G"
64410 End If

64420 g.Rows = 2
64430 g.AddItem ""
64440 g.RemoveItem 1
64450 g.Visible = False

64460 sql = "Select distinct SourcePanelName from SourcePanels where " & _
            "SourcePanelType = '" & SourcePanelType & "'"
64470 Set tb = New Recordset
64480 RecOpenClient 0, tb, sql
           
64490 pb.Visible = True
64500 pb.max = tb.RecordCount + 2
64510 pb = 0

64520 Do While Not tb.EOF
64530   pb = pb + 1
64540   total = 0
64550   TotCoag = 0
64560   TotHaem = 0
64570   TotBio = 0
64580   s = tb!SourcePanelName & vbTab
64590   sql = "Select Content from SourcePanels where " & _
              "SourcePanelName = '" & tb!SourcePanelName & "' " & _
              "and SourcePanelType = '" & SourcePanelType & "'"
64600   Set tb1 = New Recordset
64610   RecOpenClient 0, tb1, sql
64620   Do While Not tb1.EOF
          'Totals
64630     sql = "Select Count (*) as Tot from Demographics where "
64640     Select Case SourcePanelType
            Case "W": sql = sql & "Ward"
64650       Case "C": sql = sql & "Clinician"
64660       Case "G": sql = sql & "GP"
64670     End Select
64680     sql = sql & " = '" & AddTicks(tb1!Content & "") & "' " & _
                      "and RunDate between '" & _
                      Format$(dtFrom, "dd/mmm/yyyy") & _
                      "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
64690     Set tbD = New Recordset
64700     RecOpenClient 0, tbD, sql
64710     total = total + tbD!Tot
          
          'Coag
64720     sql = "Select distinct Demographics.SampleID from Demographics, CoagResults where "
64730     Select Case SourcePanelType
            Case "W": sql = sql & "Ward"
64740       Case "C": sql = sql & "Clinician"
64750       Case "G": sql = sql & "GP"
64760     End Select
64770     sql = sql & " = '" & AddTicks(tb1!Content & "") & "' " & _
                      "and Demographics.SampleID = CoagResults.SampleID " & _
                      "and Demographics.RunDate between '" & _
                      Format$(dtFrom, "dd/mmm/yyyy") & _
                      "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
64780     Set tbD = New Recordset
64790     RecOpenClient 0, tbD, sql
64800     If Not tbD.EOF Then
64810       TotCoag = TotCoag + tbD.RecordCount
64820     End If
          
          'Bio
64830     sql = "Select distinct Demographics.SampleID from Demographics, BioResults where "
64840     Select Case SourcePanelType
            Case "W": sql = sql & "Ward"
64850       Case "C": sql = sql & "Clinician"
64860       Case "G": sql = sql & "GP"
64870     End Select
64880     sql = sql & " = '" & AddTicks(tb1!Content & "") & "' " & _
                      "and Demographics.SampleID = BioResults.SampleID " & _
                      "and Demographics.RunDate between '" & _
                      Format$(dtFrom, "dd/mmm/yyyy") & _
                      "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
64890     Set tbD = New Recordset
64900     RecOpenClient 0, tbD, sql
64910     If Not tbD.EOF Then
64920       TotBio = TotBio + tbD.RecordCount
64930     End If
          
          'Haem
64940     sql = "Select distinct Demographics.SampleID from Demographics, HaemResults where "
64950     Select Case SourcePanelType
            Case "W": sql = sql & "Ward"
64960       Case "C": sql = sql & "Clinician"
64970       Case "G": sql = sql & "GP"
64980     End Select
64990     sql = sql & " = '" & AddTicks(tb1!Content & "") & "' " & _
                      "and Demographics.SampleID = HaemResults.SampleID " & _
                      "and Demographics.RunDate between '" & _
                      Format$(dtFrom, "dd/mmm/yyyy") & _
                      "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
65000     Set tbD = New Recordset
65010     RecOpenClient 0, tbD, sql
65020     If Not tbD.EOF Then
65030       TotHaem = TotHaem + tbD.RecordCount
65040     End If
          
65050     tb1.MoveNext
65060   Loop
65070   s = s & Format$(total) & vbTab & _
                Format$(TotCoag) & vbTab & _
                Format$(TotBio) & vbTab & _
                Format$(TotHaem)
        
65080   g.AddItem s
65090   tb.MoveNext
65100 Loop

65110 pb.Visible = False
65120 g.Visible = True
65130 If g.Rows > 2 Then
65140   g.RemoveItem 1
65150 End If


65160 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

65170 intEL = Erl
65180 strES = Err.Description
65190 LogError "fStatSources", "FillG", intEL, strES, sql


End Sub

Private Sub bCalc_Click()

65200 FillG

End Sub

Private Sub bcancel_Click()

65210 Unload Me

End Sub


Private Sub cmdGraph_Click()

          Dim i As Integer

65220     On Error GoTo cmdGraph_Click_Error
65230     If g.TextMatrix(1, 0) = "" Then Exit Sub
65240     With frmGraph

65250         .GraphTitleText = "General Chemistry Statistics - Between " & dtFrom.Value & " and " & dtTo.Value
65260         .GraphFootNoteText = "Counts By "
65270         If oSource(0).Value = True Then
65280             .GraphFootNoteText = .GraphFootNoteText & oSource(0).Caption
65290         ElseIf oSource(0).Value = True Then
65300             .GraphFootNoteText = .GraphFootNoteText & oSource(1).Caption
65310         ElseIf oSource(0).Value = True Then
65320             .GraphFootNoteText = .GraphFootNoteText & oSource(2).Caption
65330         End If
65340         .g.Cols = 4
65350         .g.TextMatrix(0, 1) = "Caogulation"
65360         .g.TextMatrix(0, 2) = "Biochemistry"
65370         .g.TextMatrix(0, 3) = "Haematology"

65380         For i = 1 To g.Rows - 1
65390             .g.AddItem g.TextMatrix(i, 0) & vbTab & g.TextMatrix(i, 2) & vbTab & g.TextMatrix(i, 3) & vbTab & g.TextMatrix(i, 4)
65400         Next
65410         .g.RemoveItem 1
65420         .Show 1
65430     End With


65440     Exit Sub

cmdGraph_Click_Error:

          Dim strES As String
          Dim intEL As Integer

65450     intEL = Erl
65460     strES = Err.Description
65470     LogError "frmMonthlyResultsDateWise", "cmdGraph_Click", intEL, strES

End Sub

Private Sub cmdXL_Click()

65480 ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

65490 dtTo = Format$(Now, "dd/mm/yyyy")
65500 dtFrom = Format$(Now - 365, "dd/mm/yyyy")

End Sub


