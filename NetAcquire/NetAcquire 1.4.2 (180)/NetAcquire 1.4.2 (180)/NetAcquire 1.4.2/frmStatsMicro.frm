VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatsMicro 
   Caption         =   "NetAcquire - Microbiology"
   ClientHeight    =   7065
   ClientLeft      =   1425
   ClientTop       =   345
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10740
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   8610
      TabIndex        =   5
      Top             =   1140
      Width           =   2025
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Text            =   "cmbMonth"
         Top             =   510
         Width           =   1365
      End
      Begin VB.CommandButton breCalc 
         Caption         =   "Calculate"
         Height          =   945
         Left            =   300
         Picture         =   "frmStatsMicro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   870
         Width           =   1365
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   180
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2005
         Alignment       =   0
         BuddyControl    =   "lblYear"
         BuddyDispid     =   196612
         OrigLeft        =   900
         OrigTop         =   180
         OrigRight       =   1515
         OrigBottom      =   495
         Max             =   2015
         Min             =   2000
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   1755
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   3096
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   1755
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   3096
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2005"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   9
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Locations"
      Height          =   405
      Left            =   9030
      TabIndex        =   4
      Top             =   270
      Width           =   1185
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   945
      Left            =   9030
      Picture         =   "frmStatsMicro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   945
      Left            =   9030
      Picture         =   "frmStatsMicro.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5550
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11668
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Site                                      |"
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8550
      Picture         =   "frmStatsMicro.frx":0C7E
      Top             =   240
      Width           =   480
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
      Left            =   8970
      TabIndex        =   3
      Top             =   4350
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmStatsMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StartDate As String
Private StopDate As String

Private Type Source
  SourceType As String
  Name As String
  sql As String
End Type
Private Sources() As Source
Private Sub BuildSourceSelections()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim X As Integer
      Dim WCG As String

63200 On Error GoTo BuildSourceSelections_Error

63210 If Not IsNumeric(UBound(Sources)) Then
63220   Exit Sub
63230 End If

63240 For X = 0 To UBound(Sources)

63250   sql = "Select Content from SourcePanels where " & _
              "SourcePanelName = '" & Sources(X).Name & "'"
        
63260   Set tb = New Recordset
63270   RecOpenServer 0, tb, sql
        
63280   Select Case Sources(X).SourceType
          Case "W": WCG = "Ward"
63290     Case "C": WCG = "Clinician"
63300     Case "G": WCG = "GP"
63310   End Select
63320   s = "("
        
63330   Do While Not tb.EOF
63340     s = s & WCG & " = '" & Trim$(tb!Content & "") & "' or "
63350     tb.MoveNext
63360   Loop
63370   s = Left$(s, Len(s) - 3) & ")"
63380   Sources(X).sql = s
        
63390 Next

63400 Exit Sub

BuildSourceSelections_Error:

      Dim strES As String
      Dim intEL As Integer

63410 intEL = Erl
63420 strES = Err.Description
63430 LogError "frmStatsMicro", "BuildSourceSelections", intEL, strES, sql


End Sub


Private Sub FillSources()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim X As Integer

63440 On Error GoTo FillSources_Error

63450 sql = "Select Distinct SourcePanelName, SourcePanelType from SourcePanels where " & _
            "SourcePanelName is not null "
63460 Set tb = New Recordset
63470 RecOpenServer 0, tb, sql

63480 s = "<Site                                      |<Total   "
63490 X = 0
63500 Do While Not tb.EOF
        
63510   ReDim Preserve Sources(0 To X)
63520   Sources(X).Name = Trim$(tb!SourcePanelName & "")
63530   Sources(X).SourceType = Trim$(tb!SourcePanelType & "")
63540   X = X + 1
        
63550   s = s & "|<" & Trim$(tb!SourcePanelName) & " (" & Trim$(tb!SourcePanelType) & ")    "
63560   tb.MoveNext
        
63570 Loop

63580 g.FormatString = s

63590 BuildSourceSelections

63600 Exit Sub

FillSources_Error:

      Dim strES As String
      Dim intEL As Integer

63610 intEL = Erl
63620 strES = Err.Description
63630 LogError "frmStatsMicro", "FillSources", intEL, strES, sql


End Sub

Private Sub breCalc_Click()

      Dim tbSites As Recordset
      Dim tbCount As Recordset
      Dim X As Integer
      Dim sql As String
      Dim s As String
      Dim DateLimit As String
      Dim LineFound As Boolean

63640 On Error GoTo breCalc_Click_Error

63650 GetDates cmbMonth.ListIndex + 1
63660 DateLimit = " Rundate between '" & StartDate & "' and '" & StopDate & "' "

63670 g.Rows = 2
63680 g.AddItem ""
63690 g.RemoveItem 1

63700 If g.Cols = 2 Then
63710   iMsg "No Locations Set!", vbExclamation
63720   frmSetSources.Show 1
63730   FillSources
63740   Exit Sub
63750 End If

63760 pb(1) = 0
63770 pb(1).max = g.Cols - 2
63780 pb(1).Visible = True

63790 sql = "select distinct site from SiteDetails50 "
63800 Set tbSites = New Recordset
63810 RecOpenClient 0, tbSites, sql
63820 pb(0) = 0
63830 pb(0).max = tbSites.RecordCount
63840 pb(0).Visible = True
63850 Do While Not tbSites.EOF
63860   pb(0) = pb(0) + 1
63870   s = tbSites!Site & ""
63880   For X = -1 To g.Cols - 3
63890     pb(1) = X + 1
63900     If X = -1 Then
63910       sql = "Select Count (*) as Tot from Demographics as D where " & _
                  DateLimit & _
                  "and D.SampleID in " & _
                  " (Select SampleID from SiteDetails50 where " & _
                  "  Site = '" & tbSites!Site & "' )"
63920       LineFound = True
63930     Else
63940       sql = "Select Count (*) as Tot from Demographics as D where " & _
                  DateLimit & _
                  "and " & Sources(X).sql & _
                  "and D.SampleID in " & _
                  " (Select SampleID from SiteDetails50 where " & _
                  "  Site = '" & tbSites!Site & "' )"
63950     End If
63960     Set tbCount = New Recordset
63970     RecOpenServer 0, tbCount, sql
63980     If X = 0 And tbCount!Tot = 0 Then
63990       LineFound = False
64000       Exit For
64010     End If
64020     s = s & vbTab & tbCount!Tot
64030   Next
64040   If LineFound Then
64050     g.AddItem s
64060   End If
64070   tbSites.MoveNext
64080 Loop

64090 If g.Rows > 2 Then
64100   g.RemoveItem 1
64110 End If

64120 pb(0).Visible = False
64130 pb(1).Visible = False

64140 Exit Sub

breCalc_Click_Error:

      Dim strES As String
      Dim intEL As Integer

64150 intEL = Erl
64160 strES = Err.Description
64170 LogError "frmStatsMicro", "breCalc_Click", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

64180 Unload Me

End Sub

Private Sub cmbMonth_KeyPress(KeyAscii As Integer)

64190 KeyAscii = 0

End Sub


Private Sub cmdXL_Click()

64200 ExportFlexGrid g, Me

End Sub

Private Sub Command1_Click()

64210 frmSetSources.Show 1

64220 FillSources

64230 g.Rows = 2
64240 g.AddItem ""
64250 g.RemoveItem 1

End Sub

Private Sub Form_Load()

      Dim n As Integer

64260 lblYear = Format$(Now, "yyyy")

64270 For n = 1 To 12
64280   cmbMonth.AddItem Format$("01/" & Format$(n) & "/2005", "mmmm")
64290 Next
64300 cmbMonth.ListIndex = Month(Now) - 1

64310 FillSources

End Sub


Private Sub GetDates(ByVal MonthNum As Integer)

64320 StartDate = Format$("01/" & Format$(MonthNum) & "/" & lblYear, "dd/mmm/yyyy")
64330 StopDate = Format$(DateAdd("m", 1, StartDate) - 1, "dd/mmm/yyyy")

End Sub

