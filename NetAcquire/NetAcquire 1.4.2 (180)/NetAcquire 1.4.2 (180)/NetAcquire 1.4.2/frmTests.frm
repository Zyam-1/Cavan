VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTests 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Test Count"
   ClientHeight    =   5190
   ClientLeft      =   315
   ClientTop       =   585
   ClientWidth     =   11730
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5190
   ScaleWidth      =   11730
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4230
      Picture         =   "frmTests.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3780
      Width           =   825
   End
   Begin VB.PictureBox SSPanel1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   5295
      TabIndex        =   4
      Top             =   150
      Width           =   5355
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   315
         Left            =   270
         TabIndex        =   16
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218169345
         CurrentDate     =   36942
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   315
         Left            =   1950
         TabIndex        =   13
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219348993
         CurrentDate     =   36942
      End
      Begin VB.CommandButton bReCalc 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   3870
         Picture         =   "frmTests.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1050
         TabIndex        =   11
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2010
         TabIndex        =   10
         Top             =   1410
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2010
         TabIndex        =   9
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2010
         TabIndex        =   8
         Top             =   900
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   7
         Top             =   1680
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   690
         TabIndex        =   6
         Top             =   1410
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   750
         TabIndex        =   5
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2010
         TabIndex        =   15
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   240
         Width           =   345
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   330
      TabIndex        =   3
      Top             =   2400
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4875
      Left            =   5550
      TabIndex        =   2
      Top             =   180
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   8599
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<                      |<   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   420
      Picture         =   "frmTests.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1710
      Picture         =   "frmTests.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   3990
      TabIndex        =   18
      Top             =   4620
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strInHouse As String
Private strExternal As String
Private strOPD As String

Private Sub cmdXL_Click()

19040 ExportFlexGrid g, Me

End Sub


Private Sub obetween_Click(Index As Integer)

      Dim UpTo As String

19050 calFrom = BetweenDates(Index, UpTo)
19060 calTo = UpTo

End Sub

Private Sub bcancel_Click()

19070 Unload Me

End Sub

Private Sub bPrint_Click()

19080 Printer.Print
19090 Printer.Print
19100 Printer.Print
19110 Printer.Print
19120 Printer.Print "Total Tests - "; Format$(calFrom, "dd/mmm/yyyy"); " to "; Format$(calTo, "dd/mmm/yyyy")
19130 Printer.Print
19140 Printer.Print
19150 g.Col = 0
19160 g.row = 0
19170 g.ColSel = g.Cols - 1
19180 g.RowSel = g.Rows - 1
19190 Printer.Print g.Clip
19200 Printer.EndDoc

End Sub

Private Sub Form_Load()

19210 calFrom = Format$(Now, "dd/mm/yyyy")
19220 calTo = Format$(Now, "dd/mmm/yyyy")

End Sub

Private Sub GenerateLists()

      Dim tb As Recordset
      Dim sql As String

19230 On Error GoTo GenerateLists_Error

19240 sql = "SELECT DISTINCT Ward FROM Demographics WHERE " & _
            "Ward LIKE '%opd%' OR Ward LIKE '%out%' "
19250 Set tb = New Recordset
19260 RecOpenServer 0, tb, sql
19270 strOPD = ""
19280 Do While Not tb.EOF
19290   strOPD = strOPD & " Ward = '" & AddTicks(tb!Ward & "") & "' or"
19300   tb.MoveNext
19310 Loop
19320 If Len(strOPD) > 0 Then
19330   strOPD = Left$(strOPD, Len(strOPD) - 3)
19340 End If

19350 sql = "Select Text from Wards where " & _
            "Location = 'In-House'"
19360 Set tb = New Recordset
19370 RecOpenServer 0, tb, sql
19380 strInHouse = ""
19390 Do While Not tb.EOF
19400   strInHouse = strInHouse & " Ward = '" & AddTicks(tb!Text) & "' or"
      '  strUnknown = strUnknown & " Ward <> '" & AddTicks(tb!Text) & "' and"
19410   tb.MoveNext
19420 Loop
19430 If Len(strInHouse) > 0 Then
19440   strInHouse = Left$(strInHouse, Len(strInHouse) - 3)
19450 End If

19460 sql = "Select Text from Wards where " & _
            "Location = 'External'"
19470 Set tb = New Recordset
19480 RecOpenServer 0, tb, sql
19490 strExternal = ""
19500 Do While Not tb.EOF
19510   strExternal = strExternal & " Ward = '" & AddTicks(tb!Text) & "' or"
      '  strUnknown = strUnknown & " Ward <> '" & AddTicks(tb!Text) & "' and"
19520   tb.MoveNext
19530 Loop
19540 If Len(strExternal) > 0 Then
19550   strExternal = Left$(strExternal, Len(strExternal) - 3)
19560 End If
      'If Len(strUnknown) > 0 Then
      '  strUnknown = Left$(strUnknown, Len(strUnknown) - 3)
      'End If

19570 Exit Sub

GenerateLists_Error:

      Dim strES As String
      Dim intEL As Integer

19580 intEL = Erl
19590 strES = Err.Description
19600 LogError "ftests", "GenerateLists", intEL, strES, sql


End Sub

Private Sub breCalc_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim FromDate As String
      Dim ToDate As String
      Dim s As String
      Dim X As Integer

19610 On Error GoTo breCalc_Click_Error

19620 FromDate = Format$(calFrom, "dd/mmm/yyyy")
19630 ToDate = Format$(calTo, "dd/mmm/yyyy")

19640 g.Rows = 2
19650 g.AddItem ""
19660 g.RemoveItem 1

19670 sql = "Select ShortName, max(printpriority) as m " & _
            "From BioTestDefinitions " & _
            "group by shortname " & _
            "order by m asc"
19680 Set tb = New Recordset
19690 RecOpenServer 0, tb, sql
19700 s = Space$(20)
19710 Do While Not tb.EOF
19720   s = s & "|<" & Left$(tb!ShortName & Space$(5), 5) & " "
19730   tb.MoveNext
19740 Loop

19750 g.FormatString = s

19760 sql = "SELECT DISTINCT Ward FROM demographics WHERE " & _
            "RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
            "ORDER BY Ward"
19770 Set tb = New Recordset
19780 RecOpenServer 0, tb, sql
19790 Do While Not tb.EOF
19800   g.AddItem tb!Ward & ""
19810   tb.MoveNext
19820 Loop

19830 If g.Rows > 2 Then
19840   g.RemoveItem 1
19850 End If

19860 pb.max = g.Rows
19870 pb.Visible = True

19880 For n = 1 To g.Rows - 1
19890   pb = n
19900   For X = 1 To g.Cols - 1
19910     sql = "select count(DISTINCT D.SampleID) as tot " & _
                "from bioresults as r, biotestdefinitions as b, Demographics as D " & _
                "where d.rundate between '" & FromDate & "' and '" & ToDate & "' " & _
                "and b.shortname = '" & g.TextMatrix(0, X) & "' " & _
                "and r.code = b.code " & _
                "AND D.SampleID = R.SampleID " & _
                "AND Ward = '" & AddTicks(g.TextMatrix(n, 0)) & "'"
19920     Set tb = New Recordset
19930     RecOpenClient 0, tb, sql
19940     If tb!Tot <> 0 Then
19950       g.TextMatrix(n, X) = Format$(tb!Tot)
19960       g.Refresh
19970     End If
19980   Next
19990 Next

20000 pb.Visible = False

20010 Exit Sub

breCalc_Click_Error:

      Dim strES As String
      Dim intEL As Integer

20020 intEL = Erl
20030 strES = Err.Description
20040 LogError "ftests", "breCalc_Click", intEL, strES, sql

End Sub

