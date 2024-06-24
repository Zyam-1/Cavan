VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTotHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Totals for Haematology"
   ClientHeight    =   6225
   ClientLeft      =   615
   ClientTop       =   585
   ClientWidth     =   9390
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   9390
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
      Left            =   7380
      Picture         =   "frmTotHaem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   825
   End
   Begin VB.OptionButton optRet 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
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
      Index           =   0
      Left            =   7680
      TabIndex        =   20
      Top             =   1200
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optRet 
      Caption         =   "A"
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
      Left            =   8040
      TabIndex        =   19
      Top             =   1200
      Width           =   405
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   825
      Left            =   5640
      Picture         =   "frmTotHaem.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   825
   End
   Begin VB.CommandButton cmdPrint 
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
      Height          =   825
      Left            =   6510
      Picture         =   "frmTotHaem.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   825
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Index           =   1
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   120
      Width           =   3615
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
         Left            =   600
         TabIndex        =   16
         Top             =   780
         Width           =   1095
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
         Left            =   540
         TabIndex        =   15
         Top             =   1050
         Width           =   1155
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
         Left            =   300
         TabIndex        =   14
         Top             =   1320
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
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
         Left            =   510
         TabIndex        =   13
         Top             =   1590
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   12
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
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
         Left            =   390
         TabIndex        =   11
         Top             =   2130
         Width           =   1305
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
         Left            =   900
         TabIndex        =   10
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton breCalc 
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
         Left            =   1980
         Picture         =   "frmTotHaem.frx":0FDE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   930
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219217921
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219217921
         CurrentDate     =   38126
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4575
      Left            =   3960
      TabIndex        =   5
      Top             =   1410
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Source                               |^FBC    |^ESR    |^M/S    |^Ret    |^Malaria"
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
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   3960
      TabIndex        =   1
      Top             =   30
      Width           =   1605
      Begin VB.OptionButton oSource 
         Caption         =   "G.P.s"
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
         Left            =   330
         TabIndex        =   4
         Top             =   840
         Width           =   825
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Wards"
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
         Left            =   330
         TabIndex        =   3
         Top             =   570
         Width           =   855
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Clinicians"
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
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.ListBox lstSource 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      IntegralHeight  =   0   'False
      Left            =   210
      TabIndex        =   0
      Top             =   2760
      Width           =   3645
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8220
      TabIndex        =   22
      Top             =   330
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmTotHaem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
      Dim CountWBC As Long
      Dim CountESR As Long
      Dim CountMonoSpot As Long
      Dim CountRet As Long
      Dim CountMalaria As Long
      Dim s As String
      Dim Source As String
      Dim Ret As String

25690 On Error GoTo FillG_Error

25700 If oSource(0) Then
25710   Source = "Clinician"
25720 ElseIf oSource(1) Then
25730   Source = "Ward"
25740 Else
25750   Source = "GP"
25760 End If

25770 Screen.MousePointer = 11

25780 g.Rows = 2
25790 g.AddItem ""
25800 g.RemoveItem 1

25810 If optRet(0) Then Ret = "RetP" Else Ret = "RetA"

25820 For n = 0 To lstSource.ListCount - 1
        
25830   CountWBC = 0
25840   CountESR = 0
25850   CountMonoSpot = 0
25860   CountRet = 0
25870   CountMalaria = 0
        
25880   lstSource.Selected(n) = True
25890   sql = "Select count(WBC) as Tot from HaemResults where " & _
              "WBC is not null and WBC <> '' " & _
              "and SampleID in " & _
              "(Select SampleID from Demographics where " & _
              Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
              "and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "')"
25900   Set tb = New Recordset
25910   RecOpenClient 0, tb, sql
25920   CountWBC = tb!Tot
        
25930   sql = "Select count(ESR) as tot from HaemResults where " & _
              "ESR is not null and ESR <> '' " & _
              "and SampleID in " & _
              "(Select SampleID from Demographics where " & _
              Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
              "and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "')"
25940   Set tb = New Recordset
25950   RecOpenClient 0, tb, sql
25960   CountESR = tb!Tot
        
25970   sql = "Select count(MonoSpot) as tot from HaemResults where " & _
              "MonoSpot is not null and MonoSpot <> '' " & _
              "and SampleID in " & _
              "(Select SampleID from Demographics where " & _
              Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
              "and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "')"
25980   Set tb = New Recordset
25990   RecOpenClient 0, tb, sql
26000   CountMonoSpot = tb!Tot
        
26010   sql = "Select count(" & Ret & ") as tot from HaemResults where " & _
              Ret & " is not null and " & Ret & " <> '' " & _
              "and SampleID in " & _
              "(Select SampleID from Demographics where " & _
              Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
              "and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "')"
26020   Set tb = New Recordset
26030   RecOpenClient 0, tb, sql
26040   CountRet = tb!Tot
        
26050   sql = "Select count(Malaria) as tot from HaemResults where " & _
              "Malaria is not null and Malaria <> '' " & _
              "and SampleID in " & _
              "(Select SampleID from Demographics where " & _
              Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
              "and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
              "and '" & Format$(dtTo, "dd/mmm/yyyy") & "')"
26060   Set tb = New Recordset
26070   RecOpenClient 0, tb, sql
26080   CountMalaria = tb!Tot
26090   s = Left$(lstSource.List(n) & Space(20), 20) & vbTab & _
            Format$(CountWBC) & vbTab & _
            Format$(CountESR) & vbTab & _
            Format$(CountMonoSpot) & vbTab & _
            Format$(CountRet) & vbTab & _
            Format$(CountMalaria)
26100   g.AddItem s
26110 Next

26120 g.AddItem ""
        
26130 CountWBC = 0
26140 CountESR = 0
26150 CountMonoSpot = 0
26160 CountRet = 0
26170 CountMalaria = 0
26180 For n = 1 To g.Rows - 1
26190   CountWBC = CountWBC + Val(g.TextMatrix(n, 1))
26200   CountESR = CountESR + Val(g.TextMatrix(n, 2))
26210   CountMonoSpot = CountMonoSpot + Val(g.TextMatrix(n, 3))
26220   CountRet = CountRet + Val(g.TextMatrix(n, 4))
26230   CountMalaria = CountMalaria + Val(g.TextMatrix(n, 5))
26240 Next
26250 s = "Totals" & vbTab & _
          Format$(CountWBC) & vbTab & _
          Format$(CountESR) & vbTab & _
          Format$(CountMonoSpot) & vbTab & _
          Format$(CountRet) & vbTab & _
          Format$(CountMalaria)
26260 g.AddItem s

26270 If g.Rows > 2 Then
26280   g.RemoveItem 1
26290 End If

26300 Screen.MousePointer = 0

26310 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

26320 intEL = Erl
26330 strES = Err.Description
26340 LogError "ftothaem", "FillG", intEL, strES, sql

End Sub

Private Sub breCalc_Click()

26350 FilllstSource
26360 FillG

End Sub

Private Sub FilllstSource()

      Dim tb As Recordset
      Dim sql As String
      Dim strSource As String
      Dim Found As Boolean
      Dim n As Integer

26370 On Error GoTo FilllstSource_Error

26380 lstSource.Clear

26390 sql = "Select distinct "

26400 If oSource(0) Then
26410   strSource = "Clinician"
26420 ElseIf oSource(1) Then
26430   strSource = "Ward"
26440 Else
26450   strSource = "GP"
26460 End If

26470 sql = sql & strSource & " as Source from Demographics where " & _
            "RunDate between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format(dtTo, "dd/mmm/yyyy") & " 23:59'"
26480 Set tb = New Recordset
26490 RecOpenServer 0, tb, sql
26500 Do While Not tb.EOF
26510   Found = False
26520   For n = 0 To lstSource.ListCount - 1
26530     If Trim$(UCase$(lstSource.List(n))) = Trim$(UCase$(tb!Source & "")) Then
26540       Found = True
26550       Exit For
26560     End If
26570   Next
26580   If Not Found Then
26590     lstSource.AddItem Trim$(tb!Source & "")
26600   End If
26610   tb.MoveNext
26620 Loop

26630 Exit Sub

FilllstSource_Error:

      Dim strES As String
      Dim intEL As Integer

26640 intEL = Erl
26650 strES = Err.Description
26660 LogError "ftothaem", "FilllstSource", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

26670 Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim Source As String
      Dim Y As Integer

26680 If oSource(0) Then
26690   Source = "Clinicians"
26700 ElseIf oSource(1) Then
26710   Source = "Wards"
26720 Else
26730   Source = "GPs"
26740 End If

26750 Printer.Font.Name = "Courier New"

26760 Printer.Print "Totals for Haematology"
26770 Printer.Print "List of " & Source
26780 Printer.Print "Between "; Format(dtFrom, "dd/mmm/yyyy"); " and "; Format(dtTo, "dd/mmm/yyyy")
26790 Printer.Print

26800 For Y = 0 To g.Rows - 1
26810   Printer.Print g.TextMatrix(Y, 0);
26820   Printer.Print Tab(40); g.TextMatrix(Y, 1);
26830   Printer.Print Tab(46); g.TextMatrix(Y, 2);
26840   Printer.Print Tab(52); g.TextMatrix(Y, 3);
26850   Printer.Print Tab(58); g.TextMatrix(Y, 4)
26860 Next

26870 Printer.EndDoc

End Sub

Private Sub cmdXL_Click()

26880 ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CloseUp()

26890 FilllstSource

End Sub


Private Sub dtTo_CloseUp()

26900 FilllstSource

End Sub


Private Sub Form_Load()

26910 dtFrom = Format$(Now, "dd/mmm/yyyy")
26920 dtTo = dtFrom

26930 FilllstSource
26940 FillG

End Sub

Private Sub obetween_Click(Index As Integer)

      Dim UpTo As String

26950 dtFrom = Format$(BetweenDates(Index, UpTo), "dd/mmm/yyyy")
26960 dtTo = Format$(UpTo, "dd/mmm/yyyy")

26970 FilllstSource

End Sub


Private Sub oSource_Click(Index As Integer)

26980 FilllstSource

End Sub

