VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Biochemistry - Totals"
   ClientHeight    =   7845
   ClientLeft      =   90
   ClientTop       =   645
   ClientWidth     =   13695
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
   ScaleHeight     =   7845
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   9930
      TabIndex        =   14
      Top             =   630
      Width           =   3645
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
         Left            =   690
         TabIndex        =   24
         Top             =   960
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
         Left            =   630
         TabIndex        =   23
         Top             =   1230
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
         Left            =   390
         TabIndex        =   22
         Top             =   1500
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
         Left            =   600
         TabIndex        =   21
         Top             =   1770
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
         Left            =   270
         TabIndex        =   20
         Top             =   2040
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
         Left            =   480
         TabIndex        =   19
         Top             =   2310
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
         Left            =   990
         TabIndex        =   18
         Top             =   720
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
         Left            =   2070
         Picture         =   "frmTotals.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1110
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   330
         TabIndex        =   17
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219152385
         CurrentDate     =   38126
      End
   End
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
      Left            =   12000
      Picture         =   "frmTotals.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4470
      Width           =   1485
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Height          =   615
      Left            =   9930
      TabIndex        =   4
      Top             =   3480
      Width           =   3645
      Begin VB.OptionButton o 
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
         Height          =   255
         Index           =   2
         Left            =   2550
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.OptionButton o 
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
         Height          =   255
         Index           =   1
         Left            =   1410
         TabIndex        =   6
         Top             =   240
         Width           =   825
      End
      Begin VB.OptionButton o 
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7275
      Left            =   13860
      TabIndex        =   2
      Top             =   180
      Width           =   3615
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
      Left            =   10020
      Picture         =   "frmTotals.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4470
      Width           =   1485
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
      Left            =   12000
      Picture         =   "frmTotals.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6900
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7095
      Left            =   240
      TabIndex        =   3
      Top             =   630
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12515
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   "<Source               |<Samples |<Tests      |<T/S      |<Samples |<Tests      |<T/S      |<Samples |<Tests      |<T/S      "
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Endocrinology"
      ForeColor       =   &H80000018&
      Height          =   255
      Index           =   2
      Left            =   7020
      TabIndex        =   25
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Report is being generated.   Please wait."
      Height          =   495
      Left            =   10020
      TabIndex        =   13
      Top             =   5700
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Immunology"
      ForeColor       =   &H80000018&
      Height          =   255
      Index           =   1
      Left            =   4470
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biochemistry"
      ForeColor       =   &H80000018&
      Height          =   255
      Index           =   0
      Left            =   1890
      TabIndex        =   11
      Top             =   360
      Width           =   2565
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   315
      Left            =   12000
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mneDisciplines 
      Caption         =   "Disciplines"
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function GetBIEList(ByVal ListToGet As String) As String

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

23570 On Error GoTo GetBIEList_Error

23580 sql = "SELECT Code FROM BioTestDefinitions WHERE BIE = '" & ListToGet & "'"
23590 Set tb = New Recordset
23600 RecOpenServer 0, tb, sql
23610 s = "("
23620 Do While Not tb.EOF
23630   s = s & "Code = '" & tb!Code & "' OR "
23640   tb.MoveNext
23650 Loop
23660 If s = "(" Then
23670   s = ""
23680 Else
23690   s = Left$(s, Len(s) - 3) & ")"
23700 End If

23710 GetBIEList = s

23720 Exit Function

GetBIEList_Error:

      Dim strES As String
      Dim intEL As Integer

23730 intEL = Erl
23740 strES = Err.Description
23750 LogError "ftotals", "GetBIEList", intEL, strES, sql

End Function

Private Sub cmdCancel_Click()

23760 Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim n As Integer
      Dim X As Integer

23770 Printer.Print "Totals:"; dtFrom; " to "; dtTo

23780 Printer.Print
23790 For n = 0 To g.Rows - 1
23800   g.row = n
23810   For X = 0 To 3
23820     g.Col = X
23830     Printer.Print Tab(Choose(X + 1, 1, 40, 50, 60)); g;
23840   Next
23850   Printer.Print
23860 Next

23870 Printer.EndDoc

End Sub

Private Sub breCalc_Click()

23880 lblWait.Visible = True
23890 lblWait.Refresh

23900 FillList
23910 FillGrid

23920 lblWait.Visible = False

End Sub

Private Sub FillList()

      Dim sql As String
      Dim tb As Recordset
      Dim strSource As String
      Dim FromDate As String
      Dim ToDate As String

23930 On Error GoTo FillList_Error

23940 List1.Clear

23950 FromDate = Format(dtFrom, "dd/mmm/yyyy")
23960 ToDate = Format(dtTo, "dd/mmm/yyyy")

23970 If o(0) Then
23980   strSource = "Clinician"
23990 ElseIf o(1) Then
24000   strSource = "Ward"
24010 Else
24020   strSource = "GP"
24030 End If

24040 sql = "Select distinct " & strSource & " as Source " & _
            "from Demographics WHERE " & _
            "RunDate between '" & FromDate & "' AND '" & ToDate & "' " & _
            "AND " & strSource & " IS NOT NULL " & _
            "AND " & strSource & " <> '' " & _
            "AND SampleID IN ( SELECT DISTINCT SampleID FROM BioResults WHERE " & _
            "                  RunDate between '" & FromDate & "' AND '" & ToDate & "') " & _
            "Order by Source"
24050 Set tb = New Recordset
24060 RecOpenServer 0, tb, sql
24070 Do While Not tb.EOF
24080   If strSource = "Ward" Then
24090     If UCase$(tb!Source) <> "GP" Then
24100       List1.AddItem Trim$(tb!Source)
24110     End If
24120   Else
24130     List1.AddItem Trim$(tb!Source)
24140   End If
24150   tb.MoveNext
24160 Loop

24170 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

24180 intEL = Erl
24190 strES = Err.Description
24200 LogError "ftotals", "FillList", intEL, strES, sql

End Sub

Private Sub FillGrid()

      Dim tb As Recordset
      Dim sql As String
      Dim lngBioSamples As Long
      Dim lngBioTests As Long
      Dim lngImmSamples As Long
      Dim lngImmTests As Long
      Dim lngEndSamples As Long
      Dim lngEndTests As Long
      Dim strBioSamples As String
      Dim strBioTests As String
      Dim strImmSamples As String
      Dim strImmTests As String
      Dim strEndSamples As String
      Dim strEndTests As String
      Dim n As Integer
      Dim BioTPS As String
      Dim ImmTPS As String
      Dim EndTPS As String

      Dim BioList As String
      Dim ImmList As String
      Dim EndList As String

24210 On Error GoTo FillGrid_Error

24220 Screen.MousePointer = 11

24230 g.Rows = 2
24240 g.AddItem ""
24250 g.RemoveItem 1

24260 BioList = GetBIEList("B")
24270 ImmList = GetBIEList("I")
24280 EndList = GetBIEList("E")

24290 For n = 0 To List1.ListCount - 1
24300   List1.Selected(n) = True
24310   If BioList <> "" Then

24320     sql = "Select count(SampleID) as Tests, count(distinct SampleID) as Samples from BioResults where " & _
                BioList & _
                "AND SampleID in (" & _
                "  Select SampleID from demographics where " & _
                "  RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                   Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
24330     If o(0) Then
24340       sql = sql & "clinician = '"
24350     ElseIf o(1) Then
24360       sql = sql & "ward = '"
24370     Else
24380       sql = sql & "gp = '"
24390     End If
24400     sql = sql & AddTicks(List1.List(n)) & " ') )"
24410     Set tb = New Recordset
24420     RecOpenClient 0, tb, sql
24430     lngBioTests = tb!Tests
24440     lngBioSamples = tb!Samples

      '120       sql = "Select count(SampleID) as Tests from BioResults where " & _
      '                BioList & _
      '                "AND SampleID in (" & _
      '                "  Select SampleID from demographics where " & _
      '                "   " & _
      '                "  RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
      '                   Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
      '130       If o(0) Then
      '140         sql = sql & "clinician = '"
      '150       ElseIf o(1) Then
      '160         sql = sql & "ward = '"
      '170       Else
      '180         sql = sql & "gp = '"
      '190       End If
      '200       sql = sql & AddTicks(List1.List(n)) & " ') )"
      '210       Set tb = New Recordset
      '220       RecOpenClient 0, tb, sql
      '230       lngBioTests = tb!Tests
      '
      '240       sql = "Select count(distinct SampleID) as Samples from BioResults where " & _
      '                BioList & _
      '                "AND SampleID in (" & _
      '                "Select SampleID from demographics where " & _
      '                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
      '                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
      '250       If o(0) Then
      '260         sql = sql & "clinician = '"
      '270       ElseIf o(1) Then
      '280         sql = sql & "ward = '"
      '290       Else
      '300         sql = sql & "gp = '"
      '310       End If
      '320       sql = sql & AddTicks(List1.List(n)) & " ') )"
      '330       Set tb = New Recordset
      '340       RecOpenClient 0, tb, sql
      '350       lngBioSamples = tb!Samples
24450   End If
        
24460   If ImmList <> "" Then
24470     sql = "Select count(SampleID) as Tests, count(distinct SampleID) as Samples from BioResults where " & _
                ImmList & _
                "AND SampleID in (" & _
                "Select SampleID from demographics where " & _
                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
24480     If o(0) Then
24490       sql = sql & "clinician = '"
24500     ElseIf o(1) Then
24510       sql = sql & "ward = '"
24520     Else
24530       sql = sql & "gp = '"
24540     End If
24550     sql = sql & AddTicks(List1.List(n)) & " ') )"
24560     Set tb = New Recordset
24570     RecOpenClient 0, tb, sql
24580     lngImmTests = tb!Tests
          
      '500       sql = "Select count(distinct SampleID) as Samples from BioResults where " & _
      '                ImmList & _
      '                "AND SampleID in (" & _
      '                "Select SampleID from demographics where " & _
      '                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
      '                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
      '510       If o(0) Then
      '520         sql = sql & "clinician = '"
      '530       ElseIf o(1) Then
      '540         sql = sql & "ward = '"
      '550       Else
      '560         sql = sql & "gp = '"
      '570       End If
      '580       sql = sql & AddTicks(List1.List(n)) & " ') )"
      '590       Set tb = New Recordset
      '600       RecOpenClient 0, tb, sql
24590     lngImmSamples = tb!Samples
24600   End If
        
24610   If EndList <> "" Then
24620     sql = "Select count(SampleID) as Tests, count(distinct SampleID) as Samples  from BioResults where " & _
                EndList & _
                "AND SampleID in (" & _
                "Select SampleID from demographics where " & _
                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
24630     If o(0) Then
24640       sql = sql & "clinician = '"
24650     ElseIf o(1) Then
24660       sql = sql & "ward = '"
24670     Else
24680       sql = sql & "gp = '"
24690     End If
24700     sql = sql & AddTicks(List1.List(n)) & " ') )"
24710     Set tb = New Recordset
24720     RecOpenClient 0, tb, sql
24730     lngEndTests = tb!Tests
          
      '760       sql = "Select count(distinct SampleID) as Samples from BioResults where " & _
      '                EndList & _
      '                "AND SampleID in (" & _
      '                "Select SampleID from demographics where " & _
      '                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
      '                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' and ("
      '770       If o(0) Then
      '780         sql = sql & "clinician = '"
      '790       ElseIf o(1) Then
      '800         sql = sql & "ward = '"
      '810       Else
      '820         sql = sql & "gp = '"
      '830       End If
      '840       sql = sql & AddTicks(List1.List(n)) & " ') )"
      '850       Set tb = New Recordset
      '860       RecOpenClient 0, tb, sql
24740     lngEndSamples = tb!Samples
24750   End If

24760   If lngBioTests <> 0 And lngBioSamples <> 0 Then
24770     BioTPS = Format$(lngBioTests / lngBioSamples, "##.00")
24780   Else
24790     BioTPS = ""
24800   End If
        
24810   If lngImmTests <> 0 And lngImmSamples <> 0 Then
24820     ImmTPS = Format$(lngImmTests / lngImmSamples, "##.00")
24830   Else
24840     ImmTPS = ""
24850   End If
        
24860   If lngEndTests <> 0 And lngEndSamples <> 0 Then
24870     EndTPS = Format$(lngEndTests / lngEndSamples, "##.00")
24880   Else
24890     EndTPS = ""
24900   End If
        
24910   strBioTests = ""
24920   If lngBioTests <> 0 Then
24930     strBioTests = Format$(lngBioTests)
24940   End If
24950   strBioSamples = ""
24960   If lngBioSamples <> 0 Then
24970     strBioSamples = Format$(lngBioSamples)
24980   End If
        
24990   strImmTests = ""
25000   If lngImmTests <> 0 Then
25010     strImmTests = Format$(lngImmTests)
25020   End If
25030   strImmSamples = ""
25040   If lngImmSamples <> 0 Then
25050     strImmSamples = Format$(lngImmSamples)
25060   End If
        
25070   strEndTests = ""
25080   If lngEndTests <> 0 Then
25090     strEndTests = Format$(lngEndTests)
25100   End If
25110   strEndSamples = ""
25120   If lngEndSamples <> 0 Then
25130     strEndSamples = Format$(lngEndSamples)
25140   End If
        
25150   If Trim$(BioTPS & ImmTPS & EndTPS) <> "" Then
25160     g.AddItem List1.List(n) & vbTab & _
                    strBioSamples & vbTab & _
                    strBioTests & vbTab & _
                    BioTPS & vbTab & _
                    strImmSamples & vbTab & _
                    strImmTests & vbTab & _
                    ImmTPS & vbTab & _
                    strEndSamples & vbTab & _
                    strEndTests & vbTab & _
                    EndTPS

25170     If g.Rows > 28 Then
25180       g.TopRow = (g.Rows - 28)
25190     End If
          
25200     g.Refresh
25210   End If
25220 Next
25230 g.AddItem ""

25240 If g.Rows = 2 Then Exit Sub

25250 lngBioSamples = 0
25260 lngBioTests = 0
25270 lngImmSamples = 0
25280 lngImmTests = 0
25290 lngEndSamples = 0
25300 lngEndTests = 0
25310 For n = 1 To g.Rows - 1
25320   lngBioSamples = lngBioSamples + Val(g.TextMatrix(n, 1))
25330   lngBioTests = lngBioTests + Val(g.TextMatrix(n, 2))
25340   lngImmSamples = lngImmSamples + Val(g.TextMatrix(n, 4))
25350   lngImmTests = lngImmTests + Val(g.TextMatrix(n, 5))
25360   lngEndSamples = lngEndSamples + Val(g.TextMatrix(n, 7))
25370   lngEndTests = lngEndTests + Val(g.TextMatrix(n, 8))
25380 Next

25390 BioTPS = ""
25400 ImmTPS = ""
25410 EndTPS = ""
25420 If lngBioSamples <> 0 And lngBioTests <> 0 Then
25430   BioTPS = Format$(lngBioTests / lngBioSamples, "##.00")
25440 End If
25450 If lngImmSamples <> 0 And lngImmTests <> 0 Then
25460   ImmTPS = Format$(lngImmTests / lngImmSamples, "##.00")
25470 End If
25480 If lngEndSamples <> 0 And lngEndTests <> 0 Then
25490   EndTPS = Format$(lngEndTests / lngEndSamples, "##.00")
25500 End If

25510 g.AddItem "Total" & vbTab & _
                lngBioSamples & vbTab & _
                lngBioTests & vbTab & _
                BioTPS & vbTab & _
                lngImmSamples & vbTab & _
                lngImmTests & vbTab & _
                ImmTPS & vbTab & _
                lngEndSamples & vbTab & _
                lngEndTests & vbTab & _
                EndTPS & vbTab

25520 g.Refresh

25530 g.RemoveItem 1

25540 Screen.MousePointer = 0

25550 Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

25560 intEL = Erl
25570 strES = Err.Description
25580 LogError "ftotals", "FillGrid", intEL, strES, sql

End Sub
Private Sub cmdXL_Click()

25590 ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CloseUp()

25600 FillList

End Sub


Private Sub dtTo_CloseUp()

25610 FillList

End Sub


Private Sub Form_Load()

25620 dtFrom = Format$(Now, "dd/mmm/yyyy")
25630 dtTo = dtFrom

End Sub

Private Sub mneDisciplines_Click()

25640 frmSelectBioImmEnd.Show 1

End Sub

Private Sub o_Click(Index As Integer)

25650 FillList

End Sub

Private Sub obetween_Click(Index As Integer)

      Dim UpTo As String

25660 dtFrom = BetweenDates(Index, UpTo)
25670 dtTo = UpTo

25680 FillList

End Sub

