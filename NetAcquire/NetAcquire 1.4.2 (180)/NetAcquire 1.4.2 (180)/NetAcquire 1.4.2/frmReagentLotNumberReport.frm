VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReagentLotNumberReport 
   Caption         =   "NetAcquire"
   ClientHeight    =   6495
   ClientLeft      =   390
   ClientTop       =   630
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   885
      Left            =   6360
      Picture         =   "frmReagentLotNumberReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   330
      Width           =   765
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   885
      Left            =   5190
      Picture         =   "frmReagentLotNumberReport.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   330
      Width           =   765
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Calculate"
      Height          =   885
      Left            =   4140
      Picture         =   "frmReagentLotNumberReport.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   330
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameter"
      Height          =   1215
      Left            =   2340
      TabIndex        =   4
      Top             =   210
      Width           =   1605
      Begin VB.OptionButton optAnalyte 
         Caption         =   "Sickledex"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1065
      End
      Begin VB.OptionButton optAnalyte 
         Caption         =   "Malaria"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   555
         Width           =   885
      End
      Begin VB.OptionButton optAnalyte 
         Caption         =   "Monospot"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4725
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
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
      ScrollBars      =   2
      FormatString    =   "<Date/Time of Entry       |<Sample ID      |<Lot Number              |<Expiry                "
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
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1935
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   720
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   326172673
         CurrentDate     =   38408
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   326172673
         CurrentDate     =   38408
      End
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
      Left            =   4920
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmReagentLotNumberReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean
Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Long
      Dim Analyte As String

42490 On Error GoTo FillG_Error

42500 For n = 0 To 2
42510   If optAnalyte(n) Then
42520     Analyte = optAnalyte(n).Caption
42530     Exit For
42540   End If
42550 Next

42560 sql = "Select * from ReagentLotNumbers where " & _
            "EntryDateTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "and Analyte = '" & Analyte & "' " & _
            "order by EntryDateTime desc"
42570 Set tb = New Recordset
42580 RecOpenServer 0, tb, sql

42590 g.Rows = 2
42600 g.AddItem ""
42610 g.RemoveItem 1

42620 Do While Not tb.EOF
42630   s = Format$(tb!EntryDateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!SampleID & vbTab & _
            tb!LotNumber & vbTab & _
            Format$(tb!Expiry, "dd/mm/yy")
42640   g.AddItem s
42650   tb.MoveNext
42660 Loop

42670 If g.Rows > 2 Then
42680   g.RemoveItem 1
42690 End If

42700 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

42710 intEL = Erl
42720 strES = Err.Description
42730 LogError "frmReagentLotNumberReport", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

42740 Unload Me

End Sub


Private Sub cmdGo_Click()

42750 FillG

End Sub

Private Sub cmdXL_Click()

42760 ExportFlexGrid g, Me

End Sub


Private Sub Form_Load()

42770 dtFrom = Format$(Now - 7, "dd/mm/yyyy")
42780 dtTo = Format$(Now, "dd/mm/yyyy")

End Sub


Private Sub g_Click()

42790 If g.MouseRow = 0 Then
42800   If g.Col = 0 Or g.Col = 3 Then
42810     g.Sort = 9
42820   Else
42830     If SortOrder Then
42840       g.Sort = flexSortGenericAscending
42850     Else
42860       g.Sort = flexSortGenericDescending
42870     End If
42880   End If
42890   SortOrder = Not SortOrder
42900 End If

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

42910 If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
42920   Cmp = 0
42930   Exit Sub
42940 End If

42950 If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
42960   Cmp = 0
42970   Exit Sub
42980 End If

42990 d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
43000 d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

43010 If SortOrder Then
43020   Cmp = Sgn(DateDiff("s", d1, d2))
43030 Else
43040   Cmp = Sgn(DateDiff("s", d2, d1))
43050 End If

End Sub


Private Sub optAnalyte_Click(Index As Integer)

43060 FillG

End Sub


