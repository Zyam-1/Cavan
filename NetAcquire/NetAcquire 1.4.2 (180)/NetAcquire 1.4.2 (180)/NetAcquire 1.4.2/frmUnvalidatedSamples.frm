VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUnvalidatedSamples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Microbiology Unvalidated Samples"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   8550
      Picture         =   "frmUnvalidatedSamples.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   1100
      Left            =   8550
      Picture         =   "frmUnvalidatedSamples.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   1100
      Left            =   8550
      Picture         =   "frmUnvalidatedSamples.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6180
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dpStart 
      Height          =   315
      Left            =   8370
      TabIndex        =   1
      Top             =   480
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   326172673
      CurrentDate     =   40226
   End
   Begin MSComCtl2.DTPicker dpEnd 
      Height          =   315
      Left            =   8370
      TabIndex        =   2
      Top             =   840
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   326172673
      CurrentDate     =   40226
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7005
      Left            =   180
      TabIndex        =   5
      Top             =   270
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Sample ID      |<Patient Name                                          |<Sample Date               |<Demographic Date     "
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
      Left            =   8550
      TabIndex        =   9
      Top             =   4620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   8100
      TabIndex        =   4
      Top             =   900
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   7980
      TabIndex        =   3
      Top             =   540
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date Selection"
      Height          =   195
      Left            =   8370
      TabIndex        =   0
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmUnvalidatedSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortOrder As Boolean

Private Sub cmdExit_Click()
27900 Unload Me
End Sub

Private Sub cmdSearch_Click()

27910 FillList

End Sub

Private Sub cmdXL_Click()

27920 ExportFlexGrid g, Me, "List of Microbiology Unvalidated Samples" & vbCr

End Sub

Private Sub Form_Load()
27930 dpStart.Value = Date - 4
27940 dpEnd.Value = Date
27950 FillList
End Sub


Private Sub FillList()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

27960 On Error GoTo FillList_Error

27970 sql = "SELECT D.SampleID, D.PatName, SampleDate, D.DateTimeDemographics FROM demographics D " & _
              "LEFT OUTER JOIN PrintValidLog P " & _
              "ON D.SampleID = P.SampleID " & _
              "WHERE COALESCE(P.Valid, '') = '' " & _
              "AND D.SampleID Between %microoffset And %microoffset + 10000000 " & _
              "AND D.DateTimeDemographics Between '%startdate' And '%enddate'"
        
27980 sql = Replace(sql, "%microoffset", sysOptMicroOffsetOLD(0))
27990 sql = Replace(sql, "%startdate", Format(dpStart.Value, "yyyy-MM-dd 00:00:00.000"))
28000 sql = Replace(sql, "%enddate", Format(dpEnd.Value, "yyyy-MM-dd 23:59:59.000"))

28010 Set tb = New Recordset
28020 RecOpenClient 0, tb, sql

28030 If Not tb.EOF Then
28040     InitGrid
28050     While Not tb.EOF
      '110           s = Format$(Val(tb!SampleID & "") - sysOptMicroOffset(0)) & vbTab & _
      '                  tb!PatName & vbTab & _
      '                  tb!SampleDate & vbTab & _
      '                  tb!DateTimeDemographics & ""
28060         s = Format$(Val(tb!SampleID & "")) & vbTab & _
                  tb!PatName & vbTab & _
                  tb!SampleDate & vbTab & _
                  tb!DateTimeDemographics & ""
        
28070         g.AddItem s
28080         tb.MoveNext
28090     Wend
28100 End If

28110 Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

28120 intEL = Erl
28130 strES = Err.Description
28140 LogError "frmUnvalidatedSamples", "FillList", intEL, strES, sql

End Sub

Private Sub InitGrid()
28150 g.Rows = 2
28160 g.Col = 1
28170 g.FixedRows = 1
28180 g.FixedCols = 1
28190 g.Rows = 1

End Sub




Private Sub g_Click()

28200 If g.MouseRow = 0 Then
28210   If InStr(g.TextMatrix(0, g.Col), "Date") Then
28220     g.Sort = 9
28230   Else
28240     If SortOrder Then
28250       g.Sort = flexSortGenericAscending
28260     Else
28270       g.Sort = flexSortGenericDescending
28280     End If
28290   End If
28300   SortOrder = Not SortOrder
28310 End If

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

28320 If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
28330   Cmp = 0
28340   Exit Sub
28350 End If

28360 If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
28370   Cmp = 0
28380   Exit Sub
28390 End If

28400 d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
28410 d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

28420 If SortOrder Then
28430   Cmp = Sgn(DateDiff("s", d1, d2))
28440 Else
28450   Cmp = Sgn(DateDiff("s", d2, d1))
28460 End If

End Sub


