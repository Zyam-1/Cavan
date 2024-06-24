VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNameExclusions 
   Caption         =   "NetAcquire - Name Exclusions"
   ClientHeight    =   7620
   ClientLeft      =   1785
   ClientTop       =   1200
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   7620
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   5520
      Picture         =   "frmNameExclusions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   6375
      Begin VB.CommandButton bAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   5430
         TabIndex        =   4
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox tReport 
         Height          =   315
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox tGiven 
         Height          =   315
         Left            =   300
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   2325
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6555
      Left            =   300
      TabIndex        =   0
      Top             =   870
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11562
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Given Name                   |<Report As                        "
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
End
Attribute VB_Name = "frmNameExclusions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

57270 On Error GoTo FillG_Error

57280 g.Rows = 2
57290 g.AddItem ""
57300 g.RemoveItem 1

57310 sql = "Select * from NameExclusions " & _
            "Order by GivenName"
57320 Set tb = New Recordset
57330 RecOpenServer 0, tb, sql
57340 Do While Not tb.EOF
57350   s = tb!GivenName & vbTab & tb!ReportName & ""
57360   g.AddItem s
57370   tb.MoveNext
57380 Loop

57390 If g.Rows > 2 Then
57400   g.RemoveItem 1
57410 End If

57420 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

57430 intEL = Erl
57440 strES = Err.Description
57450 LogError "frmNameExclusions", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

      Dim tb As Recordset
      Dim sql As String

57460 On Error GoTo bAdd_Click_Error

57470 If Trim$(tGiven) = "" Then Exit Sub

57480 sql = "Select * from NameExclusions where " & _
            "GivenName = '" & AddTicks(tGiven) & "'"
57490 Set tb = New Recordset
57500 RecOpenServer 0, tb, sql
57510 If tb.EOF Then
57520   tb.AddNew
57530 End If
57540 tb!GivenName = tGiven
57550 tb!ReportName = tReport
57560 tb.Update

57570 FillG
57580 tGiven = ""
57590 tReport = ""

57600 Exit Sub

bAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

57610 intEL = Erl
57620 strES = Err.Description
57630 LogError "frmNameExclusions", "bAdd_Click", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

57640 Unload Me

End Sub


Private Sub Form_Load()

57650 FillG

End Sub
Private Sub g_Click()

      Dim sql As String
      Static SortOrder As Boolean

57660 On Error GoTo g_Click_Error

57670 If g.MouseRow = 0 Then
57680   If SortOrder Then
57690     g.Sort = flexSortGenericAscending
57700   Else
57710     g.Sort = flexSortGenericDescending
57720   End If
57730   SortOrder = Not SortOrder
57740   Exit Sub
57750 End If

57760 If iMsg("Remove " & g.TextMatrix(g.row, 0) & " from list?", vbQuestion + vbYesNo) = vbYes Then
57770   sql = "delete from NameExclusions where " & _
              "GivenName = '" & AddTicks(g.TextMatrix(g.row, 0)) & "'"
57780   Cnxn(0).Execute sql
57790 End If

57800 FillG

57810 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

57820 intEL = Erl
57830 strES = Err.Description
57840 LogError "frmNameExclusions", "g_Click", intEL, strES, sql


End Sub


