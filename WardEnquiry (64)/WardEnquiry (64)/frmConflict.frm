VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmConflict 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demographic Conflict"
   ClientHeight    =   6960
   ClientLeft      =   1245
   ClientTop       =   840
   ClientWidth     =   8760
   HelpContextID   =   10019
   Icon            =   "frmConflict.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewLatest 
      Caption         =   "&View this Record"
      Height          =   405
      Left            =   5520
      TabIndex        =   6
      Top             =   5910
      Width           =   1455
   End
   Begin VB.CommandButton cmdNone 
      Cancel          =   -1  'True
      Caption         =   "&None of the above"
      Height          =   525
      Left            =   7020
      Picture         =   "frmConflict.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   5400
      Picture         =   "frmConflict.frx":13FC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4740
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6800
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
      ScrollBars      =   2
      FormatString    =   "<Chart Number |<Date Of Birth |<Name                                                   |<Hospital              |<Cn  |  "
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
   Begin VB.Image Image3 
      Height          =   720
      Left            =   4860
      Picture         =   "frmConflict.frx":192E
      Top             =   5745
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7860
      Picture         =   "frmConflict.frx":27F8
      Top             =   -60
      Width           =   480
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   8040
      Picture         =   "frmConflict.frx":30C2
      Top             =   6000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   8040
      Picture         =   "frmConflict.frx":3398
      Top             =   6210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblCountWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Found 99999 matches. Refine your search criteria. Only the top 100 matches are shown. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   4290
      Width           =   8430
   End
   Begin VB.Label lblMostRecent 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmConflict.frx":366E
      Height          =   855
      Left            =   300
      TabIndex        =   5
      Top             =   5670
      Width           =   4545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   300
      X2              =   8370
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   4860
      Width           =   5145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight Patient of interest"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   5310
      TabIndex        =   3
      Top             =   90
      Width           =   2625
   End
End
Attribute VB_Name = "frmConflict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pChart As String
Private pPatName As String
Private pDoB As String

Private pRecentChart As String
Private pRecentPatName As String
Private pRecentDoB As String
Private pRecentDate As String

Private SortOrder As Boolean

Private Sub FillInfo()

      Dim tb As Recordset
      Dim sql As String
      Dim Cn As Integer
      Dim S As String
      Dim TotCount As Long
      Dim Latest As String
      Dim Earliest As String

10    On Error GoTo FillInfo_Error

20    Cn = Val(grd.TextMatrix(grd.Row, 4))

30    sql = "Select RunDate from Demographics where " & _
            "PatName = '" & AddTicks(grd.TextMatrix(grd.Row, 2)) & "' "
40    If Trim$(grd.TextMatrix(grd.Row, 0)) <> "" Then
50        sql = sql & "and Chart = '" & AddTicks(grd.TextMatrix(grd.Row, 0)) & "' "
60    Else
70        sql = sql & "and ( Chart = '' ) "
80    End If
90    If IsDate(grd.TextMatrix(grd.Row, 1)) Then
100       sql = sql & "and DoB = '" & Format$(grd.TextMatrix(grd.Row, 1), "dd/mmm/yyyy") & "' "
110   Else
120       sql = sql & "and ( DoB = '' ) "
130   End If
140   sql = sql & "order by RunDate Asc"

150   TotCount = 0
160   Latest = ""
170   Earliest = ""

180   If Cn = -1 Then
190       For Cn = 0 To intOtherHospitalsInGroup
200           Set tb = New Recordset
210           RecOpenClient Cn, tb, sql
220           If Not tb.EOF Then
230               TotCount = TotCount + tb.RecordCount
240               If Not IsDate(Earliest) Then
250                   Earliest = Format$(tb!RunDate, "dd/mmm/yyyy")
260               Else
270                   If DateDiff("d", Earliest, Format$(tb!RunDate, "dd/mmm/yyyy")) < 0 Then
280                       Earliest = Format$(tb!RunDate, "dd/mmm/yyyy")
290                   End If
300               End If
310               tb.MoveLast
320               If Not IsDate(Latest) Then
330                   Latest = Format$(tb!RunDate, "dd/mmm/yyyy")
340               Else
350                   If DateDiff("d", Latest, Format$(tb!RunDate, "dd/mmm/yyyy")) > 0 Then
360                       Latest = Format$(tb!RunDate, "dd/mmm/yyyy")
370                   End If
380               End If
390           End If
400       Next
410   Else
420       Set tb = New Recordset
          ' Set tb = Cnxn(Cn).Execute(sql)
430       RecOpenClient Cn, tb, sql
440       If Not tb.EOF Then
450           TotCount = tb.RecordCount
460           Earliest = tb!RunDate
470           tb.MoveLast
480           Latest = tb!RunDate
490       End If
500   End If

510   S = Format$(TotCount) & " Record"
520   If TotCount > 1 Then
530       S = S & "s : Earliest " & Format$(Earliest, "dd/mmm/yyyy")
540       S = S & " : Latest " & Format$(Latest, "dd/mmm/yyyy")
550   Else
560       S = S & " : RunDate " & Format$(Earliest, "dd/mmm/yyyy")
570   End If
580   lblInfo = S

590   Exit Sub

FillInfo_Error:

      Dim strES As String
      Dim intEL As Integer

600   intEL = Erl
610   strES = Err.Description
620   LogError "frmConflict", "FillInfo", intEL, strES, sql

End Sub

Private Sub cmdOK_Click()

      Dim i As Integer
      Dim Selected As Boolean
10    For i = 1 To grd.Rows - 1
20        grd.Row = i
30        grd.col = 5
40        If grd.CellPicture = imgGreenTick.Picture Then
50            Selected = True
60        End If
70    Next i
80    If Not Selected Then
90        iMsg "Please select at least one demographic entry"
100       Exit Sub
110   End If

120   If grd.MouseRow = 0 Then
130       pChart = ""
140       pPatName = ""
150       pDoB = ""
160   Else
170       pChart = grd.TextMatrix(grd.Row, 0)
180       pDoB = grd.TextMatrix(grd.Row, 1)
190       pPatName = grd.TextMatrix(grd.Row, 2)
200   End If

210   Me.Hide

End Sub

Private Sub cmdNone_Click()

10    pChart = ""
20    pPatName = ""
30    pDoB = ""

40    Me.Hide

End Sub


Private Sub cmdViewLatest_Click()

      Dim i As Integer
10    For i = 1 To grd.Rows - 1
20        grd.Row = i
30        grd.col = 5
40        If grd.TextMatrix(i, 0) = pRecentChart And grd.TextMatrix(i, 1) = pRecentDoB And grd.TextMatrix(i, 2) = pRecentPatName Then
50            Set grd.CellPicture = imgGreenTick.Picture
60        Else
70            Set grd.CellPicture = imgRedCross.Picture
80        End If
90    Next i

100   pChart = pRecentChart
110   pDoB = pRecentDoB
120   pPatName = pRecentPatName

130   Me.Hide

End Sub

Private Sub Form_Activate()

      Dim x As Integer
      Dim S As String
      Dim i As Integer
10    SingleUserUpdateLoggedOn UserName

20    grd.Row = 1
30    For x = 0 To 2
40        grd.col = x
50        grd.CellBackColor = vbYellow
60    Next

70    FillInfo

80    S = "The most recent record for this search is for" & vbCrLf & _
          pRecentPatName & " and is dated " & pRecentDate & "." & vbCrLf
90    If Trim$(pRecentChart) <> "" Then
100       S = S & "Chart Number (MRN) " & pRecentChart & "."
110   Else
120       S = S & "The Chart Number (MRN) was not specified."
130   End If
140   S = S & vbCrLf
150   If IsDate(pRecentDoB) Then
160       S = S & "D.o.B " & pRecentDoB & "."
170   Else
180       S = S & "The Patients D.o.B was not specified."
190   End If
200   lblMostRecent.Caption = S

210   For i = 1 To grd.Rows - 1
220       grd.Row = i
230       grd.col = 5
240       Set grd.CellPicture = imgRedCross.Picture
250   Next i

End Sub

Public Property Get PatName() As String

10    PatName = pPatName

End Property


Public Property Get Chart() As String

10    Chart = pChart

End Property

Public Property Get DoB() As String

10    DoB = pDoB

End Property

Private Sub Form_Load()


10    grd.ColWidth(4) = 0


End Sub

Private Sub grd_Click()

      Dim y As Integer
      Dim x As Integer
      Dim ySave As Integer

10    If grd.MouseRow = 0 Then
20        SortOrder = Not SortOrder
30        If grd.col = 1 Then
40            grd.Sort = 9
50            Exit Sub
60        End If
70        If SortOrder Then
80            grd.Sort = flexSortGenericAscending
90        Else
100           grd.Sort = flexSortGenericDescending
110       End If
120       Exit Sub
130   End If

140   If grd.MouseCol = 5 Then
150       If grd.CellPicture = imgRedCross Then
160           Set grd.CellPicture = imgGreenTick
170       Else
180           Set grd.CellPicture = imgRedCross
190       End If
200   End If

210   ySave = grd.Row

220   grd.col = 0
230   For y = 1 To grd.Rows - 1
240       grd.Row = y
250       If grd.CellBackColor = vbYellow Then
260           For x = 0 To grd.Cols - 1
270               grd.col = x
280               grd.CellBackColor = 0
290           Next
300           Exit For
310       End If
320   Next

330   grd.Row = ySave
340   For x = 0 To grd.Cols - 1
350       grd.col = x
360       grd.CellBackColor = vbYellow
370   Next



380   FillInfo

End Sub

Private Sub grd_Compare(ByVal Row1 As Long, ByVal Row2 As Long, cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(grd.TextMatrix(Row1, 1)) Then
20        cmp = 0
30        Exit Sub
40    End If

50    If Not IsDate(grd.TextMatrix(Row2, 1)) Then
60        cmp = 0
70        Exit Sub
80    End If

90    d1 = Format(grd.TextMatrix(Row1, 1), "dd/mmm/yyyy")
100   d2 = Format(grd.TextMatrix(Row2, 1), "dd/mmm/yyyy")

110   If SortOrder Then
120       cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140       cmp = -Sgn(DateDiff("D", d1, d2))
150   End If

End Sub



Public Property Let RecentPatName(ByVal strNewValue As String)

10    pRecentPatName = strNewValue

End Property
Public Property Let RecentDoB(ByVal strNewValue As String)

10    pRecentDoB = strNewValue

End Property

Public Property Let RecentChart(ByVal strNewValue As String)

10    pRecentChart = strNewValue

End Property


Public Property Let RecentDate(ByVal strNewValue As String)

10    pRecentDate = strNewValue

End Property




Public Property Let CountWarning(ByVal lNewValue As Long)

10    If lNewValue > 100 Then
20        lblCountWarning.Caption = "Found " & Format$(lNewValue) & " matches. " & _
                                    "Refine your search criteria. " & _
                                    "Only the top 100 matches are shown."
30        lblCountWarning.Visible = True
40    Else
50        lblCountWarning.Visible = False
60    End If

End Property


