VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemographicValidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Demographic Validation"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSingle 
      Caption         =   "Single Entry"
      Height          =   1155
      Left            =   12960
      Picture         =   "frmDemographicValidation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   930
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show"
      Height          =   1575
      Left            =   7560
      TabIndex        =   13
      Top             =   510
      Width           =   1755
      Begin VB.OptionButton optShow 
         Caption         =   "Only not Validated"
         Height          =   375
         Index           =   1
         Left            =   390
         TabIndex        =   15
         Top             =   900
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   14
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "&Undo All Changes"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   9900
      Picture         =   "frmDemographicValidation.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   930
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Navigation"
      Height          =   1575
      Left            =   3570
      TabIndex        =   4
      Top             =   510
      Width           =   3915
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Space Bar"
         Height          =   195
         Left            =   2700
         TabIndex        =   10
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Up Arrow"
         Height          =   195
         Left            =   2700
         TabIndex        =   9
         Top             =   750
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Down Arrow"
         Height          =   195
         Left            =   2700
         TabIndex        =   8
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Toggle Selected Row"
         Height          =   195
         Left            =   735
         TabIndex        =   7
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Move selection Up"
         Height          =   195
         Left            =   945
         TabIndex        =   6
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Move selection Down"
         Height          =   195
         Left            =   735
         TabIndex        =   5
         Top             =   360
         Width           =   1545
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   2370
         Picture         =   "frmDemographicValidation.frx":1794
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2370
         Picture         =   "frmDemographicValidation.frx":2196
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2370
         Picture         =   "frmDemographicValidation.frx":2B98
         Top             =   360
         Width           =   240
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   3570
      TabIndex        =   3
      Top             =   2310
      Visible         =   0   'False
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7485
      Left            =   180
      TabIndex        =   2
      Top             =   2610
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   13203
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   65535
      ForeColorSel    =   255
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"frmDemographicValidation.frx":359A
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1155
      Left            =   14760
      Picture         =   "frmDemographicValidation.frx":36AE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   1155
      Left            =   10950
      Picture         =   "frmDemographicValidation.frx":4578
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2370
      Left            =   150
      TabIndex        =   11
      Top             =   60
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   219742210
      CurrentDate     =   40933
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmDemographicValidation.frx":5442
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmDemographicValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

          Dim s As String
          Dim dx As Demographic
          Dim DV As DemogValidation
          Dim X As Integer
          Dim Y As Integer

30910     On Error GoTo FillG_Error

30920     g.Rows = 2
30930     g.AddItem ""
30940     g.RemoveItem 1
30950     g.Visible = False

          Dim DVs As New DemogValidations
30960     DVs.LoadByDate mvDate.Value

          Dim Dxs As New Demographics
30970     Dxs.LoadAllByEntryDate mvDate.Value

30980     pBar = 0
30990     pBar.max = Dxs.Count + 1
31000     pBar.Visible = True

31010     For Each dx In Dxs
31020         pBar = pBar + 1
31030         Set DV = DVs(dx.SampleID)
31040         If optShow(0) Or (optShow(1) And DV Is Nothing) Then
31050             s = dx.SampleID & vbTab & _
                      dx.Chart & vbTab & _
                      dx.PatName & vbTab & _
                      dx.DoB & vbTab & _
                      dx.Addr0 & " " & dx.Addr1 & vbTab & _
                      dx.GP & vbTab & _
                      dx.Ward & vbTab & _
                      dx.Clinician & vbTab & _
                      dx.Operator & vbTab & _
                      vbTab & _
                      vbTab & _
                      Format$(dx.DateTimeDemographics, "dd/MM/YY HH:nn:ss")
31060             g.AddItem s
31070             If Not DV Is Nothing Then
31080                 g.row = g.Rows - 1
31090                 g.Col = 9
31100                 Set g.CellPicture = imgGreenTick.Picture
31110                 g.CellPictureAlignment = flexAlignLeftCenter
31120                 g = DV.ValidatedBy
31130                 g.TextMatrix(g.row, 10) = "V"
31140             End If
31150         End If
31160     Next
        
31170     If g.Rows > 2 Then
31180         g.RemoveItem 1
31190         For X = 0 To 3
31200             g.Col = X
31210             For Y = 1 To g.Rows - 1
31220                 g.row = Y
31230                 g.CellFontBold = True
31240             Next
31250         Next
31260     End If

31270     g.Visible = True
31280     pBar.Visible = False
31290     cmdExit.Enabled = True
31300     mvDate.Enabled = True

31310     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

31320     intEL = Erl
31330     strES = Err.Description
31340     LogError "frmDemographicValidation", "FillG", intEL, strES
31350     g.Visible = True
31360     pBar.Visible = False
31370     cmdExit.Enabled = True
31380     mvDate.Enabled = True

End Sub

Private Sub SaveG()

          Dim n As Integer
          Dim DVs As New DemogValidations
          Dim DV As DemogValidation
          Dim SaveRow As Integer

31390     On Error GoTo SaveG_Error

31400     SaveRow = g.row

31410     For n = 1 To g.Rows - 1
31420         If g.TextMatrix(n, 0) <> "" Then
31430             If g.TextMatrix(n, 10) <> "V" Then
31440                 If g.TextMatrix(n, 9) <> "" Then
31450                     Set DV = New DemogValidation
31460                     DV.SampleID = g.TextMatrix(n, 0)
31470                     DV.EnteredBy = g.TextMatrix(n, 8)
31480                     DV.EnteredDateTime = g.TextMatrix(n, 11)
31490                     DV.ValidatedBy = g.TextMatrix(n, 9)
31500                     DVs.Add DV
        
31510                     g.TextMatrix(n, 10) = "V"
31520                     g.row = n
31530                     g.Col = 9
31540                     Set g.CellPicture = imgGreenTick.Picture
31550                 End If
31560             End If
31570         End If
31580     Next

31590     If DVs.Count > 0 Then
31600         DVs.SaveAll
31610     End If

31620     g.row = SaveRow
31630     g.Col = 0
31640     g.ColSel = g.Cols - 1

31650     Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer

31660     intEL = Erl
31670     strES = Err.Description
31680     LogError "frmDemographicValidation", "SaveG", intEL, strES

End Sub

Private Sub cmdExit_Click()

31690     Unload Me

End Sub


Private Sub cmdRevert_Click()

31700     FillG
31710     cmdRevert.Enabled = False
31720     cmdSave.Enabled = False
31730     mvDate.Enabled = True

End Sub


Private Sub cmdSave_Click()

31740     SaveG

31750     cmdRevert.Enabled = False
31760     cmdSave.Enabled = False
31770     cmdExit.Enabled = True
31780     mvDate.Enabled = True

End Sub


Private Sub cmdSingle_Click()

31790     frmDemographicValidationSingle.Show 1

End Sub

Private Sub Form_Load()

31800     On Error GoTo Form_Load_Error

31810     mvDate.Value = Now

31820     g.ColWidth(10) = 0

31830     FillG

31840     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

31850     intEL = Erl
31860     strES = Err.Description
31870     LogError "frmDemographicValidation", "Form_Load", intEL, strES

End Sub



Private Sub g_Click()

31880     On Error GoTo g_Click_Error

31890     If g.MouseRow = 0 Then
31900         g.Col = g.MouseCol
31910         If InStr(UCase$(g.TextMatrix(0, g.Col)), "DATE") <> 0 Then
31920             g.Sort = 9
31930         Else
31940             If SortOrder Then
31950                 g.Sort = flexSortGenericAscending
31960             Else
31970                 g.Sort = flexSortGenericDescending
31980             End If
31990         End If
32000         SortOrder = Not SortOrder
32010         Exit Sub
32020     End If

32030     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

32040     intEL = Erl
32050     strES = Err.Description
32060     LogError "frmDemographicValidation", "g_Click", intEL, strES

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

32070     On Error GoTo g_Compare_Error

32080     If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
32090         Cmp = 0
32100         Exit Sub
32110     End If

32120     If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
32130         Cmp = 0
32140         Exit Sub
32150     End If

32160     d1 = Format(g.TextMatrix(Row1, g.Col), "dd/mmm/yyyy hh:mm:ss")
32170     d2 = Format(g.TextMatrix(Row2, g.Col), "dd/mmm/yyyy hh:mm:ss")

32180     If SortOrder Then
32190         Cmp = Sgn(DateDiff("s", d1, d2))
32200     Else
32210         Cmp = Sgn(DateDiff("s", d2, d1))
32220     End If

32230     Exit Sub

g_Compare_Error:

          Dim strES As String
          Dim intEL As Integer

32240     intEL = Erl
32250     strES = Err.Description
32260     LogError "frmDemographicValidation", "g_Compare", intEL, strES

End Sub


Private Function IsAnyNotSaved() As Boolean

          Dim Y As Integer
          Dim RetVal As Boolean

32270     On Error GoTo IsAnyNotSaved_Error

32280     RetVal = False

32290     For Y = 1 To g.Rows - 1
32300         If g.TextMatrix(Y, 9) <> "" And g.TextMatrix(Y, 10) <> "V" Then
32310             RetVal = True
32320             Exit For
32330         End If
32340     Next

32350     IsAnyNotSaved = RetVal

32360     Exit Function

IsAnyNotSaved_Error:

          Dim strES As String
          Dim intEL As Integer

32370     intEL = Erl
32380     strES = Err.Description
32390     LogError "frmDemographicValidation", "IsAnyNotSaved", intEL, strES

End Function

Private Sub g_KeyUp(KeyCode As Integer, Shift As Integer)

32400     On Error GoTo g_KeyUp_Error

32410     Debug.Print KeyCode

32420     If g.TextMatrix(g.row, 0) <> "" And g.row > 0 Then
32430         If KeyCode = vbKeySpace Then
32440             If g.row > 0 Then
32450                 If g.TextMatrix(g.row, 8) <> UserName Then ' I didn't make the first entry
32460                     If g.TextMatrix(g.row, 10) <> "V" Then 'not valid
32470                         If g.TextMatrix(g.row, 9) <> "" Then 'name present
32480                             g.TextMatrix(g.row, 9) = ""
32490                             If IsAnyNotSaved() Then
32500                                 cmdSave.Enabled = True
32510                                 cmdRevert.Enabled = True
32520                                 cmdExit.Enabled = False
32530                                 mvDate.Enabled = False
32540                             Else
32550                                 cmdSave.Enabled = False
32560                                 cmdRevert.Enabled = False
32570                                 cmdExit.Enabled = True
32580                                 mvDate.Enabled = True
32590                             End If
32600                         Else 'no name present
32610                             g.TextMatrix(g.row, 9) = UserName
32620                             If IsAnyNotSaved() Then
32630                                 cmdSave.Enabled = True
32640                                 cmdRevert.Enabled = True
32650                                 cmdExit.Enabled = False
32660                                 mvDate.Enabled = False
32670                             Else
32680                                 cmdSave.Enabled = False
32690                                 cmdRevert.Enabled = False
32700                                 cmdExit.Enabled = True
32710                                 mvDate.Enabled = True
32720                             End If
32730                         End If
32740                     End If
32750                 End If
32760                 SendKeys "{DOWN}"
32770             End If
32780         End If
32790     End If

32800     Exit Sub

g_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

32810     intEL = Erl
32820     strES = Err.Description
32830     LogError "frmDemographicValidation", "g_KeyUp", intEL, strES

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

32840     FillG

End Sub


Private Sub optShow_Click(Index As Integer)

32850     FillG

End Sub

