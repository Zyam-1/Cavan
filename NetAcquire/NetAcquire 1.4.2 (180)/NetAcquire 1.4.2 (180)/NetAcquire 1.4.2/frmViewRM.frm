VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmViewRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - View Running Means"
   ClientHeight    =   4740
   ClientLeft      =   1980
   ClientTop       =   1890
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Smoothing"
      Height          =   885
      Left            =   4980
      TabIndex        =   10
      Top             =   60
      Width           =   1455
      Begin VB.OptionButton optSmoothing 
         Caption         =   "0.99"
         Height          =   195
         Index           =   5
         Left            =   780
         TabIndex        =   16
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optSmoothing 
         Caption         =   "0.97"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   15
         Top             =   420
         Width           =   645
      End
      Begin VB.OptionButton optSmoothing 
         Alignment       =   1  'Right Justify
         Caption         =   "0.93"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optSmoothing 
         Caption         =   "0.95"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton optSmoothing 
         Alignment       =   1  'Right Justify
         Caption         =   "0.91"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   645
      End
      Begin VB.OptionButton optSmoothing 
         Alignment       =   1  'Right Justify
         Caption         =   "0.89"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parameter"
      Height          =   885
      Left            =   1710
      TabIndex        =   8
      Top             =   60
      Width           =   2085
      Begin VB.ComboBox lstParameter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Department"
      Height          =   885
      Left            =   210
      TabIndex        =   5
      Top             =   60
      Width           =   1485
      Begin VB.OptionButton optBio 
         Caption         =   "Biochemistry"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optHaem 
         Caption         =   "Haematology"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Points"
      Height          =   885
      Left            =   3810
      TabIndex        =   2
      Top             =   60
      Width           =   1155
      Begin ComCtl2.UpDown UpDown1 
         Height          =   405
         Left            =   690
         TabIndex        =   3
         Top             =   270
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "lblDataPoints"
         BuddyDispid     =   196617
         OrigLeft        =   240
         OrigTop         =   660
         OrigRight       =   960
         OrigBottom      =   915
         Increment       =   10
         Max             =   500
         Min             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label lblDataPoints 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   645
      Left            =   6540
      Picture         =   "frmViewRM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   975
   End
   Begin MSChart20Lib.MSChart g 
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmViewRM.frx":066A
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   7515
   End
End
Attribute VB_Name = "frmViewRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawHaemGraph()

      Dim rs As Recordset
      Dim sql As String
      Dim X As Long
45390 On Error GoTo DrawHaemGraph_Error

45400 ReDim D(1 To lblDataPoints, 1 To 1) As Variant
      Dim maxy As Single
      Dim miny As Single
      Dim rm As Single
      Dim fld As Field
      Dim Run(500) As Single
      Dim varTot(500) As Single
      Dim varMax As Single
      Dim varMin As Single
      Dim varRm(500) As Single
      Dim varNo As Long
      Dim n As Long
      Dim Smoothing As Single

45410 If lstParameter = "" Then Exit Sub

45420 For n = 0 To 5
45430   If optSmoothing(n) Then
45440     Smoothing = optSmoothing(n).Caption
45450     Exit For
45460   End If
45470 Next

45480 If Val(lblDataPoints) < 50 Then
45490   lblDataPoints = "50"
45500 ElseIf Val(lblDataPoints) > 500 Then
45510   lblDataPoints = "500"
45520 End If

45530 g.Visible = False

45540 sql = "SELECT TOP " & lblDataPoints & " " & lstParameter & _
            " FROM HaemResults  where " & lstParameter & " <> '' " & _
            "ORDER BY rundatetime DESC"

45550 Set rs = New Recordset
45560 RecOpenClient 0, rs, sql

45570 maxy = 0
45580 miny = 999
45590 n = 1

45600 Do While Not rs.EOF
45610   Run(n) = Val(rs("" & lstParameter & ""))
45620   n = n + 1
45630   rs.MoveNext
45640 Loop
45650 varMin = 10000
45660 varMax = 0
45670 For n = 1 To (Val(lblDataPoints) - 21)
45680     For X = n To n + 21
45690        If varMin > Run(X) Then varMin = Run(X)
45700        If varMax < Run(X) Then varMax = Run(X)
45710        varTot(n) = varTot(n) + Run(X)
45720     Next
45730     varTot(n) = varTot(n) - (varMin + varMax)
45740     varTot(n) = varTot(n) / 20
45750     If n = 1 Then
45760         varRm(n) = varTot(n)
45770     Else
45780         varRm(n) = (varTot(n) * (1 - Smoothing)) + (varRm(n - 1) * Smoothing)
45790     End If
45800     varMin = 10000
45810     varMax = 0
45820     varNo = n
45830 Next
       
45840 miny = 9999
45850 maxy = 0
45860 X = varNo - 1
45870 For n = 1 To varNo
45880     X = X - 1
45890     If X = 0 Then Exit For
45900     D(X, 1) = varRm(n)
45910     If varRm(n) < miny Then
45920       miny = varRm(n)
45930     End If
45940     If varRm(n) > maxy Then
45950       maxy = varRm(n)
45960     End If
45970 Next
       
45980 With g.Plot.Axis(VtChAxisIdY).ValueScale
45990   .Auto = False
46000   .Maximum = Int(maxy * 1.05 + 0.5)
46010   If .Maximum = 0 Then
46020     .Maximum = 1
46030   End If
46040   If maxy > .Maximum Then
46050     .Maximum = .Maximum + 1
46060   End If
46070   .Minimum = Int(miny * 0.95 - 0.5)
46080   If .Minimum < 0 Then
46090     .Minimum = 0
46100   End If
46110 End With

46120 g.ChartData = D
46130 g.Visible = True

46140 Exit Sub

DrawHaemGraph_Error:

      Dim strES As String
      Dim intEL As Integer

46150 intEL = Erl
46160 strES = Err.Description
46170 LogError "frmViewRM", "DrawHaemGraph", intEL, strES, sql

End Sub



Private Sub DrawBioGraph()

      Dim tb As Recordset
      Dim sql As String
      Dim X As Long
46180 On Error GoTo DrawBioGraph_Error

46190 ReDim D(1 To lblDataPoints, 1 To 1) As Variant
      Dim maxy As Single
      Dim miny As Single
      Dim rm As Single
      Dim fld As Field
      Dim Run(500) As Single
      Dim varTot(500) As Single
      Dim varMax As Single
      Dim varMin As Single
      Dim varRm(500) As Single
      Dim varNo As Long
      Dim n As Long
      Dim Smoothing As Single

46200 If lstParameter = "" Then Exit Sub

46210 For n = 0 To 5
46220   If optSmoothing(n) Then
46230     Smoothing = optSmoothing(n).Caption
46240     Exit For
46250   End If
46260 Next

46270 If Val(lblDataPoints) < 50 Then
46280   lblDataPoints = "50"
46290 ElseIf Val(lblDataPoints) > 500 Then
46300   lblDataPoints = "500"
46310 End If

46320 g.Visible = False

46330 sql = "SELECT TOP " & lblDataPoints & " Result " & _
            "FROM BioResults " & _
            "WHERE Code IN ( SELECT Code FROM BioTestDefinitions " & _
            "                WHERE LongName = '" & lstParameter.Text & "') " & _
            "AND Result <> '' " & _
            "ORDER BY runtime DESC"

46340 Set tb = New Recordset
46350 RecOpenClient 0, tb, sql

46360 maxy = 0
46370 miny = 999
46380 n = 1

46390 Do While Not tb.EOF
46400   Run(n) = Val(tb!Result)
46410   Debug.Print n, tb!Result
46420   n = n + 1
46430   tb.MoveNext
46440 Loop
46450 varMin = 10000
46460 varMax = 0
46470 For n = 1 To (Val(lblDataPoints) - 21)
46480     For X = n To n + 21
46490        If varMin > Run(X) Then varMin = Run(X)
46500        If varMax < Run(X) Then varMax = Run(X)
46510        varTot(n) = varTot(n) + Run(X)
46520     Next
46530     varTot(n) = varTot(n) - (varMin + varMax)
46540     varTot(n) = varTot(n) / 20
46550     If n = 1 Then
46560         varRm(n) = varTot(n)
46570     Else
46580         varRm(n) = (varTot(n) * (1 - Smoothing)) + (varRm(n - 1) * Smoothing)
46590     End If
46600     varMin = 10000
46610     varMax = 0
46620     varNo = n
46630 Next
       
46640 miny = 9999
46650 maxy = 0
46660 X = varNo - 1
46670 For n = 1 To varNo
46680     X = X - 1
46690     If X = 0 Then Exit For
46700     D(X, 1) = varRm(n)
46710     If varRm(n) < miny Then
46720       miny = varRm(n)
46730     End If
46740     If varRm(n) > maxy Then
46750       maxy = varRm(n)
46760     End If
46770 Next
       
46780 With g.Plot.Axis(VtChAxisIdY).ValueScale
46790   .Auto = False
46800   .Maximum = Int(maxy * 1.05 + 0.5)
46810   If .Maximum = 0 Then
46820     .Maximum = 1
46830   End If
46840   If maxy > .Maximum Then
46850     .Maximum = .Maximum + 1
46860   End If
46870   .Minimum = Int(miny * 0.95 - 0.5)
46880   If .Minimum < 0 Then
46890     .Minimum = 0
46900   End If
46910 End With

46920 g.ChartData = D
46930 g.Visible = True

46940 Exit Sub

DrawBioGraph_Error:

      Dim strES As String
      Dim intEL As Integer

46950 intEL = Erl
46960 strES = Err.Description
46970 LogError "frmViewRM", "DrawBioGraph", intEL, strES, sql

End Sub

Private Sub FillParameterList()

46980 If optHaem Then
46990   FillWithHaem
47000 ElseIf optBio Then
47010   FillWithBio
47020 End If

End Sub

Private Sub FillWithHaem()

      Dim sql As String
      Dim tb As Recordset

47030 On Error GoTo FillWithHaem_Error

47040 lstParameter.Clear

47050 sql = "Select distinct AnalyteName from HaemTestDefinitions"
47060 Set tb = New Recordset
47070 RecOpenServer 0, tb, sql
47080 Do While Not tb.EOF
47090   lstParameter.AddItem tb!AnalyteName & ""
47100   tb.MoveNext
47110 Loop

47120 Exit Sub

FillWithHaem_Error:

      Dim strES As String
      Dim intEL As Integer

47130 intEL = Erl
47140 strES = Err.Description
47150 LogError "frmViewRM", "FillWithHaem", intEL, strES, sql


End Sub

Private Sub FillWithBio()

      Dim sql As String
      Dim tb As Recordset

47160 On Error GoTo FillWithBio_Error

47170 lstParameter.Clear

47180 sql = "Select distinct LongName, PrintPriority from BioTestDefinitions " & _
            "order by PrintPriority"
47190 Set tb = New Recordset
47200 RecOpenClient 0, tb, sql

47210 Do While Not tb.EOF
47220   lstParameter.AddItem tb!LongName
47230   tb.MoveNext
47240 Loop

47250 Exit Sub

FillWithBio_Error:

      Dim strES As String
      Dim intEL As Integer

47260 intEL = Erl
47270 strES = Err.Description
47280 LogError "frmViewRM", "FillWithBio", intEL, strES, sql


End Sub


Private Sub cmdExit_Click()

47290 Unload Me

End Sub

Private Sub Form_Load()

47300 FillParameterList

End Sub

Private Sub lstParameter_Click()

47310 If optHaem Then
47320   DrawHaemGraph
47330 Else
47340   DrawBioGraph
47350 End If

End Sub

Private Sub optBio_Click()

47360 FillParameterList

End Sub

Private Sub optHaem_Click()

47370 FillParameterList

End Sub

Private Sub optSmoothing_Click(Index As Integer)

47380 If optHaem Then
47390   DrawHaemGraph
47400 Else
47410   DrawBioGraph
47420 End If

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


47430 If optHaem Then
47440   DrawHaemGraph
47450 Else
47460   DrawBioGraph
47470 End If

End Sub
