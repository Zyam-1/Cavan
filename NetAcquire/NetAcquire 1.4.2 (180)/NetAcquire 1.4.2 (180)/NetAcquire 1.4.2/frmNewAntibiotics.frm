VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNewAntibiotics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9450
      Top             =   3870
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9450
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   390
      TabIndex        =   5
      Top             =   120
      Width           =   8175
      Begin VB.Frame Frame4 
         Caption         =   "Penicillin Allergy"
         Height          =   555
         Left            =   5820
         TabIndex        =   20
         Top             =   810
         Width           =   1875
         Begin VB.OptionButton optPenAll 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   22
            Top             =   240
            Width           =   885
         End
         Begin VB.OptionButton optPenAll 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   270
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   420
         Width           =   1620
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   795
         Left            =   2490
         Picture         =   "frmNewAntibiotics.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox txtAntibiotic 
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   1020
         Width           =   3105
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pregnancy"
         Height          =   555
         Left            =   5820
         TabIndex        =   12
         Top             =   210
         Width           =   1875
         Begin VB.OptionButton optPreg 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optPreg 
            Caption         =   "Exclude"
            Height          =   195
            Index           =   1
            Left            =   870
            TabIndex        =   13
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Out Patients"
         Height          =   555
         Left            =   3570
         TabIndex        =   9
         Top             =   210
         Width           =   1875
         Begin VB.OptionButton optOP 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton optOP 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   10
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Children 
         Caption         =   "Children"
         Height          =   555
         Left            =   3570
         TabIndex        =   6
         Top             =   810
         Width           =   1875
         Begin VB.OptionButton optChildren 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   8
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optChildren 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Antibiotic"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   210
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8910
      Picture         =   "frmNewAntibiotics.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4380
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8910
      Picture         =   "frmNewAntibiotics.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   8910
      Picture         =   "frmNewAntibiotics.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6870
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8910
      Picture         =   "frmNewAntibiotics.frx":1330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdAB 
      Height          =   5925
      Left            =   390
      TabIndex        =   0
      Top             =   1680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10451
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   "<Code       |<Antibiotic Name                  |^Pregnancy|^Out-Patients|^Children|^Pen.All |^View  "
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
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   9210
      Picture         =   "frmNewAntibiotics.frx":199A
      Top             =   1680
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9210
      Picture         =   "frmNewAntibiotics.frx":1C70
      Top             =   1200
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmNewAntibiotics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer
Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim TickSave(2 To 6) As Boolean
      Dim VisibleRows As Integer

57850 With grdAB
57860   If .row = .Rows - 1 Then Exit Sub
        
57870   n = .row

57880   FireCounter = FireCounter + 1
57890   If FireCounter > 5 Then
57900     tmrDown.Interval = 100
57910   End If
        
57920   VisibleRows = .height \ .RowHeight(1) - 1
        
57930   .Visible = False
        
57940   s = .TextMatrix(n, 0) & vbTab & .TextMatrix(n, 1)
57950   For X = 2 To 6
57960     .Col = X
57970     TickSave(X) = .CellPicture = imgSquareTick.Picture
57980   Next
        
57990   .RemoveItem n
58000   If n < .Rows Then
58010     .AddItem s, n + 1
58020     .row = n + 1
58030   Else
58040     .AddItem s
58050     .row = .Rows - 1
58060   End If
        
58070   For X = 0 To .Cols - 1
58080     .Col = X
58090     .CellBackColor = vbYellow
58100   Next
        
58110   For X = 2 To 6
58120     .Col = X
58130     .CellPictureAlignment = flexAlignCenterCenter
58140     Set .CellPicture = IIf(TickSave(X), imgSquareTick.Picture, imgSquareCross.Picture)
58150   Next
        
58160   If Not .RowIsVisible(.row) Or .row = .Rows - 1 Then
58170     If .row - VisibleRows + 1 > 0 Then
58180       .TopRow = .row - VisibleRows + 1
58190     End If
58200   End If
        
58210   .Visible = True
        
58220 End With

58230 cmdSave.Enabled = True

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim TickSave(2 To 6) As Boolean

58240 With grdAB
58250   If .row = 1 Then Exit Sub
        
58260   FireCounter = FireCounter + 1
58270   If FireCounter > 5 Then
58280     tmrUp.Interval = 100
58290   End If

58300   n = .row
        
58310   .Visible = False

58320   s = .TextMatrix(n, 0) & vbTab & .TextMatrix(n, 1)
          
58330   For X = 2 To 6
58340     .Col = X
58350     TickSave(X) = .CellPicture = imgSquareTick.Picture
58360   Next
        
58370   .RemoveItem n
58380   .AddItem s, n - 1
        
58390   .row = n - 1
58400   For X = 0 To .Cols - 1
58410     .Col = X
58420     .CellBackColor = vbYellow
58430   Next
        
58440   For X = 2 To 6
58450     .Col = X
58460     .CellPictureAlignment = flexAlignCenterCenter
58470     Set .CellPicture = IIf(TickSave(X), imgSquareTick.Picture, imgSquareCross.Picture)
58480   Next
        
58490   If Not .RowIsVisible(.row) Then
58500     .TopRow = .row
58510   End If
        
58520   .Visible = True
        
58530 End With

58540 cmdSave.Enabled = True

End Sub





Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

58550 On Error GoTo FillG_Error

58560 grdAB.Rows = 2
58570 grdAB.AddItem ""
58580 grdAB.RemoveItem 1

58590 sql = "Select * from Antibiotics " & _
            "order by ListOrder asc"
58600 Set tb = New Recordset
58610 RecOpenClient 0, tb, sql
58620 Do While Not tb.EOF
58630   s = tb!Code & vbTab & tb!AntibioticName
58640   grdAB.AddItem s
58650   grdAB.row = grdAB.Rows - 1
          
58660   grdAB.Col = 2
58670   grdAB.CellPictureAlignment = flexAlignCenterCenter
58680   If Not IsNull(tb!AllowIfPregnant) Then
58690     If tb!AllowIfPregnant Then
58700       Set grdAB.CellPicture = imgSquareTick.Picture
58710     Else
58720       Set grdAB.CellPicture = imgSquareCross.Picture
58730     End If
58740   Else
58750     Set grdAB.CellPicture = imgSquareTick.Picture
58760   End If
        
58770   grdAB.Col = 3
58780   grdAB.CellPictureAlignment = flexAlignCenterCenter
58790   If Not IsNull(tb!AllowIfOutPatient) Then
58800     If tb!AllowIfOutPatient <> 0 Then
58810       Set grdAB.CellPicture = imgSquareTick.Picture
58820     Else
58830       Set grdAB.CellPicture = imgSquareCross.Picture
58840     End If
58850   Else
58860     Set grdAB.CellPicture = imgSquareTick.Picture
58870   End If
        
58880   grdAB.Col = 4
58890   grdAB.CellPictureAlignment = flexAlignCenterCenter
58900   If Not IsNull(tb!AllowIfChild) Then
58910     If tb!AllowIfChild <> 0 Then
58920       Set grdAB.CellPicture = imgSquareTick.Picture
58930     Else
58940       Set grdAB.CellPicture = imgSquareCross.Picture
58950     End If
58960   Else
58970     Set grdAB.CellPicture = imgSquareTick.Picture
58980   End If
        
58990   grdAB.Col = 5
59000   grdAB.CellPictureAlignment = flexAlignCenterCenter
59010   If Not IsNull(tb!AllowIfPenAll) Then
59020     If tb!AllowIfPenAll <> 0 Then
59030       Set grdAB.CellPicture = imgSquareTick.Picture
59040     Else
59050       Set grdAB.CellPicture = imgSquareCross.Picture
59060     End If
59070   Else
59080     Set grdAB.CellPicture = imgSquareTick.Picture
59090   End If

59100   grdAB.Col = 6
59110   grdAB.CellPictureAlignment = flexAlignCenterCenter
59120   If Not IsNull(tb!ViewInGrid) Then
59130     If tb!ViewInGrid <> 0 Then
59140       Set grdAB.CellPicture = imgSquareTick.Picture
59150     Else
59160       Set grdAB.CellPicture = imgSquareCross.Picture
59170     End If
59180   Else
59190     Set grdAB.CellPicture = imgSquareTick.Picture
59200   End If

59210   tb.MoveNext
59220 Loop

59230 If grdAB.Rows > 2 Then
59240   grdAB.RemoveItem 1
59250 End If

59260 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

59270 intEL = Erl
59280 strES = Err.Description
59290 LogError "frmNewAntibiotics", "FillG", intEL, strES, sql


End Sub

Private Sub cmdAdd_Click()

      Dim s As String
      Dim tb As Recordset
      Dim sql As String

59300 On Error GoTo cmdAdd_Click_Error

59310 If Trim$(txtAntibiotic) = "" Then Exit Sub
59320 If Trim$(txtCode) = "" Then Exit Sub

59330 sql = "Select * from Antibiotics where Code = '" & Trim$(txtCode) & "'"
59340 Set tb = New Recordset
59350 RecOpenServer 0, tb, sql
59360 If Not tb.EOF Then
59370   iMsg "Code Used already!", vbExclamation
59380   txtCode = ""
59390   Exit Sub
59400 End If

59410 s = txtCode & vbTab & txtAntibiotic
59420 grdAB.AddItem s
59430 grdAB.row = grdAB.Rows - 1
          
59440 grdAB.Col = 2
59450 grdAB.CellPictureAlignment = flexAlignCenterCenter
59460 If optPreg(0) Then
59470   Set grdAB.CellPicture = imgSquareTick.Picture
59480 Else
59490   Set grdAB.CellPicture = imgSquareCross.Picture
59500 End If

59510 grdAB.Col = 3
59520 grdAB.CellPictureAlignment = flexAlignCenterCenter
59530 If optOP(0) Then
59540   Set grdAB.CellPicture = imgSquareTick.Picture
59550 Else
59560   Set grdAB.CellPicture = imgSquareCross.Picture
59570 End If

59580 grdAB.Col = 4
59590 grdAB.CellPictureAlignment = flexAlignCenterCenter
59600 If optChildren(0) Then
59610   Set grdAB.CellPicture = imgSquareTick.Picture
59620 Else
59630   Set grdAB.CellPicture = imgSquareCross.Picture
59640 End If

59650 grdAB.Col = 5
59660 grdAB.CellPictureAlignment = flexAlignCenterCenter
59670 If optPenAll(0) Then
59680   Set grdAB.CellPicture = imgSquareTick.Picture
59690 Else
59700   Set grdAB.CellPicture = imgSquareCross.Picture
59710 End If

59720 grdAB.Col = 6
59730 grdAB.CellPictureAlignment = flexAlignCenterCenter
59740 Set grdAB.CellPicture = imgSquareTick.Picture

59750 txtAntibiotic = ""
59760 txtCode = ""
59770 cmdSave.Enabled = True

59780 Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

59790 intEL = Erl
59800 strES = Err.Description
59810 LogError "frmNewAntibiotics", "cmdadd_Click", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

59820 If cmdSave.Enabled Then
59830   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
59840     Exit Sub
59850   End If
59860 End If

59870 Unload Me
        
End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

59880 FireDown

59890 tmrDown.Interval = 250
59900 FireCounter = 0

59910 tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

59920 tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

59930 FireUp

59940 tmrUp.Interval = 250
59950 FireCounter = 0

59960 tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

59970 tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

59980 On Error GoTo cmdSave_Click_Error

59990 For n = 1 To grdAB.Rows - 1
60000   If grdAB.TextMatrix(n, 1) <> "" Then
60010     sql = "Select * from Antibiotics where " & _
                "AntibioticName = '" & grdAB.TextMatrix(n, 1) & "'"
          
60020     Set tb = New Recordset
60030     RecOpenClient 0, tb, sql
60040     If tb.EOF Then
60050       tb.AddNew
60060     End If
60070     tb!Code = grdAB.TextMatrix(n, 0)
60080     tb!AntibioticName = grdAB.TextMatrix(n, 1)
60090     tb!ListOrder = n
60100     grdAB.row = n
60110     grdAB.Col = 2
60120     tb!AllowIfPregnant = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
60130     grdAB.Col = 3
60140     tb!AllowIfOutPatient = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
60150     grdAB.Col = 4
60160     tb!AllowIfChild = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
60170     grdAB.Col = 5
60180     tb!AllowIfPenAll = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
60190     grdAB.Col = 6
60200     tb!ViewInGrid = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
60210     tb.Update
60220   End If
60230 Next

60240 cmdSave.Enabled = False

60250 Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60260 intEL = Erl
60270 strES = Err.Description
60280 LogError "frmNewAntibiotics", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

60290 FillG

End Sub

Private Sub grdAB_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim xSave As Integer

60300 With grdAB
60310   ySave = .row
60320   xSave = .Col
60330   .Visible = False
60340   .Col = 0
60350   For Y = 1 To .Rows - 1
60360     .row = Y
60370     If .CellBackColor = vbYellow Then
60380       For X = 0 To .Cols - 1
60390         .Col = X
60400         .CellBackColor = 0
60410       Next
60420       Exit For
60430     End If
60440   Next
60450   .row = ySave
60460   .Col = xSave
60470   .Visible = True
60480 End With

60490 If grdAB.MouseRow = 0 Then
60500   If SortOrder Then
60510     grdAB.Sort = flexSortGenericAscending
60520   Else
60530     grdAB.Sort = flexSortGenericDescending
60540   End If
60550   SortOrder = Not SortOrder
60560   Exit Sub
60570 End If

60580 Select Case grdAB.Col
        Case 0:
60590     grdAB.Enabled = False
60600     grdAB.TextMatrix(grdAB.row, 0) = Trim$(UCase$(iBOX("Code for " & Trim$(grdAB.TextMatrix(grdAB.row, 1)) & " ?", "Antibiotic Code", grdAB.TextMatrix(grdAB.row, 0))))
60610     grdAB.Enabled = True
60620     cmdSave.Enabled = True
60630   Case 1:
60640     For X = 0 To grdAB.Cols - 1
60650       grdAB.Col = X
60660       grdAB.CellBackColor = vbYellow
60670     Next
60680     cmdMoveUp.Enabled = True
60690     cmdMoveDown.Enabled = True
60700   Case 2, 3, 4, 5, 6:
60710     If grdAB.CellPicture = imgSquareTick.Picture Then
60720       Set grdAB.CellPicture = imgSquareCross.Picture
60730     Else
60740       Set grdAB.CellPicture = imgSquareTick.Picture
60750     End If
60760     cmdSave.Enabled = True
60770 End Select

End Sub

Private Sub tmrDown_Timer()

60780 FireDown

End Sub


Private Sub tmrUp_Timer()

60790 FireUp

End Sub


