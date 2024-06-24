VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAntiD 
   Caption         =   "Anti-D Prophylaxis"
   ClientHeight    =   5490
   ClientLeft      =   300
   ClientTop       =   600
   ClientWidth     =   8055
   ControlBox      =   0   'False
   DrawWidth       =   10
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
   Icon            =   "frmAntiD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   8055
   Begin VB.Frame Frame2 
      Height          =   2385
      Left            =   4020
      TabIndex        =   20
      Top             =   150
      Width           =   3855
      Begin VB.TextBox tabscreen 
         Height          =   285
         Left            =   1410
         TabIndex        =   27
         Top             =   1860
         Width           =   2175
      End
      Begin VB.ComboBox mward 
         Height          =   315
         Left            =   870
         TabIndex        =   26
         Top             =   960
         Width           =   2715
      End
      Begin VB.OptionButton rhmother 
         Caption         =   "Neg"
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
         Left            =   2970
         TabIndex        =   25
         Top             =   1500
         Width           =   675
      End
      Begin VB.OptionButton rhmother 
         Alignment       =   1  'Right Justify
         Caption         =   "Pos"
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
         Left            =   2310
         TabIndex        =   24
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox tmgroup 
         Height          =   285
         Left            =   1410
         TabIndex        =   23
         Top             =   1500
         Width           =   675
      End
      Begin VB.TextBox tmnum 
         Height          =   285
         Left            =   870
         TabIndex        =   22
         Top             =   360
         Width           =   2715
      End
      Begin VB.TextBox tmname 
         Height          =   285
         Left            =   870
         TabIndex        =   21
         Top             =   660
         Width           =   2715
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "A/B Screen"
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
         Left            =   480
         TabIndex        =   32
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
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
         Left            =   360
         TabIndex        =   31
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label label8 
         AutoSize        =   -1  'True
         Caption         =   "Group"
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
         Left            =   900
         TabIndex        =   30
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hosp #"
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
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Left            =   330
         TabIndex        =   28
         Top             =   720
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   3855
      Begin VB.ListBox code 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3030
         TabIndex        =   14
         Top             =   1500
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox tcoombs 
         Height          =   285
         Left            =   810
         TabIndex        =   13
         Top             =   1800
         Width           =   1875
      End
      Begin VB.ComboBox bward 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   810
         TabIndex        =   12
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton rhbaby 
         Caption         =   "Neg"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   1500
         Width           =   645
      End
      Begin VB.OptionButton rhbaby 
         Alignment       =   1  'Right Justify
         Caption         =   "Pos"
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
         Left            =   1410
         TabIndex        =   10
         Top             =   1500
         Width           =   585
      End
      Begin VB.TextBox tbgroup 
         Height          =   285
         Left            =   810
         TabIndex        =   9
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox tbdob 
         Height          =   285
         Left            =   810
         TabIndex        =   8
         Top             =   660
         Width           =   2655
      End
      Begin VB.TextBox tbnum 
         Height          =   285
         Left            =   810
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Coombs"
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
         Left            =   90
         TabIndex        =   19
         Top             =   1860
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.O.B"
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
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hosp #"
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
         Left            =   135
         TabIndex        =   17
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
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
         Left            =   270
         TabIndex        =   16
         Top             =   1020
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Group"
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
         Left            =   225
         TabIndex        =   15
         Top             =   1560
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gad 
      Height          =   2055
      Left            =   270
      TabIndex        =   5
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   8
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLines       =   2
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
      Height          =   495
      Left            =   6540
      TabIndex        =   3
      Top             =   3300
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
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
      Height          =   495
      Left            =   6540
      TabIndex        =   2
      Top             =   4500
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   1
      Top             =   3900
      Width           =   1335
   End
   Begin VB.CommandButton btnsuggestx 
      Appearance      =   0  'Flat
      Caption         =   "Suggest Nos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   0
      Top             =   2700
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   33
      Top             =   5220
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lresult 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   270
      TabIndex        =   4
      Top             =   2700
      Width           =   5535
   End
End
Attribute VB_Name = "frmAntiD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

      Dim n As Integer
      Dim rq As Integer
      Dim s As String
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "Anti-D Immunoglobulin Prophylaxis"
40    Printer.Print
50    Printer.Print "Mother"; Tab(40); "Baby"
60    Printer.Print "Hosp# "; tmnum; Tab(40); tbnum
70    Printer.Print "Name  "; tmname; Tab(40); "DoB "; tbdob
80    Printer.Print "Ward  "; mward; Tab(40); bward
90    Printer.Print "Group "; tmgroup; " ";
100   Printer.Print Tab(40); tbgroup;
110   Printer.Print Tab(40); "Coombs "; tcoombs
120   Printer.Print
130   Printer.Print lresult
140   Printer.Print

150   For n = 1 To gad.Rows - 1
160     gad.Col = 4
170     gad.Row = n
180     If Trim$(gad.Text) <> "" Then
190       rq = Val(gad.Text)
200       s = Choose(rq, "One Dose", "Two Doses", "Three Doses", "Four Doses")
210       Printer.Print s; " of Serial # ";
220       gad.Col = 0
230       Printer.Print gad.Text;
240       gad.Col = 1
250       Printer.Print ". (Expiry "; gad.Text; ")"
260       Exit For
270     End If
280   Next

290   Printer.EndDoc

300   For Each Px In Printers
310     If Px.DeviceName = OriginalPrinter Then
320       Set Printer = Px
330       Exit For
340     End If
350   Next

End Sub

Private Sub cmdSave_Click()


      Dim tb As Recordset
      Dim groupdone As Integer
      Dim serial As String
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    If Trim$(tmnum) = "" And Trim$(tmname) = "" Then
30      iMsg "Require either Mothers Name" & vbCrLf & "or number.", vbInformation
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60    End If

70    If Trim$(tbnum) = "" And Trim$(tbdob) = "" Then
80      iMsg "Require either Babys D.o.B." & vbCrLf & "or number.", vbInformation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   groupdone = True
130   If Trim$(tmgroup) = "" Or Trim$(tbgroup) = "" Then groupdone = False
140   If rhmother(0) = False And rhmother(1) = False Then groupdone = False
150   If rhbaby(0) = False And rhbaby(1) = False Then groupdone = False
160   If Trim$(tcoombs) = "" Then groupdone = False
170   If Not groupdone Then
180     iMsg "Enter Group/Rhesus/Coombs.", vbInformation
190     If TimedOut Then Unload Me: Exit Sub
200     Exit Sub
210   End If

220   sql = "Select top 1 * from Anti_D"
230   Set tb = New Recordset
240   RecOpenServerBB 0, tb, sql
250   tb.AddNew

260   tb!mname = tmname
270   tb!mnum = tmnum
280   tb!mward = mward.Text
290   tb!mGroup = tmgroup
300   tb!bnum = tbnum
310   If IsDate(tbdob) Then
320     tb!bDoB = Format(tbdob, "dd/mmm/yyyy")
330   Else
340     tb!bDoB = Null
350   End If
360   tb!bward = bward.Text
370   tb!bgroup = tbgroup
380   tb!coombs = (tcoombs = "Positive")
390   tb!mRh = rhmother(0)
400   tb!brh = rhbaby(0)
410   tb!Op = UserCode
420   tb!DateTime = Now

430   tb!given = serial
440   tb.Update

450   ClearDetails

460   tmnum.SetFocus

470   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmAntiD", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub bward_LostFocus()

      Dim intN As Integer
      Dim strFind As String

10    strFind = UCase$(Trim$(bward))

20    For intN = 0 To code.ListCount - 1
30      If strFind = code.List(intN) Then
40        bward = bward.List(intN)
50        Exit For
60      End If
70    Next

End Sub

Private Sub checkallgroups()

      Dim done As Integer

10    done = True
20    If Not rhbaby(0) And Not rhbaby(1) Then done = False
30    If Not rhmother(0) And Not rhmother(1) Then done = False

40    If Not done Then lresult.Caption = "": Exit Sub

50    If rhmother(1) And rhbaby(0) Then
60      lresult.Caption = "Mother suitable for Anti-D Immunoglobulin"
70    Else
80      lresult.Caption = "Anti-D is not indicated."
90    End If

End Sub

Private Sub ClearDetails()

      Dim n As Integer

10    gad.Col = 0
20    gad.Row = 1
30    gad.ColSel = gad.Cols - 1
40    gad.RowSel = gad.Rows - 1
50    gad.Clip = ""
60    bward = ""
70    mward = ""
80    lresult = ""
90    For n = 0 To 1: rhbaby(n) = 0: rhmother(n) = 0: Next
100   tbdob = ""
110   tbgroup = ""
120   tbnum = ""
130   tcoombs = ""
140   tmgroup = ""
150   tmname = ""
160   tmnum = ""

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo Form_Load_Error

20    bward.Clear
30    mward.Clear

40    sql = "Select * from Wards order by ListOrder"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    Do While Not tb.EOF
80      bward.AddItem tb!Text & ""
90      mward.AddItem tb!Text & ""
100     tb.MoveNext
110   Loop

120   gad.Row = 0
130   For n = 0 To 3
140     gad.ColWidth(n) = Choose(n + 1, 1400, 1100, 900, 900, 850)
150     gad.ColAlignment(n) = 2
160     gad.FixedAlignment(n) = 2
170     gad.Col = n
180     gad.Text = Choose(n + 1, "Serial", "Exp", "Supp", "Use")
190   Next

200   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmAntiD", "Form_Load", intEL, strES, sql


End Sub

Private Sub gad_Click()

10    gad.Col = 0
20    If Trim$(gad.Text) = "" Then Beep: Exit Sub
30    If gad.Row = 0 Then Beep: Exit Sub

40    gad.Col = 3
50    If gad.Text = "Allocate" Then
60      gad.Text = ""
70    Else
80      gad.Text = "Allocate"
90    End If

End Sub

Private Function groupchange(ByVal g As String) As String

10    g = Trim$(g)

20    Select Case g
        Case "":   g = "O"
30      Case "O":  g = "A"
40      Case "A":  g = "B"
50      Case "B":  g = "AB"
60      Case "AB": g = ""
70    End Select

80    groupchange = g

End Function

Private Sub mward_LostFocus()

      Dim intN As Integer
      Dim strFind As String

10    strFind = UCase$(Trim$(mward))

20    For intN = 0 To code.ListCount - 1
30      If strFind = code.List(intN) Then
40        mward = mward.List(intN)
50        Exit For
60      End If
70    Next


End Sub

Private Sub rhbaby_Click(Index As Integer)

10    checkallgroups

End Sub

Private Sub rhmother_Click(Index As Integer)

10    checkallgroups

End Sub

Private Sub tbgroup_Change()

10    checkallgroups

End Sub

Private Sub tbgroup_Click()

10    tbgroup = groupchange(tbgroup.Text)

End Sub

Private Sub tbnum_LostFocus()


      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo tbnum_LostFocus_Error

20    sql = "select * from patientdetails where " & _
            "patnum = '" & tbnum & "'"

30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then
60      bward = ""
70      tbdob = ""
80      rhbaby(0) = False
90      rhbaby(1) = False
100     tbgroup = ""
110     tcoombs = ""
120     Exit Sub
130   End If

140   tb.MoveLast
150   bward = tb("ward") & ""
160   If Not IsNull(tb!DoB) Then
170     tbdob = Format(tb!DoB, "dd/mm/yyyy")
180   Else
190     tbdob = ""
200   End If
210   tbgroup = Left$(tb!fGroup & "  ", 2)
220   If tb!previousrh = "+" Then rhbaby(0) = True
230   If tb!previousrh = "-" Then rhbaby(1) = True

240   Exit Sub

tbnum_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmAntiD", "tbnum_LostFocus", intEL, strES, sql


End Sub

Private Sub tcoombs_Click()

10    If Trim$(tcoombs) = "" Then
20      tcoombs = "Negative"
30    ElseIf tcoombs = "Negative" Then
40      tcoombs = "Positive"
50    Else
60      tcoombs = ""
70    End If

End Sub

Private Sub tmgroup_Change()

10    checkallgroups

End Sub

Private Sub tmgroup_Click()

10    tmgroup = groupchange(tmgroup)

End Sub

Private Sub tmnum_LostFocus()


      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo tmnum_LostFocus_Error

20    sql = "select * from patientdetails where " & _
            "patnum = '" & tmnum & "'"

30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    If tb.EOF Then
60      tmname = ""
70      mward = ""
80      rhmother(0) = False
90      rhmother(1) = False
100     tmgroup = ""
110     Exit Sub
120   End If

130   tb.MoveLast

140   tmname = tb("name") & ""
150   mward = tb("ward") & ""
160   tmgroup = Left$(tb("fgroup") & "  ", 2)
170   If tb("previousrh") = "+" Then rhmother(0) = True
180   If tb("previousrh") = "-" Then rhmother(1) = True
190   If Trim$(tb("anti3reported") & "") = "" Then
200     tabscreen = "Negative"
210   Else
220     tabscreen = tb("anti3reported")
230   End If

240   Exit Sub

tmnum_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmAntiD", "tmnum_LostFocus", intEL, strES, sql

End Sub

