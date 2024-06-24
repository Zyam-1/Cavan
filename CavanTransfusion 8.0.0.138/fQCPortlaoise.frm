VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fQCPortlaoise 
   Caption         =   "Daily ABO-Rh QC"
   ClientHeight    =   5670
   ClientLeft      =   1140
   ClientTop       =   855
   ClientWidth     =   7305
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
   Icon            =   "fQCPortlaoise.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5670
   ScaleWidth      =   7305
   Begin MSFlexGridLib.MSFlexGrid grdTarget 
      Height          =   1575
      Left            =   2850
      TabIndex        =   20
      Top             =   3810
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdTitle 
      Height          =   375
      Index           =   1
      Left            =   2850
      TabIndex        =   19
      Top             =   2640
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdTitle 
      Height          =   375
      Index           =   0
      Left            =   2850
      TabIndex        =   18
      Top             =   720
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdreaction 
      Height          =   1575
      Left            =   2850
      TabIndex        =   17
      Top             =   1080
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdlotnos 
      Height          =   1545
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   2725
      _Version        =   393216
      Rows            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   5
      Left            =   4260
      MaxLength       =   10
      TabIndex        =   13
      Top             =   3420
      Width           =   1155
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   4
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3420
      Width           =   1155
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   3
      Left            =   1980
      MaxLength       =   10
      TabIndex        =   11
      Top             =   3420
      Width           =   1155
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   2
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   9
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   1
      Left            =   3180
      MaxLength       =   10
      TabIndex        =   8
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox lot 
      Height          =   285
      Index           =   0
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   7
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton btnprint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   6
      Top             =   2100
      Width           =   1995
   End
   Begin VB.CommandButton btnloadprev 
      Appearance      =   0  'Flat
      Caption         =   "Load Previous"
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
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   2115
   End
   Begin VB.CommandButton btnview 
      Appearance      =   0  'Flat
      Caption         =   "View Targets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   2115
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   1
      Top             =   1140
      Width           =   1995
   End
   Begin VB.CommandButton btncancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   3
      Top             =   4800
      Width           =   1995
   End
   Begin VB.TextBox txtinput 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   660
      Width           =   840
   End
   Begin VB.TextBox txttitinput 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7440
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox tinput 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1860
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1110
      TabIndex        =   21
      Top             =   5460
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line8 
      X1              =   4260
      X2              =   4020
      Y1              =   3420
      Y2              =   3000
   End
   Begin VB.Line Line7 
      X1              =   3120
      X2              =   3480
      Y1              =   3420
      Y2              =   3000
   End
   Begin VB.Line Line6 
      X1              =   1980
      X2              =   2880
      Y1              =   3420
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   5460
      X2              =   4620
      Y1              =   540
      Y2              =   780
   End
   Begin VB.Line Line3 
      X1              =   4320
      X2              =   4080
      Y1              =   540
      Y2              =   780
   End
   Begin VB.Line Line2 
      X1              =   3180
      X2              =   3420
      Y1              =   540
      Y2              =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   2100
      X2              =   2940
      Y1              =   540
      Y2              =   780
   End
   Begin VB.Line Line9 
      X1              =   4620
      X2              =   5400
      Y1              =   3000
      Y2              =   3420
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      Height          =   1275
      Left            =   2640
      Top             =   1860
      Width           =   2235
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   2640
      Top             =   660
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lot Number"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   3450
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lot Number"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   270
      Width           =   825
   End
End
Attribute VB_Name = "fQCPortlaoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prevtarget As String

Private Sub btnCancel_Click()

10    Unload Me

End Sub

Private Sub btnloadprev_Click()

      Dim mt As Recordset
      Dim sql As String

10    On Error GoTo btnloadprev_Click_Error

20    sql = "Select * from ABORhQC where " & _
            "DateTime = '01/01/1950'"
30    Set mt = New Recordset
40    RecOpenServerBB 0, mt, sql
50    If mt.EOF Then Exit Sub
60    mt.MoveLast

70    grdTitle(0).Col = 0
80    grdTitle(1).Col = 0
90    grdTitle(0).Row = 0
100   grdTitle(1).Row = 0
110   grdTitle(0).ColSel = 2
120   grdTitle(1).ColSel = 2
130   grdTitle(0).RowSel = 0
140   grdTitle(1).RowSel = 0

150   grdreaction.Col = 0
160   grdreaction.Row = 0
170   grdreaction.ColSel = 2
180   grdreaction.RowSel = 5

190   grdTarget.Col = 0
200   grdTarget.Row = 0
210   grdTarget.ColSel = 2
220   grdTarget.RowSel = 5

230   grdLotNos.Col = 0
240   grdLotNos.Row = 0
250   grdLotNos.ColSel = 1
260   grdLotNos.RowSel = 5

270   grdTitle(0).Clip = mt!Title
280   grdTitle(1).Clip = mt!title1 & ""
290   grdLotNos.Clip = mt!lotnos
300   grdreaction.Clip = mt("reaction")
310   grdTarget.Clip = mt("target")

320   prevtarget = mt("target")

330   lot(0) = mt!lot0 & ""
340   lot(1) = mt!lot1 & ""
350   lot(2) = mt!lot2 & ""
360   lot(3) = mt!lot3 & ""
370   lot(4) = mt!lot4 & ""
380   lot(5) = mt!lot5 & ""

390   Exit Sub

btnloadprev_Click_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "fQCPortlaoise", "btnloadprev_Click", intEL, strES, sql


End Sub

Private Sub btnprint_Click()

      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    btnprint.Visible = False
40    btnloadprev.Visible = False
50    cmdSave.Visible = False
60    btncancel.Visible = False
70    btnview.Visible = False

80    PrintForm

90    btnprint.Visible = True
100   btnloadprev.Visible = True
110   cmdSave.Visible = True
120   btncancel.Visible = True
130   btnview.Visible = True

140   For Each Px In Printers
150     If Px.DeviceName = OriginalPrinter Then
160       Set Printer = Px
170       Exit For
180     End If
190   Next

End Sub

Private Sub cmdSave_Click()

      Dim n As Integer
      Dim Y As Integer
      Dim full As Integer
      Dim reaction As String
      Dim target As String
      Dim mt As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    full = True
30    For n = 0 To 2
40      grdTitle(0).Col = n
50      If Trim$(grdTitle(0)) = "" Then full = False
        'grdtitle(1).Col = n
        'If trim$(grdtitle(1)) = "" Then full = False
60    Next
70    If Not full Then
80      iMsg "Titles not filled", vbCritical
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   For n = 0 To 2
130     If Trim$(lot(n)) = "" Then full = False
140   Next
150   If Not full Then
160     iMsg "Lot Numbers of Reagents not filled", vbCritical
170     If TimedOut Then Unload Me: Exit Sub
180     Exit Sub
190   End If

200   For n = 0 To 1
210     grdLotNos.Col = n
220     For Y = 0 To 5
230       grdLotNos.Row = Y
240       If Trim$(grdLotNos) = "" Then full = False
250     Next
260   Next
270   If Not full Then
280     iMsg "Lot Numbers or Type not entered.", vbCritical
290     If TimedOut Then Unload Me: Exit Sub
300     Exit Sub
310   End If

320   grdTitle(0).Col = 0
330   grdTitle(1).Col = 0
340   grdTitle(0).Row = 0
350   grdTitle(1).Row = 0
360   grdTitle(0).ColSel = 2
370   grdTitle(1).ColSel = 2
380   grdTitle(0).RowSel = 0
390   grdTitle(1).RowSel = 0

400   grdreaction.Col = 0
410   grdreaction.Row = 0
420   grdreaction.ColSel = 2
430   grdreaction.RowSel = 5

440   grdTarget.Col = 0
450   grdTarget.Row = 0
460   grdTarget.ColSel = 2
470   grdTarget.RowSel = 5

480   grdLotNos.Col = 0
490   grdLotNos.Row = 0
500   grdLotNos.ColSel = 1
510   grdLotNos.RowSel = 5

520   reaction = grdreaction.Clip
530   target = grdTarget.Clip

540   If reaction <> target Then
550     iMsg "Reactions and Targets don't match.", vbCritical
560     If TimedOut Then Unload Me: Exit Sub
570     Exit Sub
580   End If

590   sql = "Select * from ABORhQC where Title = 'x'"
600   Set mt = New Recordset
610   RecOpenServerBB 0, mt, sql
620   mt.AddNew
630   mt!Title = grdTitle(0).Clip
640   mt!title1 = grdTitle(1).Clip
650   mt!lotnos = grdLotNos.Clip
660   mt!reaction = grdreaction.Clip
670   mt!target = grdTarget.Clip
680   mt!Operator = UserCode
690   mt!DateTime = Now
700   For n = 0 To 5
710     mt("lot" & Format(n)) = lot(n)
720   Next

730   mt.Update

740   Unload Me

750   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

760   intEL = Erl
770   strES = Err.Description
780   LogError "fQCPortlaoise", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub btnview_Click()

      Dim n As Integer
      Dim Y As Integer
      Dim full As Integer

10    full = True
20    For n = 0 To 2
30      grdreaction.Col = n
40      For Y = 0 To 5
50        grdreaction.Row = Y
60        If Trim$(grdreaction.Text) = "" Then full = False
70      Next
80    Next
90    If Not full Then
100     iMsg "Reactions not filled.", vbCritical
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   grdTarget.Col = 0
150   grdTarget.Row = 2
160   grdTarget.ColSel = 0
170   grdTarget.RowSel = 5
180   grdTarget.Clip = prevtarget

End Sub

Private Sub Form_Load()

      Dim n As Integer

10    For n = 0 To 2
20      grdreaction.ColAlignment(n) = 2
30      grdreaction.ColWidth(n) = grdreaction.Width / 3
40      grdTarget.ColAlignment(n) = 2
50      grdTarget.ColWidth(n) = grdTarget.Width / 3
60      grdTitle(0).ColAlignment(n) = 2
70      grdTitle(0).ColWidth(n) = grdTitle(0).Width / 3
80      grdTitle(1).ColAlignment(n) = 2
90      grdTitle(1).ColWidth(n) = grdTitle(1).Width / 3
100   Next

110   For n = 0 To 1
120     grdLotNos.ColAlignment(n) = 2
130     grdLotNos.ColWidth(n) = Choose(n + 1, 1000, 1000)
140   Next

End Sub

Private Sub grdlotnos_Click()

10    txtInput.Text = grdLotNos.Text
20    txtInput.SetFocus

End Sub

Private Sub grdreaction_Click()

      Dim t As String

10    t = grdreaction.Text
20    Select Case t
        Case "O": t = "+"
30      Case "+": t = ""
40      Case Else: t = "O"
50    End Select
60    grdreaction.Text = t

End Sub

Private Sub grdtarget_Click()

      Dim t As String

10    t = grdTarget.Text
20    Select Case t
        Case "O": t = "+"
30      Case "+": t = ""
40      Case Else: t = "O"
50    End Select
60    grdTarget.Text = t

End Sub

Private Sub grdtitle_Click(Index As Integer)

10    If Index = 0 Then
20      txttitinput = grdTitle(0)
30      txttitinput.SetFocus
40    Else
50      tInput = grdTitle(1)
60      tInput.SetFocus
70    End If

End Sub

Private Sub grdtitle_GotFocus(Index As Integer)

10    If Index = 0 Then
20      txttitinput = ""
30    Else
40      tInput = ""
50    End If

End Sub

Private Sub tInput_Change()

10    grdTitle(1) = tInput

End Sub

Private Sub txtinput_Change()

10    grdLotNos.Text = txtInput.Text

End Sub

Private Sub txttitinput_Change()

10    grdTitle(0) = txttitinput

End Sub

