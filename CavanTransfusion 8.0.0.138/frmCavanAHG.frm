VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCavanAHG 
   Caption         =   "NetAcquire --- AHG Quality Assurance"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "frmCavanAHG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIgGC3d 
      Height          =   285
      Left            =   900
      TabIndex        =   7
      Top             =   420
      Width           =   2895
   End
   Begin VB.TextBox txtIgG 
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   900
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   765
      Left            =   3000
      Picture         =   "frmCavanAHG.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5970
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   4530
      Picture         =   "frmCavanAHG.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5970
      Width           =   1155
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1440
      Picture         =   "frmCavanAHG.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5970
      Width           =   1155
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   5340
      Width           =   5235
   End
   Begin MSComCtl2.DTPicker dtIgGExpiry 
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   930
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   38418
   End
   Begin MSComCtl2.DTPicker dtC3dExpiry 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   420
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   38418
   End
   Begin MSFlexGridLib.MSFlexGrid grdReactions 
      Height          =   1035
      Left            =   510
      TabIndex        =   8
      Top             =   4020
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1826
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"frmCavanAHG.frx":1C08
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
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   1035
      Left            =   510
      TabIndex        =   9
      Top             =   2490
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1826
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   $"frmCavanAHG.frx":1C96
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   870
      TabIndex        =   16
      Top             =   6990
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Card Lot Number"
      Height          =   195
      Index           =   0
      Left            =   1740
      TabIndex        =   15
      Top             =   210
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IgG C3d"
      Height          =   195
      Left            =   270
      TabIndex        =   14
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "IgG"
      Height          =   195
      Left            =   600
      TabIndex        =   13
      Top             =   930
      Width           =   255
   End
   Begin VB.Label lblLastEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   900
      TabIndex        =   12
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   11
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   5370
      Width           =   660
   End
End
Attribute VB_Name = "frmCavanAHG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdLoad_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdLoad_Click_Error

20    sql = "Select top 1 * from StLukesAHG " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      txtComment = tb!Comment & ""
70      txtIgGC3d = tb!C3dCardLot & ""
80      txtIgG = tb!IgGCardLot & ""
90      dtC3dExpiry = tb!c3dexpiry & ""
100     dtIgGExpiry = tb!iggexpiry & ""
110     grdLotNos.TextMatrix(1, 1) = tb!AntiDLot & ""
120     grdLotNos.TextMatrix(1, 2) = Format$(tb!AntiDExpiry, "dd/mm/yyyy")
130     grdLotNos.TextMatrix(2, 1) = tb!ABSerumLot & ""
140     grdLotNos.TextMatrix(2, 2) = Format$(tb!ABSerumExpiry, "dd/mm/yyyy")
150     grdLotNos.TextMatrix(3, 1) = tb!OLot & ""
160     grdLotNos.TextMatrix(3, 2) = Format$(tb!OExpiry, "dd/mm/yyyy")
170     grdReactions.TextMatrix(1, 1) = tb!Reaction11 & ""
180     grdReactions.TextMatrix(1, 2) = tb!Reaction12 & ""
190     grdReactions.TextMatrix(2, 1) = tb!Reaction21 & ""
200     grdReactions.TextMatrix(2, 2) = tb!Reaction22 & ""
210     grdReactions.TextMatrix(3, 1) = tb!Reaction31 & ""
220     grdReactions.TextMatrix(3, 2) = tb!Reaction32 & ""
230   Else
240     txtIgGC3d = ""
250     txtIgG = ""
260     dtC3dExpiry = Format$(Now, "dd/mm/yyyy")
270     dtIgGExpiry = Format$(Now, "dd/mm/yyyy")
280     grdLotNos.TextMatrix(1, 1) = ""
290     grdLotNos.TextMatrix(1, 2) = ""
300     grdLotNos.TextMatrix(2, 1) = ""
310     grdLotNos.TextMatrix(2, 2) = ""
320     grdLotNos.TextMatrix(3, 1) = ""
330     grdLotNos.TextMatrix(3, 2) = ""
340     grdReactions.TextMatrix(1, 1) = ""
350     grdReactions.TextMatrix(1, 2) = ""
360     grdReactions.TextMatrix(2, 1) = ""
370     grdReactions.TextMatrix(2, 2) = ""
380     grdReactions.TextMatrix(3, 1) = ""
390     grdReactions.TextMatrix(3, 2) = ""
400   End If

410   cmdSave.Enabled = True

420   Exit Sub

cmdLoad_Click_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmCavanAHG", "cmdLoad_Click", intEL, strES, sql


End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    sql = "Select top 1 * from StLukesAHG where " & _
            "DateTime = '01/01/2000'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    tb.AddNew
60    tb!Comment = txtComment

70    tb!C3dCardLot = txtIgGC3d
80    tb!IgGCardLot = txtIgG
90    tb!c3dexpiry = Format$(dtC3dExpiry, "dd/mmm/yyyy")
100   tb!iggexpiry = Format$(dtIgGExpiry, "dd/mmm/yyyy")
110   tb!DateTime = Format$(Now, "dd/mm/yyyy hh:mm:ss")
120   tb!AntiDLot = grdLotNos.TextMatrix(1, 1)

130   If IsDate(grdLotNos.TextMatrix(1, 2)) Then
140     tb!AntiDExpiry = Format$(grdLotNos.TextMatrix(1, 2), "dd/mm/yyyy")
150   Else
160     tb!AntiDExpiry = Null
170   End If

180   tb!ABSerumLot = grdLotNos.TextMatrix(2, 1)

190   If IsDate(grdLotNos.TextMatrix(2, 2)) Then
200     tb!ABSerumExpiry = Format$(grdLotNos.TextMatrix(2, 2), "dd/mm/yyyy")
210   Else
220     tb!ABSerumExpiry = Null
230   End If

240   tb!OLot = grdLotNos.TextMatrix(3, 1)
250   If IsDate(grdLotNos.TextMatrix(3, 2)) Then
260     tb!OExpiry = Format$(grdLotNos.TextMatrix(3, 2), "dd/mm/yyyy")
270   Else
280     tb!OExpiry = Null
290   End If
300   tb!Reaction11 = Left$(grdReactions.TextMatrix(1, 1), 1)
310   tb!Reaction12 = Left$(grdReactions.TextMatrix(1, 2), 1)
320   tb!Reaction21 = Left$(grdReactions.TextMatrix(2, 1), 1)
330   tb!Reaction22 = Left$(grdReactions.TextMatrix(2, 2), 1)
340   tb!Reaction31 = Left$(grdReactions.TextMatrix(3, 1), 1)
350   tb!Reaction32 = Left$(grdReactions.TextMatrix(3, 2), 1)
360   tb!Operator = UserName

370   tb.Update

380   cmdSave.Enabled = False

390   Unload Me

400   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmCavanAHG", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo Form_Load_Error

20    sql = "Select top 1 * from StLukesAHG " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      s = "Last Entered by " & tb!Operator & vbCrLf & _
            "on " & Format$(tb!DateTime, "dd/mm/yyyy") & _
            " at " & Format$(tb!DateTime, "hh:mm:ss")
70      lblLastEntered = s
80    End If

90      cmdLoad.Enabled = True
100     grdLotNos.RowHeight(2) = 0

110   dtC3dExpiry = Format(Now, "dd/MM/yyyy")
120   dtIgGExpiry = Format(Now, "dd/MM/yyyy")

130   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmCavanAHG", "Form_Load", intEL, strES, sql


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Enabled Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub


Private Sub grdLotNos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim strIP As String
      Dim s As String
      Dim f As Form

10    With grdLotNos
20      If .MouseRow = 0 Then Exit Sub
  
30      If .MouseCol = 1 Then
40        s = "Enter Lot Number for " & .TextMatrix(.Row, 0)
50        strIP = iBOX(s, , .TextMatrix(.Row, .Col))
60        If TimedOut Then Unload Me: Exit Sub
70        If .Row = 3 And Len(strIP) = 10 Then
80          strIP = "8SS" & Right$(strIP, 3)
90        End If
100     ElseIf .MouseCol = 2 Then
110       Set f = frmAskDate
120       If .TextMatrix(.Row, .Col) <> "" Then
130         f.DisplayDate = Format(.TextMatrix(.Row, .Col), "dd/MMM/yyyy")
140       Else
150         f.DisplayDate = Format(Now, "dd/MMM/yyyy")
160       End If
170       f.Show 1
180       strIP = f.DisplayDate
190       Set f = Nothing
200     End If
210     .TextMatrix(.Row, .Col) = strIP
220     cmdSave.Enabled = True
230   End With

End Sub


Private Sub grdReactions_Click()

10    With grdReactions
20      If .MouseRow <> 0 And .MouseCol <> 0 Then
30        Select Case .TextMatrix(.Row, .Col)
            Case "": .TextMatrix(.Row, .Col) = "0"
40          Case "0": .TextMatrix(.Row, .Col) = "1"
50          Case "1": .TextMatrix(.Row, .Col) = "2"
60          Case "2": .TextMatrix(.Row, .Col) = "3"
70          Case "3": .TextMatrix(.Row, .Col) = "4"
80          Case "4": .TextMatrix(.Row, .Col) = "+"
90          Case "+": .TextMatrix(.Row, .Col) = "Not Tested"
100         Case Else: .TextMatrix(.Row, .Col) = ""
110       End Select
120       cmdSave.Enabled = True
130     End If
140   End With

End Sub


Private Sub txtIgG_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


Private Sub txtIgGC3d_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


