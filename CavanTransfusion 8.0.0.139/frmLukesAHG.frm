VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLukesAHG 
   Caption         =   "NetAcquire --- AHG Quality Assurance"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   6780
   Icon            =   "frmLukesAHG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   6780
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   5310
      Width           =   5235
   End
   Begin MSComCtl2.DTPicker dtIgGExpiry 
      Height          =   315
      Left            =   3960
      TabIndex        =   13
      Top             =   1050
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
      TabIndex        =   12
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   38418
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1440
      Picture         =   "frmLukesAHG.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   4530
      Picture         =   "frmLukesAHG.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   765
      Left            =   3000
      Picture         =   "frmLukesAHG.frx":159E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5850
      Width           =   1155
   End
   Begin VB.TextBox txtIgG 
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   1020
      Width           =   2895
   End
   Begin VB.TextBox txtIgGC3d 
      Height          =   285
      Left            =   900
      TabIndex        =   2
      Top             =   540
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid grdReactions 
      Height          =   885
      Left            =   390
      TabIndex        =   0
      Top             =   4320
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1561
      _Version        =   393216
      Rows            =   3
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
      FormatString    =   $"frmLukesAHG.frx":1C08
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
      Height          =   1845
      Left            =   390
      TabIndex        =   6
      Top             =   2340
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   3254
      _Version        =   393216
      Rows            =   7
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
      FormatString    =   $"frmLukesAHG.frx":1C77
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
      Left            =   690
      TabIndex        =   16
      Top             =   6690
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   5340
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Index           =   0
      Left            =   3960
      TabIndex        =   11
      Top             =   330
      Width           =   420
   End
   Begin VB.Label lblLastEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   900
      TabIndex        =   10
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "IgG"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1050
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IgG C3d"
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Card Lot Number"
      Height          =   195
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   330
      Width           =   1200
   End
End
Attribute VB_Name = "frmLukesAHG"
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
170     grdLotNos.TextMatrix(4, 1) = tb!BlissLot & ""
180     grdLotNos.TextMatrix(4, 2) = Format$(tb!BlissExpiry, "dd/mm/yyyy")
190     grdLotNos.TextMatrix(5, 1) = tb!SalineLot & ""
200     grdLotNos.TextMatrix(5, 2) = Format$(tb!SalineExpiry, "dd/mm/yyyy")
210     grdLotNos.TextMatrix(6, 1) = tb!PBSBufferLot & ""
220     grdLotNos.TextMatrix(6, 2) = Format$(tb!PBSBufferExpiry, "dd/mm/yyyy")
230     grdReactions.TextMatrix(1, 1) = tb!Reaction11 & ""
240     grdReactions.TextMatrix(1, 2) = tb!Reaction12 & ""
250     grdReactions.TextMatrix(2, 1) = tb!Reaction21 & ""
260     grdReactions.TextMatrix(2, 2) = tb!Reaction22 & ""
270   Else
280     txtIgGC3d = ""
290     txtIgG = ""
300     dtC3dExpiry = Format$(Now, "dd/mm/yyyy")
310     dtIgGExpiry = Format$(Now, "dd/mm/yyyy")
320     grdLotNos.TextMatrix(1, 1) = ""
330     grdLotNos.TextMatrix(1, 2) = ""
340     grdLotNos.TextMatrix(2, 1) = ""
350     grdLotNos.TextMatrix(2, 2) = ""
360     grdLotNos.TextMatrix(3, 1) = ""
370     grdLotNos.TextMatrix(3, 2) = ""
380     grdReactions.TextMatrix(1, 1) = ""
390     grdReactions.TextMatrix(1, 2) = ""
400     grdReactions.TextMatrix(2, 1) = ""
410     grdReactions.TextMatrix(2, 2) = ""
420   End If

430   cmdSave.Enabled = True

440   Exit Sub

cmdLoad_Click_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmLukesAHG", "cmdLoad_Click", intEL, strES, sql


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

300   tb!BlissLot = grdLotNos.TextMatrix(4, 1)
310   If IsDate(grdLotNos.TextMatrix(4, 2)) Then
320     tb!BlissExpiry = Format$(grdLotNos.TextMatrix(4, 2), "dd/mm/yyyy")
330   Else
340     tb!BlissExpiry = Null
350   End If

360   tb!SalineLot = grdLotNos.TextMatrix(5, 1)
370   If IsDate(grdLotNos.TextMatrix(5, 2)) Then
380     tb!SalineExpiry = Format$(grdLotNos.TextMatrix(5, 2), "dd/mm/yyyy")
390   Else
400     tb!SalineExpiry = Null
410   End If

420   tb!PBSBufferLot = grdLotNos.TextMatrix(6, 1)
430   If IsDate(grdLotNos.TextMatrix(6, 2)) Then
440     tb!PBSBufferExpiry = Format$(grdLotNos.TextMatrix(6, 2), "dd/mm/yyyy")
450   Else
460     tb!PBSBufferExpiry = Null
470   End If

480   tb!Reaction11 = Left$(grdReactions.TextMatrix(1, 1), 1)
490   tb!Reaction12 = Left$(grdReactions.TextMatrix(1, 2), 1)
500   tb!Reaction21 = Left$(grdReactions.TextMatrix(2, 1), 1)
510   tb!Reaction22 = Left$(grdReactions.TextMatrix(2, 2), 1)
520   tb!Operator = UserName

530   tb.Update

540   cmdSave.Enabled = False

550   Unload Me

560   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

570   intEL = Erl
580   strES = Err.Description
590   LogError "frmLukesAHG", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()


      '10    EnsureColumnExistsBB "StLukesAHG", "BlissLot", "nvarchar(50)"
      '20    EnsureColumnExistsBB "StLukesAHG", "BlissExpiry", "DateTime"
      '30    EnsureColumnExistsBB "StLukesAHG", "SalineLot", "nvarchar(50)"
      '40    EnsureColumnExistsBB "StLukesAHG", "SalineExpiry", "DateTime"
      '50    EnsureColumnExistsBB "StLukesAHG", "PBSBufferLot", "nvarchar(50)"
      '60    EnsureColumnExistsBB "StLukesAHG", "PBSBufferExpiry", "DateTime"

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

90        cmdLoad.Enabled = True
100       grdLotNos.RowHeight(2) = 0

110   dtC3dExpiry = Format(Now, "dd/MM/yyyy")
120   dtIgGExpiry = Format(Now, "dd/MM/yyyy")

130   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmLukesAHG", "Form_Load", intEL, strES, sql

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


Private Sub txtIgGC3d_LostFocus()

      Dim Expiry As String
      Dim lot As String

10    If Len(txtIgGC3d) = 20 Then

20      If Mid$(txtIgGC3d, 7, 2) <> "22" Then
30        iMsg "This is not an IgG C3d (AHG) Card", vbCritical
40        If TimedOut Then Unload Me: Exit Sub
50        Exit Sub
60      End If
  
70      Expiry = Left$(txtIgGC3d, 2) & "/" & Mid$(txtIgGC3d, 3, 2) & "/" & Mid$(txtIgGC3d, 5, 2)
80      Expiry = Format$(Expiry, "dd/MMM/yyyy")
90      dtC3dExpiry = Expiry
  
100     lot = Mid$(txtIgGC3d, 15, 3)
110     txtIgGC3d = "AHC" & lot & "A"
120   End If

End Sub


