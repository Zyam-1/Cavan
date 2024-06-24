VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLukesPhenotype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Patient Phenotype Quality Assurance"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1500
      TabIndex        =   5
      Top             =   3120
      Width           =   5235
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   765
      Left            =   4050
      Picture         =   "frmLukesPhenotype.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3660
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   5580
      Picture         =   "frmLukesPhenotype.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3660
      Width           =   1155
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Previous"
      Enabled         =   0   'False
      Height          =   765
      Left            =   2490
      Picture         =   "frmLukesPhenotype.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3660
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   1605
      Left            =   750
      TabIndex        =   3
      Top             =   1335
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   2831
      _Version        =   393216
      Rows            =   6
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
      FormatString    =   $"frmLukesPhenotype.frx":133E
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
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblLastEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   750
      TabIndex        =   7
      Top             =   480
      Width           =   5985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Index           =   1
      Left            =   750
      TabIndex        =   6
      Top             =   3165
      Width           =   660
   End
End
Attribute VB_Name = "frmLukesPhenotype"
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

20    sql = "Select top 1 * from StLukesPhenotype " & _
            "Order by DateTime desc"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    If Not tb.EOF Then
60      txtComment = tb!Comment & ""
70      grdLotNos.TextMatrix(1, 1) = tb!AntiKLotNumber & ""
80      grdLotNos.TextMatrix(1, 2) = Format$(tb!AntiKExpiry, "dd/mm/yyyy")
90      grdLotNos.TextMatrix(2, 1) = tb!AntiE0LotNumber & ""
100     grdLotNos.TextMatrix(2, 2) = Format$(tb!AntiE0Expiry, "dd/mm/yyyy")
110     grdLotNos.TextMatrix(3, 1) = tb!AntiE1LotNumber & ""
120     grdLotNos.TextMatrix(3, 2) = Format$(tb!AntiE1Expiry, "dd/mm/yyyy")
130     grdLotNos.TextMatrix(4, 1) = tb!AntiC0LotNumber & ""
140     grdLotNos.TextMatrix(4, 2) = Format$(tb!AntiC0Expiry, "dd/mm/yyyy")
150     grdLotNos.TextMatrix(5, 1) = tb!AntiC1LotNumber & ""
160     grdLotNos.TextMatrix(5, 2) = Format$(tb!AntiC1Expiry, "dd/mm/yyyy")
  
170   Else
  
 
180     grdLotNos.TextMatrix(1, 1) = ""
190     grdLotNos.TextMatrix(1, 2) = ""
200     grdLotNos.TextMatrix(2, 1) = ""
210     grdLotNos.TextMatrix(2, 2) = ""
220     grdLotNos.TextMatrix(3, 1) = ""
230     grdLotNos.TextMatrix(3, 2) = ""
240     grdLotNos.TextMatrix(4, 1) = ""
250     grdLotNos.TextMatrix(4, 2) = ""
260     grdLotNos.TextMatrix(5, 1) = ""
270     grdLotNos.TextMatrix(5, 2) = ""
  
280   End If

290   cmdSave.Enabled = True

300   Exit Sub

cmdLoad_Click_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmLukesPhenotype", "cmdLoad_Click", intEL, strES, sql
End Sub

Private Sub cmdSave_Click()
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    sql = "Select top 1 * from StLukesPhenotype where " & _
            "DateTime = '01/01/2000'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    tb.AddNew
60    tb!Comment = txtComment
70    tb!DateTime = Format$(Now, "dd/mm/yyyy hh:mm:ss")
80    tb!AntiKLotNumber = grdLotNos.TextMatrix(1, 1)

90    If IsDate(grdLotNos.TextMatrix(1, 2)) Then
100     tb!AntiKExpiry = Format$(grdLotNos.TextMatrix(1, 2), "dd/mm/yyyy")
110   Else
120     tb!AntiKExpiry = Null
130   End If

140   tb!AntiE0LotNumber = grdLotNos.TextMatrix(2, 1)

150   If IsDate(grdLotNos.TextMatrix(2, 2)) Then
160     tb!AntiE0Expiry = Format$(grdLotNos.TextMatrix(2, 2), "dd/mm/yyyy")
170   Else
180     tb!AntiE0Expiry = Null
190   End If

200   tb!AntiE1LotNumber = grdLotNos.TextMatrix(3, 1)
210   If IsDate(grdLotNos.TextMatrix(3, 2)) Then
220     tb!AntiE1Expiry = Format$(grdLotNos.TextMatrix(3, 2), "dd/mm/yyyy")
230   Else
240     tb!AntiE1Expiry = Null
250   End If

260   tb!AntiC0LotNumber = grdLotNos.TextMatrix(4, 1)

270   If IsDate(grdLotNos.TextMatrix(4, 2)) Then
280     tb!AntiC0Expiry = Format$(grdLotNos.TextMatrix(4, 2), "dd/mm/yyyy")
290   Else
300     tb!AntiC0Expiry = Null
310   End If

320   tb!AntiC1LotNumber = grdLotNos.TextMatrix(5, 1)
330   If IsDate(grdLotNos.TextMatrix(5, 2)) Then
340     tb!AntiC1Expiry = Format$(grdLotNos.TextMatrix(5, 2), "dd/mm/yyyy")
350   Else
360     tb!AntiC1Expiry = Null
370   End If


380   tb!Operator = UserName

390   tb.Update

400   cmdSave.Enabled = False

410   Unload Me

420   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmLukesPhenotype", "cmdSave_Click", intEL, strES, sql
End Sub

Private Sub Form_Load()
       Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo Form_Load_Error

20    sql = "Select top 1 * from StLukesPhenotype " & _
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
110   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmLukesPhenotype", "Form_Load", intEL, strES, sql
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
      '    If .Row = 3 And Len(strIP) = 10 Then
      '      strIP = "8SS" & Right$(strIP, 3)
      '    End If
70      ElseIf .MouseCol = 2 Then
80        Set f = frmAskDate
90        If .TextMatrix(.Row, .Col) <> "" Then
100         f.DisplayDate = Format(.TextMatrix(.Row, .Col), "dd/MMM/yyyy")
110       Else
120         f.DisplayDate = Format(Now, "dd/MMM/yyyy")
130       End If
140       f.Show 1
150       strIP = f.DisplayDate
160       Set f = Nothing
170     End If
180     .TextMatrix(.Row, .Col) = strIP
190     cmdSave.Enabled = True
200   End With
End Sub
