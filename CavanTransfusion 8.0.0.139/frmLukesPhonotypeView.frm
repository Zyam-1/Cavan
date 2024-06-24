VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLukesPhenotypeView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Phonotype QC History"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   3900
      Picture         =   "frmLukesPhonotypeView.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   6750
      Picture         =   "frmLukesPhonotypeView.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1275
      Left            =   330
      TabIndex        =   2
      Top             =   180
      Width           =   3255
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   1860
         Picture         =   "frmLukesPhonotypeView.frx":0CD4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93388801
         CurrentDate     =   38302
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   93388801
         CurrentDate     =   38302
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   795
      Left            =   5325
      Picture         =   "frmLukesPhonotypeView.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   4185
      Left            =   300
      TabIndex        =   1
      Top             =   1620
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   13
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmLukesPhonotypeView.frx":1780
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   7
      Top             =   6000
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmLukesPhenotypeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim x As Long
      Dim Y As Long
      Dim n As Integer

10    On Error GoTo FillG_Error


20    With grdLotNos
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With

70    sql = "Select * from StLukesPhenotype where " & _
            "DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' order by DateTime desc"
80    Set tb = New Recordset
90    RecOpenServerBB 0, tb, sql

100   Do While Not tb.EOF
110     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            tb!AntiKLotNumber & vbTab & _
            Format$(tb!AntiKExpiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiE0LotNumber & vbTab & _
            Format$(tb!AntiE0Expiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiE1LotNumber & vbTab & _
            Format$(tb!AntiE1Expiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiC0LotNumber & vbTab & _
            Format$(tb!AntiC0Expiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiC1LotNumber & vbTab & _
            Format$(tb!AntiC1Expiry, "dd/mm/yyyy")

      
120     grdLotNos.AddItem s
  
  
130     tb.MoveNext
140   Loop


150   With grdLotNos
160     If .Rows > 2 Then
170       .RemoveItem 1
180       .ColWidth(2) = 0
190       For n = 1 To .Rows - 1
200         If .TextMatrix(n, 2) <> "" Then
210           .ColWidth(2) = TextWidth("Comment ")
220           Exit For
230         End If
240       Next
250     End If
260   End With

270   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmLukesPhenotypeView", "FillG", intEL, strES, sql


End Sub
'Private Sub PrintReport(ByVal blnpreview As Boolean)
'
'Dim y As Long
'Dim f As Form
'
'On Error GoTo PrintReport_Error
'
'Set f = New frmPreviewRTF
'
'With f
'  .Dept = "TQ" 'Transfusion QC Report
'  .AdjustPaperSize "A4port"
'  .Clear
'  .WriteFormattedText "     ;", , 20, , , "Courier New"
'  .WriteFormattedText "St Lukes Hospital Rathgar", 1, 20, vbRed, 1
'  .WriteText vbCrLf
'  .WriteFormattedText "Patient Phenotype QC History", 1, 14, vbBlack, 1
'
'  .WriteFormattedText "Date/Time         Operator   Comment              IgG C3d     IgG     IgG C3d/    IgG/ ", , 10
'  .WriteFormattedText "                                                  /Anti D   /Anti D   AB Serum  AB Serum ", , 10
'  For y = 1 To grdReactions.Rows - 1
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 0) & Space$(18), 18) & ";", , 10 'DateTime
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 1) & Space$(10), 10) & " ;", , 10 'Operator
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 2) & Space$(20), 20) & "   ;", , 10 'Comment
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 3) & " ", 1) & Space$(9) & ";", , 10
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 4) & " ", 1) & Space$(9) & ";", , 10
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 5) & " ", 1) & Space$(9) & ";", , 10
'    .WriteFormattedText Left$(grdReactions.TextMatrix(y, 6) & " ", 1) & Space$(9)
'  Next
'  .WriteText vbCrLf
'  .WriteText vbCrLf
'
'  .WriteFormattedText "Lot Numbers", 1, 14, , 1
'
'  .WriteFormattedText "Date/Time         Operator   IgG C3d     IgG Card    Anti D      AB Serum    O R1wR1", , 10
'  .WriteFormattedText "                             Card Lot    Card Lot    Card Lot    Card Lot    Card Lot", , 10
'  For y = 1 To grdLotNos.Rows - 1
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 0) & Space$(18), 18) & ";", , 10 'DateTime
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 1) & Space$(10), 10) & " ;", , 10 'Operator
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 3) & Space$(12), 12) & ";", , 10
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 4) & Space$(12), 12) & ";", , 10
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 5) & Space$(12), 12) & ";", , 10
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 7) & Space$(12), 12) & ";", , 10
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 9) & Space$(12), 12), , 10
'  Next
'  .WriteText vbCrLf
'  .WriteText vbCrLf
'
'  .WriteFormattedText "Expiry Dates", 1, 14, , 1
'  .WriteFormattedText "Date/Time         Operator       Anti D      AB Serum    O R1wR1", , 10
'  .WriteFormattedText "                                 Expiry      Expiry      Expiry ", , 10
'  For y = 1 To grdLotNos.Rows - 1
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 0) & Space$(18), 18) & ";", , 10 'DateTime
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 1) & Space$(10), 10) & "     ;" 'Operator
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 6) & Space$(12), 12) & ";"
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 8) & Space$(12), 12) & ";"
'    .WriteFormattedText Left$(grdLotNos.TextMatrix(y, 10) & Space$(12), 12) & ";"
'    .WriteText vbCrLf
'  Next
'
'  If blnpreview Then
'    .Show 1
'  Else
'    .PrintRTB
'  End If
'
'End With
'
'Set f = Nothing
'
'Exit Sub
'
'PrintReport_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmLukesAHGView", "PrintReport", intEL, strES
'
'End Sub
Private Sub cmdCancel_Click()
10    Unload Me
End Sub

Private Sub cmdRefresh_Click()
10    FillG
End Sub

Private Sub Form_Load()
10    dtFrom = Format(Now - 7, "dd/mm/yyyy")
20    dtTo = Format(Now, "dd/mm/yyyy")

30    If HospName(0) = "Cavan" Then
40      cmdPreview.Enabled = False
50    Else
60      cmdPreview.Enabled = True
70    End If

End Sub


Private Sub grdLotNos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
10    AskForComment grdLotNos
End Sub

Private Sub AskForComment(ByVal g As MSFlexGrid)

      Dim Comment As String
      Dim DateTime As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo AskForComment_Error

20    If g.MouseRow = 0 Then Exit Sub

30    DateTime = g.TextMatrix(g.MouseRow, 0)
40    Comment = g.TextMatrix(g.Row, 2)

50    Answer = iMsg("Enter Comment for " & DateTime & "?", vbQuestion + vbYesNo)
60    If TimedOut Then Unload Me: Exit Sub
70    If Answer = vbNo Then
80      Exit Sub
90    End If

100   Comment = iBOX("Enter Comment", , Comment)
110   If TimedOut Then Unload Me: Exit Sub

120   sql = "Select * from StLukesPhenotype where " & _
            "DateTime = '" & Format$(DateTime, "dd/mmm/yyyy hh:mm:ss") & "'"

130   Set tb = New Recordset
140   RecOpenServerBB 0, tb, sql
150   If Not tb.EOF Then
160     tb!Comment = Comment
170   End If
180   tb.Update

190   FillG

200   Exit Sub

AskForComment_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmLukesPhenotypeView", "AskForComment", intEL, strES, sql


End Sub

