VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLukesCentrifugeView 
   Caption         =   "NetAcquire --- Centrifuge Speed / Temperature History"
   ClientHeight    =   5010
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   8610
   Icon            =   "frmLukesCentrifugeView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   675
      Left            =   4830
      Picture         =   "frmLukesCentrifugeView.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   795
      Left            =   3570
      Picture         =   "frmLukesCentrifugeView.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   6090
      Picture         =   "frmLukesCentrifugeView.frx":1376
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   7320
      Picture         =   "frmLukesCentrifugeView.frx":19E0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   3615
      Left            =   270
      TabIndex        =   3
      Top             =   1020
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   10
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
      ScrollBars      =   2
      FormatString    =   $"frmLukesCentrifugeView.frx":204A
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   735
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   3165
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38302
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38302
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   1530
      TabIndex        =   8
      Top             =   4740
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmLukesCentrifugeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Integer

10    On Error GoTo FillG_Error

20    grdReport.Rows = 2
30    grdReport.AddItem ""
40    grdReport.RemoveItem 1

50    sql = "Select * from StLukesCentrifuge where " & _
            "DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' order by dateTime desc"

60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = Format(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            Left$(tb!Operator, 10) & vbTab & _
            tb!Comment & vbTab & _
            tb!Cent1Phase1 & vbTab & _
            tb!Cent1Phase2 & vbTab & _
            tb!Cent2Phase1 & vbTab & _
            tb!Cent2Phase2 & vbTab & _
            tb!BlockL & vbTab & _
            tb!BlockR & vbTab & _
            tb!BlockS & ""
100     grdReport.AddItem s
110     tb.MoveNext
120   Loop

130   If grdReport.Rows > 2 Then
140     grdReport.RemoveItem 1
150     grdReport.ColWidth(2) = 0
160     For n = 1 To grdReport.Rows - 1
170       If grdReport.TextMatrix(n, 2) <> "" Then
180         grdReport.ColWidth(2) = TextWidth("Comment ")
190         Exit For
200       End If
210     Next
220   End If

230   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmLukesCentrifugeView", "FillG", intEL, strES, sql


End Sub

Private Sub PrintReport(ByVal blnpreview As Boolean)

      Dim Y As Long

10    On Error GoTo PrintReport_Error

20    With frmPreviewRTF
30      .Dept = "TC" 'Centrifuge QC Report
40      .AdjustPaperSize "A4port"
50      .Clear
60      .WriteFormattedText "     ;", , 20, , , "Courier New"
70      .WriteFormattedText "St Lukes Hospital Rathgar", 1, 20, vbRed, 1
80      .WriteText vbCrLf
90      .WriteFormattedText "NetAcquire - Transfusion", 1, 14, vbBlack, 1
100     .WriteFormattedText "Centrifuge and Temperature QC History", 1, 14, vbBlack, 1
110     .WriteText vbCrLf
    
120     .WriteFormattedText "Date/Time         Operator   Comment              C1    C1    C2    C2    B-L  B-R  B-S", , 10
130     .WriteFormattedText "                                                  Ph1   Ph2   Ph1   Ph2", , 10

140     For Y = 1 To grdReport.Rows - 1
150       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
160       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 10 'Operator
170       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 2) & Space$(20), 20) & " ;", , 10 'Comment
180       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 3) & Space$(6), 6) & ";", , 10
190       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 4) & Space$(6), 6) & ";", , 10
200       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 5) & Space$(6), 6) & ";", , 10
210       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 6) & Space$(6), 6) & ";", , 10
220       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 7) & Space$(5), 5) & ";", , 10
230       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 8) & Space$(5), 5) & ";", , 10
240       .WriteFormattedText Left$(grdReport.TextMatrix(Y, 9) & Space$(5), 5), , 10
250     Next
    
260     If blnpreview Then
270       .Show 1
280     Else
290       .PrintRTB
300     End If
310   End With

320   Exit Sub

PrintReport_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmLukesCentrifugeView", "PrintReport", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPreview_Click()

10    If Not SetFormPrinter() Then
20      iMsg "Can't find printer!", vbExclamation
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50    End If

60    PrintReport True

End Sub

Private Sub cmdPrint_Click()

10    If Not SetFormPrinter() Then
20      iMsg "Can't find printer!", vbExclamation
30      If TimedOut Then Unload Me: Exit Sub
40      Exit Sub
50    End If

60    PrintReport False

End Sub

Private Sub cmdRefresh_Click()

10    FillG

End Sub

Private Sub Form_Load()

10    dtFrom = Format(Now - 7, "dd/mm/yyyy")
20    dtTo = Format(Now, "dd/mm/yyyy")

End Sub


Private Sub grdReport_Click()

10    AskForComment grdReport

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

120   sql = "Select * from StLukesCentrifuge where " & _
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
230   LogError "frmLukesCentrifugeView", "AskForComment", intEL, strES, sql


End Sub

