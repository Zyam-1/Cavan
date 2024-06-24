VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLukesQCOrthoWeeklyView 
   Caption         =   "NetAcquire --- Ortho BioVue Grouping Cards (Weekly) History"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12345
   Icon            =   "frmLukesQCOrthoWeeklyView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   12345
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   675
      Left            =   9810
      Picture         =   "frmLukesQCOrthoWeeklyView.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   795
      Left            =   3390
      Picture         =   "frmLukesQCOrthoWeeklyView.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   90
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   11010
      Picture         =   "frmLukesQCOrthoWeeklyView.frx":1376
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   8640
      Picture         =   "frmLukesQCOrthoWeeklyView.frx":19E0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3165
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38302
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   2625
      Left            =   60
      TabIndex        =   5
      Top             =   930
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   19
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
      AllowUserResizing=   1
      FormatString    =   $"frmLukesQCOrthoWeeklyView.frx":204A
   End
   Begin MSFlexGridLib.MSFlexGrid grdCardReactions 
      Height          =   2625
      Left            =   60
      TabIndex        =   6
      Top             =   3630
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   31
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
      AllowUserResizing=   1
      FormatString    =   $"frmLukesQCOrthoWeeklyView.frx":2160
   End
   Begin MSFlexGridLib.MSFlexGrid grdSeraReactions 
      Height          =   2625
      Left            =   60
      TabIndex        =   7
      Top             =   6330
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   19
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
      AllowUserResizing=   1
      FormatString    =   $"frmLukesQCOrthoWeeklyView.frx":2240
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   3450
      TabIndex        =   10
      Top             =   9060
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmLukesQCOrthoWeeklyView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

120   sql = "Select * from StLukesGroupingCards where " & _
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
230   LogError "frmLukesQCOrthoWeeklyView", "AskForComment", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim X As Integer
      Dim Y As Integer
      Dim n As Integer

10    On Error GoTo FillG_Error

20    With grdLotNos
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With
70    With grdCardReactions
80      .Rows = 2
90      .AddItem ""
100     .RemoveItem 1
110   End With
120   With grdSeraReactions
130     .Rows = 2
140     .AddItem ""
150     .RemoveItem 1
160   End With

170   sql = "Select * from StLukesGroupingCards where " & _
            "DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' order by DateTime desc"

180   Set tb = New Recordset
190   RecOpenServerBB 0, tb, sql
200   Do While Not tb.EOF
210     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            tb!A1rrLot & vbTab & _
            Format$(tb!A1rrExpiry, "dd/mm/yyyy") & vbTab & _
            tb!A2rrLot & vbTab & _
            Format$(tb!A2rrExpiry, "dd/mm/yyyy") & vbTab & _
            tb!BrrLot & vbTab & _
            Format$(tb!BrrExpiry, "dd/mm/yyyy") & vbTab & _
            tb!OR1wR1Lot & vbTab & _
            Format$(tb!OR1wR1Expiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiALot & vbTab & _
            Format$(tb!AntiAExpiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiBLot & vbTab & _
            Format$(tb!AntiBExpiry, "dd/mm/yyyy") & vbTab & _
            tb!AntiDLot & vbTab & _
            Format$(tb!AntiDExpiry, "dd/mm/yyyy") & vbTab & _
            tb!CardLotNumber & vbTab & _
            Format$(tb!CardExpiry, "dd/mm/yyyy")
220     grdLotNos.AddItem s
  
230     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            vbTab
240     For Y = 1 To 4
250       For X = 1 To 6
260         s = s & tb("C" & Format$(Y) & Format$(X)) & vbTab
270       Next
280       s = s & vbTab
290     Next
300     grdCardReactions.AddItem s
  
310     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            vbTab
320     For Y = 1 To 4
330       For X = 1 To 3
340         s = s & tb("S" & Format$(Y) & Format$(X)) & vbTab
350       Next
360       s = s & vbTab
370     Next
380     grdSeraReactions.AddItem s
  
390     tb.MoveNext
400   Loop

410   With grdLotNos
420     If .Rows > 2 Then .RemoveItem 1
430     .ColWidth(2) = 0
440     For n = 1 To .Rows - 1
450       If .TextMatrix(n, 2) <> "" Then
460         .ColWidth(2) = TextWidth("Comment ")
470         Exit For
480       End If
490     Next
500   End With
510   With grdCardReactions
520     If .Rows > 2 Then .RemoveItem 1
530     .ColWidth(2) = 0
540     For n = 1 To .Rows - 1
550       If .TextMatrix(n, 2) <> "" Then
560         .ColWidth(2) = TextWidth("Comment ")
570         Exit For
580       End If
590     Next
600   End With
610   With grdSeraReactions
620     If .Rows > 2 Then .RemoveItem 1
630     .ColWidth(2) = 0
640     For n = 1 To .Rows - 1
650       If .TextMatrix(n, 2) <> "" Then
660         .ColWidth(2) = TextWidth("Comment ")
670         Exit For
680       End If
690     Next
700   End With

710   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

720   intEL = Erl
730   strES = Err.Description
740   LogError "frmLukesQCOrthoWeeklyView", "FillG", intEL, strES, sql

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

60      PrintReportCavan

End Sub

Private Sub PrintReport(ByVal blnpreview As Boolean)

      Dim Y As Long

10    On Error GoTo PrintReport_Error

20    With frmPreviewRTF
30      .Dept = "TG" 'Grouping Cards
40      .AdjustPaperSize "A4port"
50      .Clear
60      .WriteFormattedText "     ;", , 20, , , "Courier New"
70      .WriteFormattedText "St Lukes Hospital Rathgar", 1, 20, vbRed, 1
80      .WriteText vbCrLf

90      .WriteFormattedText "Ortho BioVue Grouping Cards Report", 1, 14, vbBlack, 1
100     .WriteFormattedText "Lot Numbers(1)", 1, 10, vbBlack, 1
110     .WriteFormattedText "Date/Time         Operator   A1rr             A2rr             Brr              O R1wR1", 0, 8
120     For Y = 1 To grdLotNos.Rows - 1
130       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
140       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
150       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 3) & Space$(17), 17) & ";", , 8
160       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 5) & Space$(17), 17) & ";", , 8
170       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 7) & Space$(17), 17) & ";", , 8
180       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 9) & Space$(17), 17), , 8
190     Next
200     .WriteFormattedText "Lot Numbers(2)", 1, 10, vbBlack, 1
210     .WriteFormattedText "Date/Time         Operator   Anti A           Anti B           Anti D           Card  ", 0, 8
220     For Y = 1 To grdLotNos.Rows - 1
230       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
240       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
250       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 11) & Space$(17), 17) & ";", , 8
260       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 13) & Space$(17), 17) & ";", , 8
270       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 15) & Space$(17), 17) & ";", , 8
280       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 17) & Space$(17), 17), , 8
290     Next
300     .WriteText vbCrLf

310     .WriteFormattedText "Expiry Dates(1)", 1, 10, vbBlack, 1
320     .WriteFormattedText "Date/Time         Operator   A1rr             A2rr             Brr              O R1wR1", 0, 8
330     For Y = 1 To grdLotNos.Rows - 1
340       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
350       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
360       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 4) & Space$(17), 17) & ";", , 8
370       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 6) & Space$(17), 17) & ";", , 8
380       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 8) & Space$(17), 17) & ";", , 8
390       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 10) & Space$(17), 17), , 8
400     Next
410     .WriteFormattedText "Expiry Dates(2)", 1, 10, vbBlack, 1
420     .WriteFormattedText "Date/Time         Operator   Anti A           Anti B           Anti D           Card  ", 0, 8
430     For Y = 1 To grdLotNos.Rows - 1
440       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
450       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
460       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 12) & Space$(17), 17) & ";", , 8
470       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 14) & Space$(17), 17) & ";", , 8
480       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 16) & Space$(17), 17) & ";", , 8
490       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 18) & Space$(17), 17), , 8
500     Next
510     .WriteText vbCrLf

520     .WriteFormattedText "Card Reactions(1)", 1, 10, vbBlack, 1
530     .WriteFormattedText "Date/Time         Operator   A1rr>> A  B  AB D  D  ctr   A2rr>> A  B  AB D  D  ctr ", 0, 8
540     For Y = 1 To grdCardReactions.Rows - 1
550       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
560       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 1) & Space$(10), 10) & "        ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdcardreactions.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
570       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 4) & Space$(3), 3) & ";", , 8
580       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 5) & Space$(3), 3) & ";", , 8
590       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 6) & Space$(3), 3) & ";", , 8
600       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 7) & Space$(3), 3) & ";", , 8
610       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 8) & Space$(3), 3) & ";", , 8
620       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 9) & Space$(3), 3) & "          ;", , 8
630       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 11) & Space$(3), 3) & ";", , 8
640       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 12) & Space$(3), 3) & ";", , 8
650       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 13) & Space$(3), 3) & ";", , 8
660       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 14) & Space$(3), 3) & ";", , 8
670       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 15) & Space$(3), 3) & ";", , 8
680       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 16) & Space$(3), 3), , 8
690     Next
700     .WriteFormattedText "Card Reactions(2)", 1, 10, vbBlack, 1
710     .WriteFormattedText "Date/Time         Operator    Brr>> A  B  AB D  D  ctr OR1wR1>> A  B  AB D  D  ctr", 0, 8
720     For Y = 1 To grdCardReactions.Rows - 1
730       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
740       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 1) & Space$(10), 10) & "        ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdcardreactions.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
750       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 18) & Space$(3), 3) & ";", , 8
760       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 19) & Space$(3), 3) & ";", , 8
770       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 20) & Space$(3), 3) & ";", , 8
780       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 21) & Space$(3), 3) & ";", , 8
790       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 22) & Space$(3), 3) & ";", , 8
800       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 23) & Space$(3), 3) & "          ;", , 8
810       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 25) & Space$(3), 3) & ";", , 8
820       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 26) & Space$(3), 3) & ";", , 8
830       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 27) & Space$(3), 3) & ";", , 8
840       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 28) & Space$(3), 3) & ";", , 8
850       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 29) & Space$(3), 3) & ";", , 8
860       .WriteFormattedText Left$(grdCardReactions.TextMatrix(Y, 30) & Space$(3), 3), , 8
870     Next
880     .WriteText vbCrLf

890     .WriteFormattedText "Sera Reactions(1)", 1, 10, vbBlack, 1
900     .WriteFormattedText "Date/Time         Operator   A1rr>> Anti A Anti B Anti D   A2rr>> Anti A Anti B Anti D", 0, 8
910     For Y = 1 To grdSeraReactions.Rows - 1
920       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
930       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 1) & Space$(10), 10) & "        ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdserareactions.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
940       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 4) & Space$(7), 7) & ";", , 8
950       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 5) & Space$(7), 7) & ";", , 8
960       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 6) & Space$(16), 16) & ";", , 8
970       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 8) & Space$(7), 7) & ";", , 8
980       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 9) & Space$(7), 7) & ";", , 8
990       .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 10) & Space$(7), 7), , 8
1000    Next
1010    .WriteFormattedText "Sera Reactions(2)", 1, 10, vbBlack, 1
1020    .WriteFormattedText "Date/Time         Operator    Brr>> Anti A Anti B Anti D OR1wR1>> Anti A Anti B Anti D", 0, 8
1030    For Y = 1 To grdSeraReactions.Rows - 1
1040      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 8 'DateTime
1050      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 1) & Space$(10), 10) & "        ;", , 8 'Operator
      '    .WriteFormattedText Left$(grdserareactions.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 8 'Comment
1060      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 12) & Space$(7), 7) & ";", , 8
1070      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 13) & Space$(7), 7) & ";", , 8
1080      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 14) & Space$(16), 16) & ";", , 8
1090      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 16) & Space$(7), 7) & ";", , 8
1100      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 17) & Space$(7), 7) & ";", , 8
1110      .WriteFormattedText Left$(grdSeraReactions.TextMatrix(Y, 18) & Space$(7), 7), , 8
1120    Next


1130    If blnpreview Then
1140      .Show 1
1150    Else
1160      .PrintRTB
1170    End If

1180  End With

1190  Exit Sub

PrintReport_Error:

      Dim strES As String
      Dim intEL As Integer

1200  intEL = Erl
1210  strES = Err.Description
1220  LogError "frmLukesQCOrthoWeeklyView", "PrintReport", intEL, strES

End Sub

Private Sub PrintReportCavan()

      Dim Y As Long

10    Printer.Font.Name = "Courier New"
20    Printer.Font.Size = 16
30    Printer.ForeColor = vbRed
40    Printer.Print "Cavan General Hospital Transfusion Laboratory"
50    Printer.Print

60    Printer.ForeColor = vbBlack
70    Printer.Print "Ortho BioVue Grouping Cards Report"
80    Printer.Print "Lot Numbers(1)"
90    Printer.Font.Size = 8
100   Printer.Print "Date/Time         Operator   A1rr             A2rr             Brr              O R1wR1"
110   For Y = 1 To grdLotNos.Rows - 1
120     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
130     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
140     Printer.Print Left$(grdLotNos.TextMatrix(Y, 3) & Space$(17), 17);
150     Printer.Print Left$(grdLotNos.TextMatrix(Y, 5) & Space$(17), 17);
160     Printer.Print Left$(grdLotNos.TextMatrix(Y, 7) & Space$(17), 17);
170     Printer.Print Left$(grdLotNos.TextMatrix(Y, 9) & Space$(17), 17)
180   Next
190   Printer.Print

200   Printer.Font.Size = 16
210   Printer.Print "Lot Numbers(2)"
220   Printer.Font.Size = 8
230   Printer.Print "Date/Time         Operator   Anti A           Anti B           Anti D           Card  "
240   For Y = 1 To grdLotNos.Rows - 1
250     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
260     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
270     Printer.Print Left$(grdLotNos.TextMatrix(Y, 11) & Space$(17), 17);
280     Printer.Print Left$(grdLotNos.TextMatrix(Y, 13) & Space$(17), 17);
290     Printer.Print Left$(grdLotNos.TextMatrix(Y, 15) & Space$(17), 17);
300     Printer.Print Left$(grdLotNos.TextMatrix(Y, 17) & Space$(17), 17)
310   Next
320   Printer.Print

330   Printer.Font.Size = 16
340   Printer.Print "Expiry Dates(1)"
350   Printer.Font.Size = 8

360   Printer.Print "Date/Time         Operator   A1rr             A2rr             Brr              O R1wR1"
370   For Y = 1 To grdLotNos.Rows - 1
380     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
390     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
400     Printer.Print Left$(grdLotNos.TextMatrix(Y, 4) & Space$(17), 17);
410     Printer.Print Left$(grdLotNos.TextMatrix(Y, 6) & Space$(17), 17);
420     Printer.Print Left$(grdLotNos.TextMatrix(Y, 8) & Space$(17), 17);
430     Printer.Print Left$(grdLotNos.TextMatrix(Y, 10) & Space$(17), 17)
440   Next
450   Printer.Print

460   Printer.Font.Size = 16
470   Printer.Print "Expiry Dates(2)"
480   Printer.Font.Size = 8
490   Printer.Print "Date/Time         Operator   Anti A           Anti B           Anti D           Card  "
500   For Y = 1 To grdLotNos.Rows - 1
510     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
520     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
530     Printer.Print Left$(grdLotNos.TextMatrix(Y, 12) & Space$(17), 17);
540     Printer.Print Left$(grdLotNos.TextMatrix(Y, 14) & Space$(17), 17);
550     Printer.Print Left$(grdLotNos.TextMatrix(Y, 16) & Space$(17), 17);
560     Printer.Print Left$(grdLotNos.TextMatrix(Y, 18) & Space$(17), 17)
570   Next
580   Printer.Print

590   Printer.Font.Size = 16
600   Printer.Print "Card Reactions(1)"
610   Printer.Font.Size = 8
620   Printer.Print "Date/Time         Operator   A1rr>> A  B  AB D  D  ctr   A2rr>> A  B  AB D  D  ctr "
630   For Y = 1 To grdCardReactions.Rows - 1
640     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
650     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 1) & Space$(10), 10); 'Operator
660     Printer.Print "        ";
670     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 4) & Space$(3), 3);
680     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 5) & Space$(3), 3);
690     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 6) & Space$(3), 3);
700     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 7) & Space$(3), 3);
710     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 8) & Space$(3), 3);
720     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 9) & Space$(3), 3);
730     Printer.Print "          ";
740     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 11) & Space$(3), 3);
750     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 12) & Space$(3), 3);
760     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 13) & Space$(3), 3);
770     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 14) & Space$(3), 3);
780     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 15) & Space$(3), 3);
790     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 16) & Space$(3), 3)
800   Next
810   Printer.Print

820   Printer.Font.Size = 16
830   Printer.Print "Card Reactions(2)"
840   Printer.Font.Size = 8
850   Printer.Print "Date/Time         Operator    Brr>> A  B  AB D  D  ctr OR1wR1>> A  B  AB D  D  ctr"
860   For Y = 1 To grdCardReactions.Rows - 1
870     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
880     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 1) & Space$(10), 10); 'Operator
890     Printer.Print "        ";
900     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 18) & Space$(3), 3);
910     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 19) & Space$(3), 3);
920     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 20) & Space$(3), 3);
930     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 21) & Space$(3), 3);
940     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 22) & Space$(3), 3);
950     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 23) & Space$(3), 3);
960     Printer.Print "          ";
970     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 25) & Space$(3), 3);
980     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 26) & Space$(3), 3);
990     Printer.Print Left$(grdCardReactions.TextMatrix(Y, 27) & Space$(3), 3);
1000    Printer.Print Left$(grdCardReactions.TextMatrix(Y, 28) & Space$(3), 3);
1010    Printer.Print Left$(grdCardReactions.TextMatrix(Y, 29) & Space$(3), 3);
1020    Printer.Print Left$(grdCardReactions.TextMatrix(Y, 30) & Space$(3), 3)
1030  Next
1040  Printer.Print

1050  Printer.Font.Size = 16
1060  Printer.Print "Sera Reactions(1)"
1070  Printer.Font.Size = 8
1080  Printer.Print "Date/Time         Operator   A1rr>> Anti A Anti B Anti D   A2rr>> Anti A Anti B Anti D"
1090  For Y = 1 To grdSeraReactions.Rows - 1
1100    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
1110    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 1) & Space$(10), 10); 'Operator
1120    Printer.Print "          ";
1130    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 4) & Space$(7), 7);
1140    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 5) & Space$(7), 7);
1150    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 6) & Space$(16), 16);
1160    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 8) & Space$(7), 7);
1170    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 9) & Space$(7), 7);
1180    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 10) & Space$(7), 7)
1190  Next
1200  Printer.Print

1210  Printer.Font.Size = 16
1220  Printer.Print "Sera Reactions(2)"
1230  Printer.Font.Size = 8
1240  Printer.Print "Date/Time         Operator    Brr>> Anti A Anti B Anti D OR1wR1>> Anti A Anti B Anti D"
1250  For Y = 1 To grdSeraReactions.Rows - 1
1260    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
1270    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 1) & Space$(10), 10); 'Operator
1280    Printer.Print "          ";
1290    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 12) & Space$(7), 7);
1300    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 13) & Space$(7), 7);
1310    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 14) & Space$(16), 16);
1320    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 16) & Space$(7), 7);
1330    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 17) & Space$(7), 7);
1340    Printer.Print Left$(grdSeraReactions.TextMatrix(Y, 18) & Space$(7), 7)
1350  Next

1360  Printer.EndDoc

End Sub


Private Sub cmdRefresh_Click()

10    FillG

End Sub

Private Sub Form_Load()

10    dtFrom = Format(Now - 7, "dd/mm/yyyy")
20    dtTo = Format(Now, "dd/mm/yyyy")

30      cmdPreview.Enabled = False

End Sub

Private Sub grdCardReactions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    AskForComment grdCardReactions

End Sub


Private Sub grdLotNos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    AskForComment grdLotNos

End Sub


Private Sub grdSeraReactions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    AskForComment grdSeraReactions

End Sub


