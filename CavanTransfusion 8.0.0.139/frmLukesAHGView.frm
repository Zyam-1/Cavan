VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLukesAHGView 
   Caption         =   "NetAcquire --- AHG QC History"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12285
   Icon            =   "frmLukesAHGView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   12285
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Print Preview"
      Height          =   675
      Left            =   1530
      Picture         =   "frmLukesAHGView.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   795
      Left            =   1560
      Picture         =   "frmLukesAHGView.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdReactions 
      Height          =   2925
      Left            =   4110
      TabIndex        =   6
      Top             =   300
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   "<Date/Time              |<Operator      |<Comment |^IgG C3d / Anti D |^IgG / Anti D |^IgG C3d / AB Serum |^IgG / AB Serum "
   End
   Begin MSFlexGridLib.MSFlexGrid grdLotNos 
      Height          =   2925
      Left            =   180
      TabIndex        =   5
      Top             =   3390
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   17
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
      FormatString    =   $"frmLukesAHGView.frx":1376
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   735
      Left            =   510
      TabIndex        =   2
      Top             =   630
      Width           =   3165
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   38302
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   2730
      Picture         =   "frmLukesAHGView.frx":1494
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2430
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   330
      Picture         =   "frmLukesAHGView.frx":1AFE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2430
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   3360
      TabIndex        =   9
      Top             =   6450
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmLukesAHGView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim X As Long
      Dim Y As Long
      Dim n As Integer

10    On Error GoTo FillG_Error

20    With grdReactions
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With
70    With grdLotNos
80      .Rows = 2
90      .AddItem ""
100     .RemoveItem 1
110   End With

120   sql = "Select * from StLukesAHG where " & _
            "DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
            "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' order by DateTime desc"
130   Set tb = New Recordset
140   RecOpenServerBB 0, tb, sql

150   Do While Not tb.EOF
160     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            tb!C3dCardLot & vbTab & _
            tb!IgGCardLot & vbTab & _
            tb!AntiDLot & vbTab & _
            Format$(tb!AntiDExpiry, "dd/mm/yyyy") & vbTab & _
            tb!ABSerumLot & vbTab & _
            Format$(tb!ABSerumExpiry, "dd/mm/yyyy") & vbTab & _
            tb!OLot & vbTab & _
            Format$(tb!OExpiry, "dd/mm/yyyy") & vbTab & _
            tb!BlissLot & "" & vbTab & _
            Format$(tb!BlissExpiry, "dd/mm/yyyy") & vbTab & _
            tb!SalineLot & "" & vbTab & _
            Format$(tb!SalineExpiry, "dd/mm/yyyy") & vbTab & _
            tb!PBSBufferLot & "" & vbTab & _
            Format$(tb!PBSBufferExpiry, "dd/mm/yyyy")
170     grdLotNos.AddItem s
  
180     s = Format$(tb!DateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Operator & vbTab & _
            tb!Comment & vbTab & _
            tb!Reaction11 & vbTab & _
            tb!Reaction12 & vbTab & _
            tb!Reaction21 & vbTab & _
            tb!Reaction22 & ""
190     grdReactions.AddItem s
200     tb.MoveNext
210   Loop

220   For X = 3 To 6
230     For Y = 1 To grdReactions.Rows - 1
240       If grdReactions.TextMatrix(Y, X) = "N" Then
250         grdReactions.TextMatrix(Y, X) = "Not Tested"
260       End If
270     Next
280   Next

290   With grdReactions
300     If .Rows > 2 Then
310       .RemoveItem 1
320       .ColWidth(2) = 0
330       For n = 1 To .Rows - 1
340         If .TextMatrix(n, 2) <> "" Then
350           .ColWidth(2) = TextWidth("Comment ")
360           Exit For
370         End If
380       Next
390     End If
400   End With
410   With grdLotNos
420     If .Rows > 2 Then
430       .RemoveItem 1
440       .ColWidth(2) = 0
450       For n = 1 To .Rows - 1
460         If .TextMatrix(n, 2) <> "" Then
470           .ColWidth(2) = TextWidth("Comment ")
480           Exit For
490         End If
500       Next
510     End If
520   End With

530   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "frmLukesAHGView", "FillG", intEL, strES, sql


End Sub

Private Sub PrintReport(ByVal blnpreview As Boolean)

      Dim Y As Long
      Dim f As Form

10    On Error GoTo PrintReport_Error

20    Set f = New frmPreviewRTF

30    With f
40      .Dept = "TQ" 'Transfusion QC Report
50      .AdjustPaperSize "A4port"
60      .Clear
70      .WriteFormattedText "     ;", , 20, , , "Courier New"
80      .WriteFormattedText "St Lukes Hospital Rathgar", 1, 20, vbRed, 1
90      .WriteText vbCrLf
100     .WriteFormattedText "AHG QC History", 1, 14, vbBlack, 1
    
110     .WriteFormattedText "Date/Time         Operator   Comment              IgG C3d     IgG     IgG C3d/    IgG/ ", , 10
120     .WriteFormattedText "                                                  /Anti D   /Anti D   AB Serum  AB Serum ", , 10
130     For Y = 1 To grdReactions.Rows - 1
140       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
150       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 10 'Operator
160       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 2) & Space$(20), 20) & "   ;", , 10 'Comment
170       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 3) & " ", 1) & Space$(9) & ";", , 10
180       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 4) & " ", 1) & Space$(9) & ";", , 10
190       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 5) & " ", 1) & Space$(9) & ";", , 10
200       .WriteFormattedText Left$(grdReactions.TextMatrix(Y, 6) & " ", 1) & Space$(9)
210     Next
220     .WriteText vbCrLf
230     .WriteText vbCrLf

240     .WriteFormattedText "Lot Numbers", 1, 14, , 1
  
250     .WriteFormattedText "Date/Time         Operator   IgG C3d     IgG         Anti D      AB Serum    O R1wR1", , 10
260     .WriteFormattedText "                             Card Lot    Card Lot    Card Lot    Card Lot    Card Lot", , 10
270     For Y = 1 To grdLotNos.Rows - 1
280       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
290       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 10 'Operator
300       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 3) & Space$(12), 12) & ";", , 10
310       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 4) & Space$(12), 12) & ";", , 10
320       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 5) & Space$(12), 12) & ";", , 10
330       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 7) & Space$(12), 12) & ";", , 10
340       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 9) & Space$(12), 12), , 10
350     Next
360     .WriteText vbCrLf
370     .WriteText vbCrLf



380     .WriteFormattedText "Date/Time         Operator   Bliss       Saline      PBS Buffer", , 10
390     .WriteFormattedText "                             Card Lot    Card Lot    Card Lot", , 10
400     For Y = 1 To grdLotNos.Rows - 1
410       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
420       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & " ;", , 10 'Operator
430       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 11) & Space$(12), 12) & ";", , 10
440       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 13) & Space$(12), 12) & ";", , 10
450       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 15) & Space$(12), 12), , 10
460     Next
470     .WriteText vbCrLf
480     .WriteText vbCrLf



490     .WriteFormattedText "Expiry Dates", 1, 14, , 1
500     .WriteFormattedText "Date/Time         Operator       Anti D      AB Serum    O R1wR1", , 10
510     .WriteFormattedText "                                 Expiry      Expiry      Expiry ", , 10
520     For Y = 1 To grdLotNos.Rows - 1
530       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
540       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & "     ;" 'Operator
550       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 6) & Space$(12), 12) & ";"
560       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 8) & Space$(12), 12) & ";"
570       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 10) & Space$(12), 12) & ";"
580       .WriteText vbCrLf
590     Next

600     .WriteText vbCrLf
610     .WriteText vbCrLf
620     .WriteFormattedText "Date/Time         Operator       Bliss       Saline      PBS Buffer", , 10
630     .WriteFormattedText "                                 Expiry      Expiry      Expiry ", , 10
640     For Y = 1 To grdLotNos.Rows - 1
650       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18) & ";", , 10 'DateTime
660       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10) & "     ;" 'Operator
670       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 12) & Space$(12), 12) & ";"
680       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 14) & Space$(12), 12) & ";"
690       .WriteFormattedText Left$(grdLotNos.TextMatrix(Y, 16) & Space$(12), 12) & ";"
700       .WriteText vbCrLf
710     Next
  
  
720     If blnpreview Then
730       .Show 1
740     Else
750       .PrintRTB
760     End If
  
770   End With

780   Set f = Nothing

790   Exit Sub

PrintReport_Error:

      Dim strES As String
      Dim intEL As Integer

800   intEL = Erl
810   strES = Err.Description
820   LogError "frmLukesAHGView", "PrintReport", intEL, strES

End Sub

Private Sub PrintReportCavan()

      Dim Y As Long

10    Printer.Font.Name = "Courier New"
20    Printer.Font.Size = 16
30    Printer.ForeColor = vbRed
40    Printer.Print "Cavan General Hospital Transfusion Laboratory"
50    Printer.ForeColor = vbBlack
60    Printer.Print "AHG QC History"
70    Printer.Font.Size = 10
80    Printer.Print "Date/Time         Operator   Comment              IgG C3d     IgG     IgG C3d/    IgG/"
90    Printer.Print "                                                  /Anti D   /Anti D   AB Serum  AB Serum "
100   For Y = 1 To grdReactions.Rows - 1
110     Printer.Print Left$(grdReactions.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
120     Printer.Print Left$(grdReactions.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
130     Printer.Print Left$(grdReactions.TextMatrix(Y, 2) & Space$(20), 20); "    "; 'Comment
140     Printer.Print Left$(grdReactions.TextMatrix(Y, 3) & " ", 1) & Space$(9);
150     Printer.Print Left$(grdReactions.TextMatrix(Y, 4) & " ", 1) & Space$(9);
160     Printer.Print Left$(grdReactions.TextMatrix(Y, 5) & " ", 1) & Space$(9);
170     Printer.Print Left$(grdReactions.TextMatrix(Y, 6) & " ", 1) & Space$(9)
180   Next
190   Printer.Print
200   Printer.Print
  
210   Printer.Font.Size = 14
220   Printer.Print "Lot Numbers"
230   Printer.Font.Size = 10
  
240   Printer.Print "Date/Time         Operator   IgG C3d     IgG Card    Anti D      AB Serum    O R1wR1"
250   Printer.Print "                             Card Lot    Card Lot    Card Lot    Card Lot    Card Lot"
260   For Y = 1 To grdLotNos.Rows - 1
270     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
280     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); " "; 'Operator
290     Printer.Print Left$(grdLotNos.TextMatrix(Y, 3) & Space$(12), 12);
300     Printer.Print Left$(grdLotNos.TextMatrix(Y, 4) & Space$(12), 12);
310     Printer.Print Left$(grdLotNos.TextMatrix(Y, 5) & Space$(12), 12);
320     Printer.Print Left$(grdLotNos.TextMatrix(Y, 7) & Space$(12), 12);
330     Printer.Print Left$(grdLotNos.TextMatrix(Y, 9) & Space$(12), 12)
340   Next
350   Printer.Print
360   Printer.Print

370   Printer.Font.Size = 14
380   Printer.Print "Expiry Dates"
390   Printer.Font.Size = 10
400   Printer.Print "Date/Time         Operator       Anti D      AB Serum    O R1wR1"
410   Printer.Print "                                 Expiry      Expiry      Expiry "
420   For Y = 1 To grdLotNos.Rows - 1
430     Printer.Print Left$(grdLotNos.TextMatrix(Y, 0) & Space$(18), 18); 'DateTime
440     Printer.Print Left$(grdLotNos.TextMatrix(Y, 1) & Space$(10), 10); "   "; 'Operator
450     Printer.Print Left$(grdLotNos.TextMatrix(Y, 6) & Space$(12), 12);
460     Printer.Print Left$(grdLotNos.TextMatrix(Y, 8) & Space$(12), 12);
470     Printer.Print Left$(grdLotNos.TextMatrix(Y, 10) & Space$(12), 12);
480     Printer.Print
490   Next

500   Printer.EndDoc
  
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

Private Sub cmdRefresh_Click()

10    FillG

End Sub

Private Sub Form_Load()

10    dtFrom = Format(Now - 7, "dd/mm/yyyy")
20    dtTo = Format(Now, "dd/mm/yyyy")

30      cmdPreview.Enabled = False

End Sub


Private Sub grdLotNos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    AskForComment grdLotNos

End Sub


Private Sub grdReactions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    AskForComment grdReactions

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

120   sql = "Select * from StLukesAHG where " & _
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
230   LogError "frmLukesAHGView", "AskForComment", intEL, strES, sql


End Sub



