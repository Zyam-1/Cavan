VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmArchive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   855
      Left            =   11760
      TabIndex        =   10
      Top             =   1140
      Width           =   1545
      Begin VB.OptionButton optSearch 
         Caption         =   "Products"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   510
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Patients"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      Height          =   855
      Left            =   11760
      TabIndex        =   7
      Top             =   180
      Width           =   1575
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   495
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Only Changes"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchive.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6180
      Width           =   1200
   End
   Begin VB.TextBox txtSampleId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11610
      TabIndex        =   4
      Top             =   3045
      Width           =   1845
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchive.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1200
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8505
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15002
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmArchive.frx":1D94
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchive.frx":1E16
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   7680
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample ID"
      Height          =   360
      Left            =   11610
      TabIndex        =   5
      Top             =   2715
      Width           =   1845
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pTableNameAudit As String




Private Sub FillCurrent()

      Dim sql As String
      Dim tb As Recordset
      Dim dt As String
      Dim Op As String
      Dim FieldChanged As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ADT As Integer
      Dim AB As Integer
      Dim NameDisplayed As Boolean
      Dim yy As Integer
      Dim Previous As String
      Dim SearchBy As String

10    On Error GoTo FillCurrent_Error

20    rtb.Text = ""
30    rtb.SelFontSize = 12

40    If Trim$(txtSampleId) = "" Then Exit Sub

50    If optSearch(0) Then
60      SearchBy = "LabNumber"
70    Else
80      SearchBy = "ISBT128"
90    End If

100   sql = "SELECT 1 Tag, *, '1/1/2030' ArchiveDateTime, 'u' ArchivedBy FROM " & pTableName & " WHERE " & _
            SearchBy & " = '" & txtSampleId & "' " & _
            "UNION " & _
            "SELECT 2 Tag, * FROM " & pTableNameAudit & " WHERE " & _
            SearchBy & " = '" & txtSampleId & "' " & _
            "ORDER BY ArchiveDateTime DESC"
110   Set tb = New Recordset
120   RecOpenClientBB 0, tb, sql
130   If tb.EOF Then
140     rtb.SelText = "No Current Record found." & vbCrLf
150     Exit Sub
160   End If

170   ReDim Records(0 To tb.RecordCount - 1, 0 To tb.Fields.Count - 1)
180   ReDim titles(0 To tb.Fields.Count - 1)
190   Y = -1
200   Do While Not tb.EOF
210     Y = Y + 1
220     For X = 0 To tb.Fields.Count - 1
230       titles(X) = tb.Fields(X).Name
240       Records(Y, X) = tb.Fields(X).Value
250     Next
260     tb.MoveNext
270   Loop

280   dt = "Unknown            "
290   tb.MoveFirst
300   If Not IsNull(tb!DateTime) Then
310     If IsDate(tb!DateTime) Then
320       dt = Format$(tb!DateTime, "dd/MM/yyyy HH:nn:ss")
330     End If
340   End If
350   Op = "Unknown"
360   If Trim$(tb!Operator & "") <> "" Then
370     Op = tb!Operator
380   End If

390   rtb.SelBold = True
400   rtb.SelColor = vbBlue
410   rtb.SelFontSize = 12
420   rtb.SelUnderline = True
430   rtb.SelText = "Current Record entered by " & Op & " at " & dt & vbCrLf & vbCrLf

440   For X = 1 To UBound(titles) - 2
450     FieldChanged = False
460     For Y = 1 To UBound(Records)
470       If Trim$(Records(Y, X) & "") & "" <> Trim$(Records(0, X) & "") Then
480         FieldChanged = True
490         Exit For
500       End If
510     Next
520     If optShow(0) Or (optShow(1) And FieldChanged) Then
530       rtb.SelFontName = "Courier New"
540       rtb.SelBold = True
550       rtb.SelColor = vbBlue
560       rtb.SelFontSize = 12
570       rtb.SelText = Left$(titles(X) & Space$(20), 20)
580       rtb.SelFontSize = 12
590       rtb.SelBold = True
600     End If
610     If FieldChanged Then
620       rtb.SelColor = vbRed
630       rtb.SelText = Left$(Records(0, X) & Space$(25), 25)
640       rtb.SelText = " See below for changes"
650     Else
660       If optShow(0) Then
670         rtb.SelText = Left$(Records(0, X) & Space$(25), 25)
680       End If
690     End If
700     If optShow(0) Or (optShow(1) And FieldChanged) Then
710       rtb.SelText = vbCrLf
720     End If
730   Next

740   rtb.SelText = vbCrLf
750   rtb.SelText = vbCrLf
760   rtb.SelFontSize = 16
770   rtb.SelColor = vbBlack
780   rtb.SelText = String(40, "-") & vbCrLf
790   rtb.SelFontSize = 16
800   rtb.SelColor = vbBlack
810   rtb.SelText = "Audit Records:" & vbCrLf


820   If UBound(Records) = 0 Then
830     rtb.SelFontSize = 16
840     rtb.SelColor = vbBlack
850     rtb.SelText = "No Changes Made"
860   Else
870     For X = UBound(titles) - 1 To UBound(titles)
880       If UCase$(titles(X)) = "ARCHIVEDATETIME" Then
890         ADT = X
900       End If
910       If UCase$(titles(X)) = "ARCHIVEDBY" Then
920         AB = X
930       End If
940     Next
  
950     For X = 1 To UBound(titles)
960       FieldChanged = False
970       If X <> ADT And X <> AB Then
980         For Y = 1 To UBound(Records)
990           If Trim$(Records(Y, X) & "") <> Trim$(Records(0, X) & "") Then
1000            FieldChanged = True
1010            Exit For
1020          End If
1030        Next
1040      End If
1050      If FieldChanged Then
1060        Previous = Trim$(Records(0, X) & "")
1070        NameDisplayed = False
1080        For yy = 1 To UBound(Records)
1090          If Previous <> Trim$(Records(yy, X) & "") Then
1100            If Not NameDisplayed Then
1110              rtb.SelText = vbCrLf
1120              rtb.SelBold = True
1130              rtb.SelColor = vbBlue
1140              rtb.SelFontSize = 12
1150              rtb.SelUnderline = True
1160              rtb.SelText = titles(X) & vbCrLf
1170              NameDisplayed = True
1180            End If
    
1190            rtb.SelFontSize = 12
1200            rtb.SelColor = vbBlack
1210            rtb.SelText = Records(yy, ADT) & " "
1220            rtb.SelColor = vbBlue
1230            rtb.SelText = Records(yy, AB) & ""
1240            rtb.SelColor = vbBlack
1250            rtb.SelText = " Changed "
1260            rtb.SelColor = vbRed
1270            rtb.SelBold = True
1280            If Records(yy, X) = "" Then
1290              rtb.SelText = "<Blank> "
1300            Else
1310              rtb.SelText = Records(yy, X) & ""
1320            End If
1330            rtb.SelColor = vbBlack
1340            rtb.SelBold = False
1350            rtb.SelText = " to "
1360            rtb.SelBold = True
1370            rtb.SelColor = vbRed
1380            If Trim$(Previous) = "" Then
1390              rtb.SelText = "<Blank>" & vbCrLf
1400            Else
1410              rtb.SelText = Previous & vbCrLf
1420            End If
1430            Previous = Trim$(Records(yy, X) & "")
1440          End If
1450        Next
1460      End If
1470    Next
  
1480    If NameDisplayed Then
1490      rtb.SelText = vbCrLf
1500    End If

1510  End If

1520  Exit Sub

FillCurrent_Error:

      Dim strES As String
      Dim intEL As Integer

1530  intEL = Erl
1540  strES = Err.Description
1550  LogError "frmArchive", "FillCurrent", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdPrint_Click()

10    On Error GoTo cmdPrint_Click_Error

20    rtb.SelStart = 0
30    rtb.SelLength = 10000000#
40    rtb.SelPrint Printer.hDC

50    Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmArchive", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdStart_Click()

10    FillCurrent

End Sub

Public Property Let TableName(ByVal sNewValue As String)

10    pTableName = sNewValue
20    pTableNameAudit = sNewValue & "Audit"

End Property
Public Property Let SampleID(ByVal sNewValue As String)

10    txtSampleId = sNewValue

End Property


Private Sub optSearch_Click(Index As Integer)

10    If Index = 0 Then
20      pTableName = "PatientDetails"
30      lblTitle = "SampleID"
40    Else
50      pTableName = "Latest"
60      lblTitle = "Pack Number"
70    End If

80    pTableNameAudit = pTableName & "Audit"

90    rtb.Text = ""
100   txtSampleId = ""

End Sub


Private Sub optShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FillCurrent

End Sub

