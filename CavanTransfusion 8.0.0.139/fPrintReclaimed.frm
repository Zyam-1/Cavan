VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form fPrintReclaimed 
   Caption         =   "NetAcquire - Reclaimed Product"
   ClientHeight    =   8340
   ClientLeft      =   180
   ClientTop       =   720
   ClientWidth     =   13635
   ControlBox      =   0   'False
   Icon            =   "fPrintReclaimed.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   13635
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   10560
      Picture         =   "fPrintReclaimed.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   135
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   12195
      Picture         =   "fPrintReclaimed.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7050
      Left            =   180
      TabIndex        =   0
      Top             =   930
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   12435
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"fPrintReclaimed.frx":159E
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
      Left            =   225
      TabIndex        =   4
      Top             =   8040
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1020
      Picture         =   "fPrintReclaimed.frx":1688
      Top             =   450
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Patient Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   180
      Width           =   2205
   End
End
Attribute VB_Name = "fPrintReclaimed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim sn As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from reclaimed order by name"
60    Set sn = New Recordset
70    RecOpenServerBB 0, sn, sql
80    Do While Not sn.EOF
90      s = sn!Name & vbTab & _
            sn!Typenex & vbTab & _
            sn!DoB & vbTab & _
            sn!Chart & vbTab & _
            sn!Ward & vbTab & _
            sn!Unit & vbTab & _
            sn!Group & vbTab & _
            sn!Product & vbTab & _
            sn!xmdate & vbTab & _
            sn!Operator & vbTab & _
            Format$(sn!DateTime, "dd/mm/yyyy hh:nn:ss")
100     g.AddItem s
110     sn.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fPrintReclaimed", "FillG", intEL, strES, sql

End Sub

Private Sub PrintReclaimed()

      Dim n As Integer
      Dim strName As String
      Dim strTypenex As String
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim Count As Integer
      Dim X As Integer
      Dim LineOfInterest As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim SampleID As String

10    On Error GoTo PrintReclaimed_Error
  
20    LineOfInterest = 0
30    g.Col = 0
40    For n = 1 To g.Rows - 1
50      g.Row = n
60      If g.CellBackColor = vbRed Then
70        LineOfInterest = n
80        Exit For
90      End If
100   Next
110   If LineOfInterest = 0 Then Exit Sub

      'Find labNumber/SampleId for unit/batch number been reclaimed
      'Note would be easier to save labnumber in Reclaimed table.

120   If IsProductBloodUnit(g.TextMatrix(LineOfInterest, 7)) Then
130       sql = "SELECT Top 1 LabNumber FROM Product WHERE "
          If Len(Trim$(g.TextMatrix(LineOfInterest, 5))) > 13 Then 'Is ISBT128 number
               sql = sql & "ISBT128 = '" & g.TextMatrix(LineOfInterest, 5) & "' "
          Else 'Codabar number
               sql = sql & "Number = '" & g.TextMatrix(LineOfInterest, 5) & "' "
          End If
          sql = sql & "AND LabNumber <> '' AND DateTime < '" & Format(g.TextMatrix(LineOfInterest, 10), "dd/mmm/yyyy hh:mm:ss") & "'" & _
                " ORDER BY DateTime DESC"
                
140   Else
150       sql = "SELECT top 1 labnumber from PatientDetails, Batchdetails  where PatientDetails.labnumber = batchdetails.sampleid " & _
          "and PatientDetails.name = '" & AddTicks(g.TextMatrix(LineOfInterest, 3)) & "'"
160       If Len(g.TextMatrix(LineOfInterest, 2)) > 0 Then 'DOB
170           sql = sql & " and PatientDetails.DOB = '" & Format(g.TextMatrix(LineOfInterest, 2), "dd/MMM/yyyy") & "'"
180       End If
190       If Len(g.TextMatrix(LineOfInterest, 3)) > 0 Then 'Chart
200           sql = sql & " and PatientDetails.PatNum = '" & g.TextMatrix(LineOfInterest, 3) & "'"
210       End If
220       sql = sql & " and Batchdetails.BatchNumber = '" & g.TextMatrix(LineOfInterest, 5) & "'"
230   End If
    
240   Set tb = New Recordset
250   RecOpenServerBB 0, tb, sql
260   If Not tb.EOF Then
270       SampleID = tb!LabNumber & ""
280   Else
290       SampleID = "0"
300   End If

310   Count = 0

320   OriginalPrinter = Printer.DeviceName

330   If Not SetFormPrinter() Then Exit Sub

340   For X = 0 To Count
350     Printer.Font.Name = "Courier New"
360     Printer.Font.Size = 14
370     Printer.Font.Bold = True
  
380     Printer.ForeColor = vbRed
390     PrintHeadingCavan SampleID
  
400     Printer.ForeColor = vbBlack
  
410     Printer.Font.Name = "Courier New"
420     Printer.Font.Size = 12
430     Printer.Font.Bold = False
  
440     g.Col = 0
450     For n = 1 To g.Rows - 1
460       g.Row = n
470       If g.CellBackColor = vbRed Then
480         strName = g
490         strTypenex = g.TextMatrix(n, 1)
500         Printer.ForeColor = vbRed
510         Printer.Font.Size = 16
520         Printer.Print
530         Printer.Font.Bold = True
540         Printer.Print "The following units have been reclaimed from this Patient"
550         Printer.Print "at " & Format$(g.TextMatrix(n, 10), "hh:nn") & _
                          " on "; Format$(g.TextMatrix(n, 10), "dd/mm/yyyy") & "."
560         Printer.ForeColor = vbBlack
570         Printer.Print
580         Printer.Font.Size = 12
590         Printer.Print "Unit             Group  Product                               Date Cross Matched"
600         For Y = g.Row To g.Rows - 1
610           g.Row = Y
620           g.Col = 0
630           If g = strName And g.TextMatrix(Y, 1) = strTypenex Then
640             Printer.Print g.TextMatrix(Y, 5); 'Unit
650             Printer.Print Tab(18); g.TextMatrix(Y, 6); 'Group
660             Printer.Print Tab(25); Left$(g.TextMatrix(Y, 7), 37); 'Product
670             If g.TextMatrix(Y, 8) <> "" Then
680               Printer.Print Tab(63); g.TextMatrix(Y, 8);
690             End If
700             Printer.Print
710           End If
720         Next
730         Printer.Print
740         Printer.ForeColor = vbRed
750         Printer.Font.Size = 22
760         Printer.Print "  THE ABOVE UNITS ARE NO LONGER AVAILABLE"
770         Printer.Print "              FOR THIS PATIENT"
  
780          Do While Printer.CurrentY < 6700
790             Printer.Print
800           Loop

810         Printer.ForeColor = vbRed
820         Printer.Font.Size = 4
830         Printer.Print String$(230, "-")

840         Printer.Font.Size = 10
850         Printer.Font.Bold = False
  
860         Printer.Print "Report Date:"; Format(Now, "dd/mm/yyyy hh:nn");
  
870         Printer.Print "    Reported By "; UserName;
880         Printer.EndDoc
890         Exit For
900       End If
910     Next
920   Next

930   For Each Px In Printers
940     If Px.DeviceName = OriginalPrinter Then
950       Set Printer = Px
960       Exit For
970     End If
980   Next

990   Answer = iMsg("Was Print run successful?", vbQuestion + vbYesNo)
1000  If TimedOut Then Unload Me: Exit Sub
1010  If Answer = vbYes Then
1020    sql = "delete from reclaimed where " & _
              "name = '" & AddTicks(strName) & "'"
1030    CnxnBB(0).Execute sql
  
1040    FillG
1050  End If

1060  Exit Sub

PrintReclaimed_Error:

      Dim strES As String
      Dim intEL As Integer

1070  intEL = Erl
1080  strES = Err.Description
1090  LogError "fPrintReclaimed", "PrintReclaimed", intEL, strES, sql

End Sub

Public Function IsProductBloodUnit(ByVal strProd As String) As Boolean
      Dim sql As String
      Dim rsRec As Recordset

10    sql = "Select * from Lists where Listtype = 'B' and text = '" & strProd & "'"
20    Set rsRec = New Recordset
30    RecOpenServerBB 0, rsRec, sql
40    If Not rsRec.EOF Then
50        IsProductBloodUnit = False
60    Else
70        IsProductBloodUnit = True
80    End If

End Function

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bprint_Click()

      Dim n As Integer
      Dim Found As Boolean

10    g.Col = 0
20    For n = 1 To g.Rows - 1
30      g.Row = n
40      If g.CellBackColor = vbRed Then
50        Found = True
60        Exit For
70      End If
80    Next
90    If Not Found Then
100     iMsg "Select Patient", vbInformation
110     If TimedOut Then Unload Me: Exit Sub
120   Else
130     PrintReclaimed
140   End If

End Sub


Private Sub Form_Load()

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************

End Sub

Private Sub g_Click()

      Dim strName As String
      Dim strTypenex As String
      Dim n As Integer
      Dim Highlighting As Boolean

10    If g.MouseRow = 0 Then Exit Sub

20    strName = g.TextMatrix(g.Row, 0)
30    strTypenex = g.TextMatrix(g.Row, 1)

40    g.Col = 0
50    Highlighting = g.CellBackColor <> vbRed
60    For n = 1 To g.Rows - 1
70      g.Row = n
80      If g.TextMatrix(n, 0) = strName And g.TextMatrix(n, 1) = strTypenex Then
90        If Highlighting Then
100         g.CellBackColor = vbRed
110         g.CellForeColor = vbYellow
120       Else
130         g.CellBackColor = &H80000018
140         g.CellForeColor = &H8000000D
150       End If
160     Else
170       g.CellBackColor = &H80000018
180       g.CellForeColor = &H8000000D
190     End If
200   Next

End Sub


