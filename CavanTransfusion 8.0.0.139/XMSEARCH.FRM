VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form xmatchedsearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for Cross Matches"
   ClientHeight    =   8865
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   14460
   ControlBox      =   0   'False
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
   Icon            =   "XMSEARCH.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8865
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdUnit 
      Height          =   7635
      Left            =   105
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   13467
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"XMSEARCH.frx":08CA
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6210
      Picture         =   "XMSEARCH.frx":09A5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   1830
      TabIndex        =   5
      Top             =   390
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   191692801
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   390
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   191692801
      CurrentDate     =   36963
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   7635
      Left            =   90
      TabIndex        =   3
      Top             =   930
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   13467
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"XMSEARCH.frx":0CAF
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "Start Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4890
      Picture         =   "XMSEARCH.frx":0D86
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton btnprint 
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
      Height          =   705
      Left            =   8610
      Picture         =   "XMSEARCH.frx":11C8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton btncancel 
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
      Height          =   705
      Left            =   12960
      Picture         =   "XMSEARCH.frx":1832
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   8550
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   720
      Picture         =   "XMSEARCH.frx":1E9C
      Top             =   540
      Width           =   360
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Unit Number to show Unit History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   11
      Top             =   90
      Width           =   1500
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7200
      TabIndex        =   8
      Top             =   210
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Between Dates"
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
      Left            =   1830
      TabIndex        =   6
      Top             =   120
      Width           =   2925
   End
End
Attribute VB_Name = "xmatchedsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub btnCancel_Click()

10    Unload Me

End Sub

Private Sub btnprint_Click()
  
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Orientation = vbPRORLandscape

40    Printer.FontName = "Courier New"
50    Printer.FontSize = 9

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print "                                                          XMatched Products"
80    Printer.Print
90    Printer.Print "Search Results for xmatched products From "; dtFrom.Value; " To "; dtTo.Value
      Dim i As Integer
100   For i = 1 To 152
110       Printer.Print "_";
120   Next i
      '****Report body

130   Printer.Print
140   Printer.Print FormatString("Unit", 16, "|");
150   Printer.Print FormatString("Exp.Date", 14, "|");
160   Printer.Print FormatString("Product", 40, "|");
170   Printer.Print FormatString("Group", 8, "|");
180   Printer.Print FormatString("Pat.ID", 10, "|");
190   Printer.Print FormatString("A and E", 10, "|");
200   Printer.Print FormatString("Name", 30, "|");
210   Printer.Print FormatString("Op", 5, "|");
220   Printer.Print FormatString("Date", 16, "|")
230   Printer.Font.Bold = False
240   For i = 1 To 152
250     Printer.Print "-";
260     If i = 152 Then
270         Printer.Print
280     End If
290   Next i

300   With grdMain
310   For Y = 1 To .Rows - 1
320       Printer.Print FormatString(.TextMatrix(Y, 0), 16, "|");
330       Printer.Print FormatString(.TextMatrix(Y, 1), 14, "|");
340       Printer.Print FormatString(.TextMatrix(Y, 2), 40, "|");
350       Printer.Print FormatString(.TextMatrix(Y, 3), 8, "|");
360       Printer.Print FormatString(.TextMatrix(Y, 4), 10, "|");
370       Printer.Print FormatString(.TextMatrix(Y, 5), 10, "|");
380       Printer.Print FormatString(.TextMatrix(Y, 6), 30, "|");
390       Printer.Print FormatString(.TextMatrix(Y, 7), 5, "|");
400       Printer.Print FormatString(.TextMatrix(Y, 8), 16, "|")
 
410   Next
420   End With

430   Printer.EndDoc
440   Printer.Orientation = vbPRORLandscape
450   For Each Px In Printers
460     If Px.DeviceName = OriginalPrinter Then
470       Set Printer = Px
480       Exit For
490     End If
500   Next

End Sub

Private Sub cmdSearch_Click()

      Dim sn As Recordset
      Dim sql As String
      Dim s As String
      Dim fdate As String
      Dim tdate As String

10    On Error GoTo cmdSearch_Click_Error
20    grdUnit.Visible = False

30    cmdXL.Visible = True
40    btnprint.Visible = True

50    With grdMain
60      .Visible = True
70      .Rows = 2
80      .AddItem ""
90      .RemoveItem 1
100   End With

110   fdate = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
120   tdate = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

130   sql = "SELECT L.Wording, P.ISBT128, P.DateExpiry, LS.Text GroupRh, " & _
            "P.PatID, P.PatName, P.Operator, P.[DateTime], PD.AandE, PD.SampleDate " & _
            "FROM Product P JOIN ProductList L ON L.BarCode = P.BarCode " & _
            "JOIN PatientDetails PD ON P.LabNumber = PD.LabNumber " & _
            "JOIN Lists Ls ON Ls.Code = P.GroupRH and Ls.ListType = 'GroupBarCodes' " & _
            "WHERE " & _
            "P.[DateTime] BETWEEN '" & fdate & "' AND '" & tdate & "' " & _
            "AND P.Event = 'X' " & _
            "AND L.BarCode = P.BarCode " & _
            "AND P.ISBT128 IN ( " & _
            "   SELECT DISTINCT ISBT128 FROM Product WHERE " & _
            "   [DateTime] BETWEEN '" & fdate & "' AND '" & tdate & "' " & _
            "   AND Event = 'X' ) " & _
            "ORDER BY PD.SampleDate desc"
140   Set sn = New Recordset
150   RecOpenServerBB 0, sn, sql
160   If sn.EOF Then
170     iMsg "No Records found.", vbInformation
180     If TimedOut Then Unload Me: Exit Sub
190     Exit Sub
200   End If

210   grdMain.Visible = False

220   Do While Not sn.EOF
230     s = sn!ISBT128 & vbTab & _
            Format(sn!DateExpiry, "dd/mm/yyyy HH:ss") & vbTab & _
            sn!Wording & vbTab & _
            sn!GroupRh & vbTab & _
            sn!Patid & vbTab & _
            sn!AandE & vbTab & _
            sn!PatName & vbTab & _
            sn!Operator & vbTab & _
            Format$(sn!DateTime, "dd/MM/yy HH:nn") & vbTab & _
            Format$(sn!SampleDate, "dd/mm/yy HH:nn")
240     grdMain.AddItem s
250     sn.MoveNext
260   Loop

270   With grdMain
280   If .Rows > 2 Then
290     .RemoveItem 1
300   End If
'
'260   .Col = 7
'270   .Sort = 9
310   .Visible = True
320   End With

330   Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "xmatchedsearch", "cmdSearch_Click", intEL, strES, sql
370   grdMain.Visible = True

End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String

10    strHeading = "XMatched Products" & vbCr
20    strHeading = strHeading & "Search Results for xmatched products From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid grdMain, Me, strHeading

End Sub

Private Sub Form_Load()

10    dtFrom = Format(Now - 7, "dd/mmm/yyyy")
20    dtTo = Format(Now, "dd/mmm/yyyy")

End Sub

Private Sub grdMain_Click()

      Dim sn As Recordset
      Dim sql As String
      Dim s As String
      Dim Criteria As String

10    On Error GoTo GrdMain_Click_Error

20    With grdMain
30    If .MouseRow = 0 Then
40      If SortOrder Then
50        If .Col = 1 Or .Col = 8 Then
60          .Sort = 9
70        Else
80          .Sort = flexSortGenericAscending
90        End If
100     Else
110       If .Col = 1 Or .Col = 8 Then
120         .Sort = 9
130       Else
140         .Sort = flexSortGenericDescending
150       End If
160     End If
170     SortOrder = Not SortOrder
180     Exit Sub
190   End If

200   Criteria = .TextMatrix(.Row, 0)
210   If Trim$(Criteria) = "" Then Exit Sub
220   .Visible = False

230   End With

240   cmdXL.Visible = False
250   btnprint.Visible = False

260   With grdUnit
270   .Visible = True
280   .Rows = 2
290   .AddItem ""
300   .RemoveItem 1

310   sql = "SELECT * FROM Product " & _
            "WHERE ISBT128 = '" & Criteria & "' " & _
            "ORDER BY Counter desc"
320   Set sn = New Recordset
330   RecOpenServerBB 0, sn, sql

340   Do While Not sn.EOF
350     s = sn!ISBT128 & "" & vbTab & _
            sn!DateExpiry & vbTab & _
            ProductWordingFor(sn!BarCode & "") & vbTab & _
            Bar2Group(sn!GroupRh & "") & vbTab & _
            vbTab & sn!Patid & vbTab & _
            sn!PatName & vbTab & _
            gEVENTCODES(sn!Event & "").Text & vbTab & _
            sn!Operator & vbTab & _
            Format$(sn!DateTime, "dd/MM/yy HH:nn")

360     .AddItem s, 1
370     sn.MoveNext
380   Loop
390   End With

400   Exit Sub

GrdMain_Click_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "xmatchedsearch", "Grid1_Click", intEL, strES, sql

End Sub

Private Sub GrdMain_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    With grdMain
20    If Not IsDate(.TextMatrix(Row1, .Col)) Then
30      Cmp = 0
40      Exit Sub
50    End If

60    If Not IsDate(.TextMatrix(Row2, .Col)) Then
70      Cmp = 0
80      Exit Sub
90    End If

100   d1 = Format(.TextMatrix(Row1, .Col), "dd/MM/yy HH:nn")
110   d2 = Format(.TextMatrix(Row2, .Col), "dd/MM/yy HH:nn")

120   If SortOrder Then
130     Cmp = Sgn(DateDiff("s", d1, d2))
140   Else
150     Cmp = -Sgn(DateDiff("s", d1, d2))
160   End If
170   End With

End Sub


