VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBatchHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batch History"
   ClientHeight    =   7635
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13875
   Icon            =   "frmBatchHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintLabel 
      Appearance      =   0  'Flat
      Caption         =   "&Print Label"
      Height          =   900
      Left            =   7950
      Picture         =   "frmBatchHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "bprintlabels"
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   900
      Left            =   12780
      Picture         =   "frmBatchHistory.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   900
      Left            =   10380
      Picture         =   "frmBatchHistory.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   900
      Left            =   4680
      Picture         =   "frmBatchHistory.frx":1548
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton bPrintBoth 
      Caption         =   "Print &Both"
      Height          =   900
      Left            =   6960
      Picture         =   "frmBatchHistory.frx":19D3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "Print History"
      Height          =   900
      Left            =   11580
      Picture         =   "frmBatchHistory.frx":1CDD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton bPrintSpecific 
      Caption         =   "Print &Form"
      Height          =   900
      Left            =   5970
      Picture         =   "frmBatchHistory.frx":2347
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox tBatchNumber 
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   75
      Width           =   1905
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   1890
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmBatchHistory.frx":29B1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   30
      TabIndex        =   10
      Top             =   7380
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   975
      TabIndex        =   22
      Top             =   1155
      Width           =   435
   End
   Begin VB.Label lblGroup 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1470
      TabIndex        =   21
      ToolTipText     =   "Batch Group"
      Top             =   1110
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume/Dose"
      Height          =   195
      Left            =   105
      TabIndex        =   20
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label lblDose 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   19
      ToolTipText     =   "Volume/Dose"
      Top             =   1455
      Width           =   2520
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ml or IU"
      Height          =   195
      Left            =   4065
      TabIndex        =   18
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Current Stock"
      Height          =   195
      Left            =   2190
      TabIndex        =   17
      Top             =   1155
      Width           =   975
   End
   Begin VB.Label lblCurrentStock 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3270
      TabIndex        =   16
      Top             =   1110
      Width           =   1305
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1470
      TabIndex        =   15
      Top             =   765
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   990
      TabIndex        =   14
      Top             =   810
      Width           =   420
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1470
      TabIndex        =   13
      Top             =   420
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   855
      TabIndex        =   12
      Top             =   465
      Width           =   555
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8940
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmBatchHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    If Trim$(tBatchNumber) = "" Then Exit Sub

60    sql = "select * from BatchDetails where " & _
            "BatchNumber = '" & tBatchNumber & "' " & _
            "Order by date desc"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql

90    If tb.EOF Then
100     iMsg "Batch Number not found!", vbExclamation
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   Fill_BatchList

150   Do While Not tb.EOF
160     If lblProduct = "" Then
170       lblProduct = tb!Product & ""
180     End If
190     If lblExpiry = "" Then
200       If Not IsNull(tb!Expiry) Then
210         lblExpiry = Format(tb!Expiry, "dd/mm/yyyy")
220       End If
230     End If
  
240     s = tb!SampleID & vbTab & _
            Format(tb!Date, "dd/mm/yy hh:mm:ss") & vbTab & _
            tb!Name & vbTab & _
            gEVENTCODES(tb!Event & "").Text & vbTab & _
            tb!Chart & vbTab & _
            tb!Addr0 & vbTab & _
            "" & _
            Format(tb!Bottles) & vbTab
250     If Not IsNull(tb!EventStart) Then
260       s = s & Format(tb!EventStart, "dd/mm/yyyy HH:nn:ss")
270     End If
280     s = s & vbTab
290     If Not IsNull(tb!EventEnd) Then
300       s = s & Format(tb!EventEnd, "dd/mm/yyyy HH:nn:ss")
310     End If
320     s = s & vbTab & _
            tb!UserCode & vbTab & _
            tb!Comment & "" & vbTab & _
            Format(tb!Date, "dd/mm/yy hh:mm:ss")
330     g.AddItem s
340     tb.MoveNext
350   Loop

360   If g.Rows > 2 Then
370     g.RemoveItem 1
380   End If

390   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "frmBatchHistory", "FillG", intEL, strES, sql


End Sub

Private Sub Fill_BatchList()

      Dim sql As String
      Dim sn As Recordset

10    On Error GoTo Fill_BatchList_Error

20    sql = "select * from batchproductlist where " & _
            "batchnumber = '" & Trim$(tBatchNumber) & "'"

30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql
50    If sn.EOF Then
60      lblProduct = "No Such Batch"
70    Else
80      lblGroup = Trim$(sn!Group & "")
90      lblExpiry = (sn!DateExpiry & "")
100     lblProduct = sn!Product
110     lblDose = Trim$(sn!UnitVolume & "")
120     lblCurrentStock = sn!CurrentStock & ""
130     If lblProduct = "Anti-D" Then
140      lblDose = lblDose & " IU"
150     End If
160   End If

170   Exit Sub

Fill_BatchList_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmBatchHistory", "Fill_BatchList", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bprint_Click()

      Dim Y As Integer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Font.Size = 9
50    Printer.Orientation = vbPRORLandscape
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                                             History for Batch Number ";
90    Printer.Print tBatchNumber

      'Heading Section

100   Printer.Font.Bold = True
110   For i = 1 To 152
120       Printer.Print "-";
130   Next i

140   Printer.Print

150   Printer.Print FormatString("Lab #", 10, "|"); 'Lab Number
160   Printer.Print FormatString("Date/Time", 18, "|"); 'Date/Time
170   Printer.Print FormatString("Name", 20, "|"); 'Name
180   Printer.Print FormatString("Details", 25, "|"); 'Details
190   Printer.Print FormatString("Chart #", 10, "|"); 'Chart Number
200   Printer.Print FormatString("Address", 25, "|"); 'Address
210   Printer.Print FormatString("Units", 6, "|"); 'Units
220   Printer.Print FormatString("Op", 8, "|"); 'Operator
230   Printer.Print FormatString("Comments", 21, "|") 'Comments

240   For i = 1 To 152
250       Printer.Print "-";
260   Next i
270   Printer.Print
280   For Y = 1 To g.Rows - 1
    
              'Detail section
290           Printer.Font.Bold = False
300           Printer.Print FormatString(g.TextMatrix(Y, 0), 10, "|"); 'Lab Number
310           Printer.Print FormatString(g.TextMatrix(Y, 1), 18, "|"); 'Date/Time
320           Printer.Print FormatString(g.TextMatrix(Y, 2), 20, "|"); 'Name
330           Printer.Print FormatString(g.TextMatrix(Y, 3), 25, "|"); 'Details
340           Printer.Print FormatString(g.TextMatrix(Y, 4), 10, "|"); 'Chart Number
350           Printer.Print FormatString(g.TextMatrix(Y, 5), 25, "|"); 'Address
360           Printer.Print FormatString(g.TextMatrix(Y, 7), 6, "|"); 'Units
370           Printer.Print FormatString(g.TextMatrix(Y, 9), 8, "|"); 'Operator
380           Printer.Print FormatString(g.TextMatrix(Y, 10), 21, "|") 'Comments
390   Next


400   Printer.EndDoc

End Sub

Private Sub bPrintBoth_Click()

      Dim n As Integer
      Dim Printed As Boolean
      Dim TwoForms As Boolean
      Dim tb As Recordset
      Dim sql As String

10    Printed = False
20    TwoForms = False
30    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
40    If TimedOut Then Unload Me: Exit Sub
50    If Answer = vbYes Then
60      TwoForms = True
70    End If

80    g.col = 0
90    For n = 1 To g.Rows - 1
100     sql = "Select SampleDate,DateReceived From PatientDetails Where labnumber = '" & g.TextMatrix(n, 0) & "'"
110     Set tb = New Recordset
120     RecOpenClientBB 0, tb, sql
130     If tb.EOF Then
140       CurrentReceivedDate = ""
150     Else
160       CurrentReceivedDate = tb!DateReceived & ""
170     End If
180     g.row = n
190     If g.CellBackColor = vbRed Then
200           PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
210           If TwoForms Then
220               PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
230           End If
240       PrintBatchLabels g.TextMatrix(n, 0)
250       Printed = True
260       Exit For
270     End If
280   Next

290   If Not Printed Then
300     iMsg "Select Lab Number to print!", vbInformation
310     If TimedOut Then Unload Me: Exit Sub
320   End If

End Sub

Private Sub cmdPrintLabel_Click()

      Dim n As Integer
      Dim Printed As Boolean

10    Printed = False

20    g.col = 0
30    For n = 1 To g.Rows - 1
40      g.row = n
50      If g.CellBackColor = vbRed Then
60        PrintBatchLabels g.TextMatrix(n, 0)
70        Printed = True
80        Exit For
90      End If
100   Next

110   If Not Printed Then
120     iMsg "Select Lab Number to print!", vbInformation
130     If TimedOut Then Unload Me: Exit Sub
140   End If

End Sub

Private Sub bPrintSpecific_Click()

      Dim n As Integer
      Dim Printed As Boolean
      Dim TwoForms As Boolean
      Dim tb As Recordset
      Dim sql As String

10    Printed = False
20    TwoForms = False
30    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
40    If TimedOut Then Unload Me: Exit Sub
50    If Answer = vbYes Then
60      TwoForms = True
70    End If

80    g.col = 0
90    For n = 1 To g.Rows - 1
100     sql = "Select SampleDate,DateReceived From PatientDetails Where labnumber = '" & g.TextMatrix(n, 0) & "'"
110     Set tb = New Recordset
120     RecOpenClientBB 0, tb, sql
130     If tb.EOF Then
140       CurrentReceivedDate = ""
150     Else
160       CurrentReceivedDate = tb!DateReceived & ""
170     End If
  
180     g.row = n
190     If g.CellBackColor = vbRed Then
200           PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
210           If TwoForms Then
220               PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
230           End If
240       Printed = True
250       Exit For
260     End If
270   Next

280   If Not Printed Then
290     iMsg "Select Lab Number to print!", vbInformation
300     If TimedOut Then Unload Me: Exit Sub
310   End If

End Sub

Private Sub cmdSearch_Click()

10    lblProduct = ""
20    lblExpiry = ""

30    FillG

End Sub


Private Sub cmdXL_Click()

10    g.Cols = g.Cols - 1
20    ExportFlexGrid g, Me
30    g.FormatString = "<Lab Number     |<Date/Time          |<Name                             |<Details       |<Chart          |<Address                    |^Units |<Start Date/Time       |<End Date/Time       |<Op       |<Comment                |"
40    g.ColWidth(11) = 0
50    FillG
End Sub



Private Sub Form_Load()


10    g.ColWidth(11) = 0
      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
20        FillG
      '**************************************

End Sub





Private Sub g_Click()

      Dim n As Integer
      Dim ySave As Integer

10    If g.MouseRow = 0 And g.col <> 7 And g.col <> 8 Then
  
20      If InStr(g.TextMatrix(0, g.col), "Date") = 0 Then
30        If SortOrder Then
40          g.Sort = flexSortGenericAscending
50        Else
60          g.Sort = flexSortGenericDescending
70        End If
80      Else
90        g.Sort = 9
100     End If
110     SortOrder = Not SortOrder
120     Exit Sub
130   End If

140   ySave = g.row

150   g.col = 0
160   For n = 1 To g.Rows - 1
170     g.row = n
180     If g.CellBackColor = vbRed Then
190       g.CellBackColor = 0
200       Exit For
210     End If
220   Next
230   g.row = ySave
240   g.CellBackColor = vbRed

End Sub


Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String
      Dim GC As Integer

10    GC = g.col

20    If Not IsDate(g.TextMatrix(Row1, GC)) Then
30      Cmp = 0
40      Exit Sub
50    End If

60    If Not IsDate(g.TextMatrix(Row2, GC)) Then
70      Cmp = 0
80      Exit Sub
90    End If

100   d1 = Format(g.TextMatrix(Row1, GC), "dd/MMM/yyyy HH:mm:ss")
110   d2 = Format(g.TextMatrix(Row2, GC), "dd/MMM/yyyy HH:mm:ss")

120   If SortOrder Then
130     Cmp = Sgn(DateDiff("D", d1, d2))
140   Else
150     Cmp = Sgn(DateDiff("D", d2, d1))
160   End If

End Sub


Private Sub g_DblClick()

      Dim Y As Integer
      Dim max As Long

10    max = 0

20    For Y = 0 To g.Rows - 1
30      If TextWidth(g.TextMatrix(Y, g.col)) > max Then
40        max = TextWidth(g.TextMatrix(Y, g.col)) + TextWidth("W")
50      End If
60    Next
70    g.ColWidth(g.col) = max

End Sub

Private Sub tBatchNumber_LostFocus()

10    FillG

End Sub


