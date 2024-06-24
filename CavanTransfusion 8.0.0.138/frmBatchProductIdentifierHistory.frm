VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBatchProductIdentifierHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Batch History"
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14700
   Icon            =   "frmBatchProductIdentifierHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFating 
      Caption         =   "Fate"
      Height          =   1005
      Left            =   10530
      Picture         =   "frmBatchProductIdentifierHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton cmdPrintLabel 
      Appearance      =   0  'Flat
      Caption         =   "&Print Label"
      Height          =   1005
      Left            =   3810
      Picture         =   "frmBatchProductIdentifierHistory.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "bprintlabels"
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1005
      Left            =   13680
      Picture         =   "frmBatchProductIdentifierHistory.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1005
      Left            =   6180
      Picture         =   "frmBatchProductIdentifierHistory.frx":2368
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton bPrintBoth 
      Caption         =   "Print &Both"
      Height          =   1005
      Left            =   2850
      Picture         =   "frmBatchProductIdentifierHistory.frx":7F7A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "Print History"
      Height          =   1005
      Left            =   120
      Picture         =   "frmBatchProductIdentifierHistory.frx":8284
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7350
      Width           =   900
   End
   Begin VB.CommandButton bPrintSpecific 
      Caption         =   "Print &Form"
      Height          =   1005
      Left            =   1890
      Picture         =   "frmBatchProductIdentifierHistory.frx":88EE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7350
      Width           =   900
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6045
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   11
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
      FormatString    =   $"frmBatchProductIdentifierHistory.frx":8F58
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   8370
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Identifier"
      Height          =   195
      Left            =   3120
      TabIndex        =   23
      Top             =   150
      Width           =   600
   End
   Begin VB.Label lblBatchNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   510
      TabIndex        =   22
      Top             =   360
      Width           =   2070
   End
   Begin VB.Label lblIdentifier 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2670
      TabIndex        =   21
      Top             =   360
      Width           =   2070
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   7125
      TabIndex        =   20
      Top             =   435
      Width           =   435
   End
   Begin VB.Label lblGroup 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7620
      TabIndex        =   19
      ToolTipText     =   "Batch Group"
      Top             =   390
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume/Dose"
      Height          =   195
      Left            =   6195
      TabIndex        =   18
      Top             =   780
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
      Left            =   7590
      TabIndex        =   17
      ToolTipText     =   "Volume/Dose"
      Top             =   750
      Width           =   2520
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ml or IU"
      Height          =   195
      Left            =   10155
      TabIndex        =   16
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Current Stock"
      Height          =   195
      Left            =   11520
      TabIndex        =   15
      Top             =   360
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
      Left            =   12570
      TabIndex        =   14
      Top             =   300
      Width           =   1005
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9330
      TabIndex        =   13
      Top             =   405
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   8850
      TabIndex        =   12
      Top             =   450
      Width           =   420
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7650
      TabIndex        =   11
      Top             =   30
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   7035
      TabIndex        =   10
      Top             =   75
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
      Left            =   7110
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   150
      Width           =   1020
   End
End
Attribute VB_Name = "frmBatchProductIdentifierHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG()

      Dim s As String
      Dim BP As BatchProduct
      Dim BPs As New BatchProducts

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    If Trim$(lblIdentifier) = "" Then Exit Sub

60    BPs.LoadSpecificIdentifier lblIdentifier

70    If BPs.Count = 0 Then
80      iMsg "Identifier not found!", vbExclamation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   FillCommonDetails BPs.Item(1)

130   For Each BP In BPs
  
140     s = BP.SampleID & vbTab & _
            gEVENTCODES(BP.EventCode).Text & vbTab & _
            Format(BP.RecordDateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            BP.PatName & vbTab & _
            BP.Chart & vbTab & _
            BP.Addr0 & vbTab
150     If Format$(BP.EventStart, "dd/MM/yyyy") <> "01/01/1900" Then
160       s = s & Format(BP.EventStart, "dd/MM/yyyy HH:nn")
170     End If
180     s = s & vbTab
190     If Format$(BP.EventEnd, "dd/MM/yyyy") <> "01/01/1900" Then
200       s = s & Format(BP.EventEnd, "dd/MM/yyyy HH:nn")
210     End If
220     s = s & vbTab & _
            BP.UserName & vbTab & _
            BP.Comment
230     g.AddItem s
240   Next

250   If g.Rows > 2 Then
260     g.RemoveItem 1
270   End If

280   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmBatchProductIdentifierHistory", "FillG", intEL, strES

End Sub

Private Sub FillCommonDetails(ByVal BP As BatchProduct)

Dim BPs As New BatchProducts

10    On Error GoTo FillCommonDetails_Error

20      lblGroup = BP.UnitGroup
30      lblExpiry = BP.DateExpiry
40      lblProduct = BP.Product
50      lblDose = BP.UnitVolume
60      lblCurrentStock = BPs.CountProductBatchInStock(BP.Product, lblBatchNumber)
70      If lblProduct = "Anti-D" Then
80       lblDose = lblDose & " IU"
90      End If

100   Exit Sub

FillCommonDetails_Error:

Dim strES As String
Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmBatchProductIdentifierHistory", "FillCommonDetails", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bprint_Click()

      Dim Y As Integer
      Dim OriginalPrinter As String
      Dim i As Integer

10    On Error GoTo bprint_Click_Error

20    OriginalPrinter = Printer.DeviceName

30    If Not SetFormPrinter() Then Exit Sub

40    Printer.FontName = "Courier New"
50    Printer.Font.Size = 9
60    Printer.Orientation = vbPRORLandscape
70    Printer.Font.Bold = True
80    Printer.Print
90    Printer.Print "                                                             History for Batch Number ";
100   Printer.Print lblBatchNumber

      'Heading Section

110   Printer.Font.Bold = True
120   For i = 1 To 152
130       Printer.Print "-";
140   Next i

150   Printer.Print

160   Printer.Print FormatString("Lab #", 10, "|"); 'Lab Number
170   Printer.Print FormatString("Date/Time", 18, "|"); 'Date/Time
180   Printer.Print FormatString("Name", 20, "|"); 'Name
190   Printer.Print FormatString("Details", 25, "|"); 'Details
200   Printer.Print FormatString("Chart #", 10, "|"); 'Chart Number
210   Printer.Print FormatString("Address", 25, "|"); 'Address
220   Printer.Print FormatString("Units", 6, "|"); 'Units
230   Printer.Print FormatString("Op", 8, "|"); 'Operator
240   Printer.Print FormatString("Comments", 21, "|") 'Comments

250   For i = 1 To 152
260       Printer.Print "-";
270   Next i
280   Printer.Print
290   For Y = 1 To g.Rows - 1
    
              'Detail section
300           Printer.Font.Bold = False
310           Printer.Print FormatString(g.TextMatrix(Y, 0), 10, "|"); 'Lab Number
320           Printer.Print FormatString(g.TextMatrix(Y, 1), 18, "|"); 'Date/Time
330           Printer.Print FormatString(g.TextMatrix(Y, 2), 20, "|"); 'Name
340           Printer.Print FormatString(g.TextMatrix(Y, 3), 25, "|"); 'Details
350           Printer.Print FormatString(g.TextMatrix(Y, 4), 10, "|"); 'Chart Number
360           Printer.Print FormatString(g.TextMatrix(Y, 5), 25, "|"); 'Address
370           Printer.Print FormatString(g.TextMatrix(Y, 7), 6, "|"); 'Units
380           Printer.Print FormatString(g.TextMatrix(Y, 9), 8, "|"); 'Operator
390           Printer.Print FormatString(g.TextMatrix(Y, 10), 21, "|") 'Comments
400   Next


410   Printer.EndDoc

420   Exit Sub

bprint_Click_Error:

Dim strES As String
Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmBatchProductIdentifierHistory", "bprint_Click", intEL, strES

End Sub

Private Sub bPrintBoth_Click()

      Dim n As Integer
      Dim Printed As Boolean
      Dim TwoForms As Boolean
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo bPrintBoth_Click_Error

20    Printed = False
30    TwoForms = False
40    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
50    If TimedOut Then Unload Me: Exit Sub
60    If Answer = vbYes Then
70      TwoForms = True
80    End If

90    g.col = 0
100   For n = 1 To g.Rows - 1
110     sql = "Select SampleDate,DateReceived From PatientDetails Where labnumber = '" & g.TextMatrix(n, 0) & "'"
120     Set tb = New Recordset
130     RecOpenClientBB 0, tb, sql
140     If tb.EOF Then
150       CurrentReceivedDate = ""
160     Else
170       CurrentReceivedDate = tb!DateReceived & ""
180     End If
190     g.row = n
200     If g.CellBackColor = vbRed Then
210           PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
220           If TwoForms Then
230               PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
240           End If
250       PrintBatchLabels g.TextMatrix(n, 0)
260       Printed = True
270       Exit For
280     End If
290   Next

300   If Not Printed Then
310     iMsg "Select Lab Number to print!", vbInformation
320     If TimedOut Then Unload Me: Exit Sub
330   End If

340   Exit Sub

bPrintBoth_Click_Error:

Dim strES As String
Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmBatchProductIdentifierHistory", "bPrintBoth_Click", intEL, strES, sql

End Sub

Private Sub cmdFating_Click()

10    With frmBatchProductMovement
20      .Identifier = lblIdentifier
30      .Show 1
40    End With

End Sub

Private Sub cmdPrintLabel_Click()

      Dim n As Integer
      Dim Printed As Boolean

10    On Error GoTo cmdPrintLabel_Click_Error

20    Printed = False

30    g.col = 0
40    For n = 1 To g.Rows - 1
50      g.row = n
60      If g.CellBackColor = vbRed Then
70        PrintBatchLabels g.TextMatrix(n, 0)
80        Printed = True
90        Exit For
100     End If
110   Next

120   If Not Printed Then
130     iMsg "Select Lab Number to print!", vbInformation
140     If TimedOut Then Unload Me: Exit Sub
150   End If

160   Exit Sub

cmdPrintLabel_Click_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmBatchProductIdentifierHistory", "cmdPrintLabel_Click", intEL, strES

End Sub

Private Sub bPrintSpecific_Click()

      Dim n As Integer
      Dim Printed As Boolean
      Dim TwoForms As Boolean
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo bPrintSpecific_Click_Error

20    Printed = False
30    TwoForms = False
40    Answer = iMsg("Do you want to print Two Forms?", vbQuestion + vbYesNo)
50    If TimedOut Then Unload Me: Exit Sub
60    If Answer = vbYes Then
70      TwoForms = True
80    End If

90    g.col = 0
100   For n = 1 To g.Rows - 1
110     sql = "Select SampleDate,DateReceived From PatientDetails Where labnumber = '" & g.TextMatrix(n, 0) & "'"
120     Set tb = New Recordset
130     RecOpenClientBB 0, tb, sql
140     If tb.EOF Then
150       CurrentReceivedDate = ""
160     Else
170       CurrentReceivedDate = tb!DateReceived & ""
180     End If
  
190     g.row = n
200     If g.CellBackColor = vbRed Then
210           PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
220           If TwoForms Then
230               PrintBatchForm g.TextMatrix(n, 0), g.TextMatrix(n, 11)
240           End If
250       Printed = True
260       Exit For
270     End If
280   Next

290   If Not Printed Then
300     iMsg "Select Lab Number to print!", vbInformation
310     If TimedOut Then Unload Me: Exit Sub
320   End If

330   Exit Sub

bPrintSpecific_Click_Error:

Dim strES As String
Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "frmBatchProductIdentifierHistory", "bPrintSpecific_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

10    g.Cols = g.Cols - 1
20    ExportFlexGrid g, Me
30    g.FormatString = "<Lab Number     |<Date/Time          |<Name                             |<Details       |<Chart          |<Address                    |^Units |<Start Date/Time       |<End Date/Time       |<Op       |<Comment                |"
40    g.ColWidth(11) = 0
50    FillG
End Sub



Private Sub Form_Activate()

10        FillG

End Sub

Private Sub g_Click()

      Dim n As Integer
      Dim ySave As Integer

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 And g.col <> 7 And g.col <> 8 Then
  
30      If InStr(g.TextMatrix(0, g.col), "Date") = 0 Then
40        If SortOrder Then
50          g.Sort = flexSortGenericAscending
60        Else
70          g.Sort = flexSortGenericDescending
80        End If
90      Else
100       g.Sort = 9
110     End If
120     SortOrder = Not SortOrder
130     Exit Sub
140   End If

150   ySave = g.row

160   g.col = 0
170   For n = 1 To g.Rows - 1
180     g.row = n
190     If g.CellBackColor = vbRed Then
200       g.CellBackColor = 0
210       Exit For
220     End If
230   Next
240   g.row = ySave
250   g.CellBackColor = vbRed

260   Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmBatchProductIdentifierHistory", "g_Click", intEL, strES

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

Public Property Let BatchNumber(ByVal sNewValue As String)

10    lblBatchNumber = sNewValue

End Property

Public Property Let Identifier(ByVal sNewValue As String)

10    lblIdentifier = sNewValue

End Property


