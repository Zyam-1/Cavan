VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form fCurrentStock 
   Caption         =   "Current Stock"
   ClientHeight    =   4020
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   7665
   DrawWidth       =   10
   Icon            =   "fCurrentStock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7665
   Begin VB.CommandButton bLog 
      Caption         =   "&Log Stock as Correct"
      Height          =   525
      Left            =   2880
      TabIndex        =   4
      Top             =   2970
      Width           =   1935
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   6150
      TabIndex        =   2
      Top             =   2970
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   570
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   9
      Cols            =   5
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   "<Blood Group             |^Cell In Stock |^Cross Matched|^FFP In Stock |^Cryo In Stock"
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
      Left            =   1140
      TabIndex        =   5
      Top             =   3840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Products In Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   7125
   End
End
Attribute VB_Name = "fCurrentStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim Generic As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo FillG_Error

20    Ps.LoadLatestBetweenExpiryDates Now, Now + 1825  '5 years
30    For Each p In Ps
40      If InStr("CRXPIE", p.PackEvent) > 0 Then
50        g.col = 0
60        Select Case p.GroupRh
            Case "51": 'O Pos
70            g.row = 1
80          Case "62": 'A Pos
90            g.row = 2
100         Case "73": 'B Pos
110           g.row = 3
120         Case "84": 'AB Pos
130           g.row = 4
140         Case "95": 'O Neg
150           g.row = 5
160         Case "06": 'A Neg
170           g.row = 6
180         Case "17": 'B Neg
190           g.row = 7
200         Case "28": 'AB Neg
210           g.row = 8
220       End Select
230       Generic = ProductGenericFor(p.BarCode)
240       If UCase(Generic) = "PLASMA" Or UCase(Generic) = "LG OCTAPLAS" Then
250         g.col = 3
260       ElseIf Generic = "Cryoprecipitate" Then
270         g.col = 4
280       ElseIf Generic = "Red Cells" Then
290         Select Case p.PackEvent
              Case "C", "R": 'Received into Stock or Restocked
300             g.col = 1
310           Case "X", "P", "I", "V": 'Cross matched, Pending or Issued
320             g.col = 2
330         End Select
340       End If
350       If g.row * g.col <> 0 Then
360         g = Format(Val(g) + 1)
370       End If
380     End If
390   Next

400   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "fCurrentStock", "FillG", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bLog_Click()

      Dim tb As Recordset
      Dim X As Integer
      Dim y As Integer
      Dim sql As String

10    On Error GoTo bLog_Click_Error

20    sql = "select * from ProductInStock"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    tb.AddNew
60    For y = 1 To 8
70      g.row = y
80      For X = 1 To 4
90        g.col = X
100       tb(Format(y) & Format(X)) = Val(g)
110     Next
120   Next
130   tb!DateTime = Format(Now, "dd/mmm/yyyy")
140   tb!Operator = UserCode
150   tb.Update

160   Exit Sub

bLog_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fCurrentStock", "bLog_Click", intEL, strES, sql


End Sub

Private Sub bprint_Click()

      Dim y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub
30    Printer.Orientation = vbPRORPortrait

40    Printer.Font.Name = "Courier New"
50    Printer.Font.Size = 10
60    Printer.Font.Bold = True
70    Printer.ForeColor = vbRed

80    Printer.Print FormatString("CAVAN GENERAL HOSPITAL : Current Stock Report", 108, , AlignCenter)
90    Printer.Print FormatString("Blood Transfusion Laboratory", 108, , AlignCenter)
      'Printer.Print "Phone 38833"

100   Printer.Font.Size = 9
110   Printer.ForeColor = vbBlack
      '****Report heading
120   For i = 1 To 108
130       Printer.Print "-";
140   Next i
150   Printer.Print

160   Printer.Print FormatString("", 0, "|");
170   Printer.Print FormatString("Blood Group", 34, "|", AlignCenter);
180   Printer.Print FormatString("Cells in Stock", 17, "|", AlignCenter);
190   Printer.Print FormatString("X Matched Stock", 17, "|", AlignCenter);
200   Printer.Print FormatString("FFP in Stock", 17, "|", AlignCenter);
210   Printer.Print FormatString("Cryo in Stock", 17, "|", AlignCenter)
220   Printer.Font.Bold = False
230   For i = 1 To 108
240       Printer.Print "-";
250   Next i
260   Printer.Print
270   For y = 1 To g.Rows - 1
280       Printer.Print FormatString("", 0, "|");
290       Printer.Print FormatString(g.TextMatrix(y, 0), 34, "|");
300       Printer.Print FormatString(g.TextMatrix(y, 1), 17, "|", AlignCenter);
310       Printer.Print FormatString(g.TextMatrix(y, 2), 17, "|", AlignCenter);
320       Printer.Print FormatString(g.TextMatrix(y, 3), 17, "|", AlignCenter);
330       Printer.Print FormatString(g.TextMatrix(y, 4), 17, "|", AlignCenter)
 
340   Next

350   Printer.Font.Italic = True
360   Printer.Print
370   Printer.Print "Report Date:"; Format(Now, "dd/mm/yyyy");
380   Printer.Print "    Reported By "; UserName;
390   Printer.Font.Italic = False

400   Printer.EndDoc

410   For Each Px In Printers
420     If Px.DeviceName = OriginalPrinter Then
430       Set Printer = Px
440       Exit For
450     End If
460   Next

End Sub


Private Sub Form_Load()

      Dim n As Integer

10    g.col = 0
20    For n = 1 To 8
30      g.row = n
40      g = Choose(n, "O Rh(D) Positive", "A Rh(D) Positive", _
                      "B Rh(D) Positive", "AB Rh(D) Positive", _
                      "O Rh(D) Negative", "A Rh(D) Negative", _
                      "B Rh(D) Negative", "AB Rh(D) Negative")
50    Next

60    FillG

End Sub


