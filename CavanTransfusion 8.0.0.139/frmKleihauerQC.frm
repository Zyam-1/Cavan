VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmKleihauerQC 
   Caption         =   "NetAcquire - Kleihauer QC"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmKleihauerQC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   8700
      Picture         =   "frmKleihauerQC.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   705
      Left            =   8700
      Picture         =   "frmKleihauerQC.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4110
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   8700
      Picture         =   "frmKleihauerQC.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   1005
      Left            =   8700
      Picture         =   "frmKleihauerQC.frx":2BC4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1185
      Left            =   8400
      TabIndex        =   1
      Top             =   90
      Width           =   1845
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   270
         TabIndex        =   3
         Top             =   750
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   39140
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   39140
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdKl 
      Height          =   7395
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   13044
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmKleihauerQC.frx":A0C6
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   165
      TabIndex        =   9
      Top             =   7740
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8700
      TabIndex        =   7
      Top             =   4830
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmKleihauerQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo FillG_Error

20    grdKl.Rows = 2
30    grdKl.AddItem ""
40    grdKl.RemoveItem 1

50    CheckKleihauerQCInDb

60    sql = "SELECT * FROM KleihauerQC WHERE " & "DateTime BETWEEN '" & _
          Format$(dtFrom, "long date") & "' AND '" & Format$(dtTo, _
          "long date") & " 23:59:59'"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql
90    Do While Not tb.EOF
100     s = tb!DateTime & vbTab & tb!SampleID & vbTab & tb!Rhesus & vbTab
110     Select Case tb!Positive & ""
          Case "P": s = s & "Pass"
120       Case "F": s = s & "Fail"
130       Case "N": s = s & "Not Checked"
140     End Select
150     s = s & vbTab
160     Select Case tb!Negative & ""
          Case "P": s = s & "Pass"
170       Case "F": s = s & "Fail"
180       Case "N": s = s & "Not Checked"
190     End Select
200     s = s & vbTab & tb!Operator & ""
210     grdKl.AddItem s
220     tb.MoveNext
230   Loop

240   If grdKl.Rows > 2 Then
250     grdKl.RemoveItem 1
260   End If

270   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmKleihauerQC", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub cmdPrint_Click()

      Dim Orig As String
      Dim Y As Integer
      Dim Px As Printer

10    Orig = UCase$(Printer.DeviceName)
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Font.Name = "Courier New"
40    Printer.Font.Size = 14
50    Printer.Print "Kleihauer QC"
60    Printer.Print
70    Printer.Font.Size = 10
80    Printer.Print "Date / Time";
90    Printer.Print Tab(23); "Sample ID";
100   Printer.Print Tab(33); "Rhesus";
110   Printer.Print Tab(46); "Pos.Cont";
120   Printer.Print Tab(60); "Neg.Cont";
130   Printer.Print Tab(74); "Operator"
140   Printer.Print

150   For Y = 1 To grdKl.Rows - 1
160     Printer.Print grdKl.TextMatrix(Y, 0);
170     Printer.Print Tab(23); grdKl.TextMatrix(Y, 1);
180     Printer.Print Tab(33); grdKl.TextMatrix(Y, 2);
190     Printer.Print Tab(46); grdKl.TextMatrix(Y, 3);
200     Printer.Print Tab(60); grdKl.TextMatrix(Y, 4);
210     Printer.Print Tab(74); grdKl.TextMatrix(Y, 5)
220   Next
230   Printer.EndDoc

240   For Each Px In Printers
250     If UCase$(Px.DeviceName) = Orig Then
260       Set Printer = Px
270       Exit For
280     End If
290   Next

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub cmdSearch_Click()

10    FillG

End Sub


Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid grdKl, Me

End Sub

Private Sub cmdXL_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub dtFrom_CallbackKeyDown(ByVal KeyCode As Integer, _
    ByVal Shift As Integer, ByVal CallbackField As String, _
    CallbackDate As Date)

End Sub

Private Sub dtFrom_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub dtTo_CallbackKeyDown(ByVal KeyCode As Integer, _
    ByVal Shift As Integer, ByVal CallbackField As String, _
    CallbackDate As Date)

End Sub

Private Sub dtTo_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub Form_Load()

10    dtFrom = Now - 30
20    dtTo = Now

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, _
    Y As Single)

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
End Sub

