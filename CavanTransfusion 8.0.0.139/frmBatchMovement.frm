VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchMovement 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Batch Product Movement"
   ClientHeight    =   8130
   ClientLeft      =   300
   ClientTop       =   525
   ClientWidth     =   14640
   ForeColor       =   &H80000008&
   Icon            =   "frmBatchMovement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8130
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExDes 
      Appearance      =   0  'Flat
      Caption         =   "Expire && Destroy"
      Enabled         =   0   'False
      Height          =   500
      Left            =   720
      TabIndex        =   17
      ToolTipText     =   "Expire & Destroy"
      Top             =   6870
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid grdBat 
      Height          =   4350
      Left            =   60
      TabIndex        =   15
      Top             =   1710
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   13
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
      FormatString    =   $"frmBatchMovement.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   7785
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "12/07/2012"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "15:46"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDispatch 
      Appearance      =   0  'Flat
      Caption         =   "Dispatch"
      Enabled         =   0   'False
      Height          =   500
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Dispatch"
      Top             =   6810
      Width           =   1500
   End
   Begin VB.CommandButton cmdTransfuse 
      Appearance      =   0  'Flat
      Caption         =   "Transfuse"
      Enabled         =   0   'False
      Height          =   500
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "Transfuse"
      Top             =   6840
      Width           =   1470
   End
   Begin VB.CommandButton cmdDestroy 
      Appearance      =   0  'Flat
      Caption         =   "Destroy"
      Enabled         =   0   'False
      Height          =   500
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "Destroy"
      Top             =   6210
      Width           =   1500
   End
   Begin VB.CommandButton cmdReplace 
      Appearance      =   0  'Flat
      Caption         =   "Return to Stock"
      Enabled         =   0   'False
      Height          =   500
      Left            =   3825
      TabIndex        =   8
      ToolTipText     =   "Return to Stock"
      Top             =   6240
      Width           =   1500
   End
   Begin VB.CommandButton cmdReturn 
      Appearance      =   0  'Flat
      Caption         =   "Return to Supplier"
      Enabled         =   0   'False
      Height          =   500
      Left            =   735
      TabIndex        =   7
      ToolTipText     =   "Return to Supplier"
      Top             =   6255
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   645
      Left            =   12570
      Picture         =   "frmBatchMovement.frx":09AC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print History"
      Top             =   6630
      Width           =   810
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   660
      Left            =   13590
      Picture         =   "frmBatchMovement.frx":0CB6
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "cancel"
      ToolTipText     =   "Exit"
      Top             =   6630
      Width           =   855
   End
   Begin VB.TextBox txtUnitNumber 
      BackColor       =   &H80000018&
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
      Left            =   825
      MaxLength       =   14
      TabIndex        =   0
      ToolTipText     =   "Batch Number"
      Top             =   195
      Width           =   2025
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   25
      Top             =   7560
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblCurrentStock 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
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
      Left            =   6360
      TabIndex        =   24
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Current Stock"
      Height          =   195
      Left            =   5340
      TabIndex        =   23
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ml or IU"
      Height          =   195
      Left            =   4380
      TabIndex        =   22
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lblSupp 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   6360
      TabIndex        =   21
      ToolTipText     =   "Supplier"
      Top             =   600
      Width           =   2955
   End
   Begin VB.Label lbSupp 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
      Height          =   195
      Left            =   5730
      TabIndex        =   20
      Top             =   660
      Width           =   570
   End
   Begin VB.Label lblDose 
      BackColor       =   &H80000018&
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
      Left            =   1410
      TabIndex        =   19
      ToolTipText     =   "Volume/Dose"
      Top             =   1080
      Width           =   2940
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume/Dose"
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   1125
      Width           =   1305
   End
   Begin VB.Label lblProd 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   825
      TabIndex        =   16
      ToolTipText     =   "Product Name"
      Top             =   645
      Width           =   4635
   End
   Begin VB.Label lblGroup 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      ToolTipText     =   "Batch Group"
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   240
      Width           =   435
   End
   Begin VB.Label lblExpiry 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3450
      TabIndex        =   4
      ToolTipText     =   "Expiry Date"
      Top             =   195
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   3000
      TabIndex        =   6
      Top             =   210
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   195
      TabIndex        =   5
      Top             =   645
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   225
      Width           =   555
   End
End
Attribute VB_Name = "frmBatchMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim s As String
Dim GroupRh As String
Dim PatName As String
Dim Patid As String
Dim NameSelected As Boolean
Dim InStock As Integer
Dim BatNum As String
Private pSampleID As String
Dim intAvailableForTransfusion As Integer

Private Sub DisableButtons()
   
10    cmdTransfuse.Enabled = False
20    cmdReturn.Enabled = False
30    cmdExDes.Enabled = False
40    cmdReplace.Enabled = False
50    cmdDestroy.Enabled = False
60    cmdDispatch.Enabled = False

End Sub

Private Sub cmdDispatch_Click()

10    If QueryValidate("F") Then
20      Validate "F"
30    End If

40    FillGrid

End Sub

Private Sub cmdCancel_Click()

10    BatNum = ""
20    Unload Me
  
End Sub

Private Sub cmdDestroy_Click()

      Dim s As String
      Dim Reason As String

10    s = "Confirm Product to be Destroyed."
20    Answer = iMsg(s, vbYesNo + vbQuestion)
30    If TimedOut Then Unload Me: Exit Sub
40    If Answer = vbYes Then
50      Reason = iBOX("Why is this being destroyed?")
60      If TimedOut Then Unload Me: Exit Sub
70      If Trim$(Reason) <> "" Then
80        LogReasonWhy Reason, "D"
90        Validate "D", Reason
100     End If
110   End If

120   FillGrid

End Sub

Private Sub cmdExDes_Click()

10    Validate "J"

20    FillGrid

End Sub

Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim OriginalPrinter As String
      Dim Px As Printer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    If Trim$(txtUnitNumber) = "" Then
40       iMsg "Unit Number?", vbQuestion
50      If TimedOut Then Unload Me: Exit Sub
60       txtUnitNumber.SetFocus
70       Exit Sub
80    End If

90    Printer.Orientation = vbPRORLandscape
      'FillGrid
100   Printer.Font.Name = "Courier New"
110   Printer.Print
120   Printer.Font.Size = 10
130   Printer.Font.Bold = True

140   Printer.Print "Unit Number : "; txtUnitNumber;
150   Printer.Print Tab(30); "Product : "; lblProd
160   Printer.Print "Expiry Date : "; lblExpiry;
170   Printer.Print Tab(30); "Group : "; Left$(lblGroup, 2)
180   Printer.Print

190   Printer.Font.Size = 9
200   For Y = 1 To 152
210       Printer.Print "-";
220   Next Y
230   Printer.Print
240   Printer.Print FormatString("Lab#", 10, "|");
250   Printer.Print FormatString("Date", 16, "|");
260   Printer.Print FormatString("Event", 25, "|");
270   Printer.Print FormatString("Patient ID", 10, "|");
280   Printer.Print FormatString("Name", 25, "|");
290   Printer.Print FormatString("User", 5, "|");
300   Printer.Print FormatString("Units", 10, "|");
310   Printer.Print FormatString("Start Date/Time", 21, "|");
320   Printer.Print FormatString("End Date/Time", 21, "|")
330   For Y = 1 To 152
340       Printer.Print "-";
350   Next Y
360   Printer.Print
370   Printer.Font.Bold = False
380   For Y = 1 To grdBat.Rows - 1
390      Printer.Print FormatString(grdBat.TextMatrix(Y, 0), 10, "|"); 'lab no
400      Printer.Print FormatString(grdBat.TextMatrix(Y, 1), 16, "|"); 'date
410      Printer.Print FormatString(grdBat.TextMatrix(Y, 2), 25, "|"); 'event
420      Printer.Print FormatString(grdBat.TextMatrix(Y, 3), 10, "|"); 'ID
430      Printer.Print FormatString(grdBat.TextMatrix(Y, 4), 25, "|"); 'Name
440      Printer.Print FormatString(grdBat.TextMatrix(Y, 8), 5, "|"); 'User
450      Printer.Print FormatString(grdBat.TextMatrix(Y, 9), 10, "|"); 'Units
460      Printer.Print FormatString(grdBat.TextMatrix(Y, 10), 21, "|"); 'Start date
470   Printer.Print FormatString(grdBat.TextMatrix(Y, 10), 21, "|") 'End Date
480   Next

490   Printer.Print

500   Printer.EndDoc

510   For Each Px In Printers
520     If Px.DeviceName = OriginalPrinter Then
530       Set Printer = Px
540       Exit For
550     End If
560   Next

End Sub

Private Sub cmdReplace_Click()

10    If QueryValidate("R") Then Validate "R"
20    FillGrid
  
End Sub

Private Sub cmdReturn_Click()

      Dim s As String
      Dim Reason As String

10    s = "Confirm Product to be Returned."
20    Answer = iMsg(s, vbYesNo + vbQuestion)
30    If TimedOut Then Unload Me: Exit Sub
40    If Answer = vbYes Then
50      Reason = iBOX("Why is this being Returned?")
60      If TimedOut Then Unload Me: Exit Sub
70      If Trim$(Reason) <> "" Then
80        LogReasonWhy Reason, "T"
90        Validate "T", Reason
100     End If
110   End If

120   FillGrid
  
End Sub

Private Sub cmdTransfuse_Click()

      Dim s As String

10    On Error GoTo cmdTransfuse_Click_Error

20    If Not NameSelected Then
30      iMsg "Select patient.", vbExclamation
40      If TimedOut Then Unload Me: Exit Sub
50       Exit Sub
60    End If

70    If grdBat.TextMatrix(grdBat.Row, 2) <> "Issued" Then
80      iMsg "Product must be Issued First!", vbInformation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110   End If

120   If grdBat.TextMatrix(grdBat.Row, 2) = "Received" Or _
         grdBat.TextMatrix(grdBat.Row, 2) = "Restocked" Or _
         grdBat.TextMatrix(grdBat.Row, 2) = "Transfused" Then Exit Sub
   
130   s = "Confirm Product Transfused." & vbCrLf & _
          "Patient Number : " & Patid & vbCrLf & _
          "  Patient Name : " & PatName
140   Answer = iMsg(s, vbYesNo + vbQuestion)
150   If TimedOut Then Unload Me: Exit Sub
160   If Answer = vbYes Then
170       Validate "S"
180   End If

190   FillGrid

200   cmdTransfuse.Enabled = False

210   Exit Sub

cmdTransfuse_Click_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmBatchMovement", "cmdTransfuse_Click", intEL, strES

End Sub

Private Sub FillGrid()

      Dim n As Integer
      Dim sn As Recordset
      Dim sql As String
      Dim s As String
      Dim C As Integer
      Dim intNumIssued As Integer
      Dim intReSorTra As Integer

10    On Error GoTo FillGrid_Error

20    InStock = 0

30    NameSelected = False

40    txtUnitNumber = Trim$(txtUnitNumber)

50    With grdBat
60       .Rows = 2
70       .AddItem ""
80       .RemoveItem 1
90    End With

100   lblDose = ""
110   If Trim$(txtUnitNumber) = "" Then Exit Sub

120   BatNum = Trim$(txtUnitNumber)

      'load all information for batchnumber
130   sql = "Select * from BatchDetails where " & _
            "BatchNumber = '" & Trim$(txtUnitNumber) & "' " & _
            "order by date desc, event desc"
140   Set sn = New Recordset
150   RecOpenServerBB 0, sn, sql
160   Do While Not sn.EOF
         'populate grid
170      s = IIf(sn!Event = "S" Or sn!Event = "I", sn!SampleID, "") & vbTab & _
             Format$(sn!Date, "dd/MM/yyyy hh:mm:ss") & vbTab & _
             gEVENTCODES(sn!Event & "").Text & vbTab & _
             IIf(sn!Event = "S" Or sn!Event = "I", sn!Chart, "") & vbTab & _
             IIf(sn!Event = "S" Or sn!Event = "I", sn!Name, "") & vbTab & _
             sn!DoB & vbTab & _
             sn!Typenex & vbTab & _
             sn!Ward & vbTab & _
             sn!UserCode & vbTab & _
             sn!Bottles & vbTab & sn!EventStart & vbTab & sn!EventEnd & vbTab & _
             sn!Comment & ""
180      grdBat.AddItem s
190      sn.MoveNext
200   Loop

210   Fill_BatchList

220   cmdPrint.Enabled = True

      'remove empty lines
230   If grdBat.Rows > 2 And grdBat.TextMatrix(1, 0) = "" Then
240      grdBat.RemoveItem 1
250   End If

      'highlight stock status
260   For n = 1 To grdBat.Rows - 1
270     grdBat.Row = n
280     If grdBat.TextMatrix(n, 0) = "Stock Status" Then
290       For C = 0 To 6
300         grdBat.Col = C
310         grdBat.CellBackColor = vbYellow
320        Next
330      ElseIf grdBat.TextMatrix(n, 1) = "Restocked." Then
340        For C = 0 To 6
350          grdBat.Col = C
360          grdBat.CellBackColor = vbGreen
370        Next
380      ElseIf grdBat.TextMatrix(n, 1) = "Received into Stock." Then
390        For C = 0 To 6
400          grdBat.Col = C
410          grdBat.CellBackColor = vbCyan
420        Next
430      End If
440   Next

450   For n = 1 To grdBat.Rows - 1
      'Count batches ready for Transfusion
460     If grdBat.TextMatrix(n, 2) = "Issued" Then
470       intNumIssued = intNumIssued + Val(grdBat.TextMatrix(n, 9))
480     End If
      'Count batches Transfused or Restocked
490     If grdBat.TextMatrix(n, 2) = "Transfused" Or grdBat.TextMatrix(n, 2) = "Restocked" Then
500       intReSorTra = intReSorTra + Val(grdBat.TextMatrix(n, 9))
510     End If

520   Next
530   intAvailableForTransfusion = intNumIssued - intReSorTra
540   If intAvailableForTransfusion < 0 Then intAvailableForTransfusion = 0

550   DisableButtons

560   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

570   intEL = Erl
580   strES = Err.Description
590   LogError "frmBatchMovement", "FillGrid", intEL, strES, sql

End Sub

Private Sub SetButtons()

10    DisableButtons

20    Select Case gEVENTCODES.CodeFor(grdBat.TextMatrix(grdBat.Row, 2))
         Case "C", "R", "?????": 'last event was "Received" or "Restocked"
30          cmdReturn.Enabled = True
40          cmdExDes.Enabled = True
            'cmdReplace.Enabled = True
50          cmdDestroy.Enabled = True
60          cmdDispatch.Enabled = True
70    Case "X", "P", "I": 'last event was "Xmatched" or "Pending" or "Issued"
80          cmdTransfuse.Enabled = True
90          cmdReturn.Enabled = True
100         cmdExDes.Enabled = True
110         cmdReplace.Enabled = True
120         cmdDestroy.Enabled = True
130         cmdDispatch.Enabled = True
140   Case "Q":
150         cmdReturn.Enabled = True
160         cmdExDes.Enabled = True
170         cmdReplace.Enabled = True
180         cmdDestroy.Enabled = True
190         cmdDispatch.Enabled = True
200   Case "W", "Z", "M":
210         cmdReturn.Enabled = True
220         cmdExDes.Enabled = True
230         cmdReplace.Enabled = True
240         cmdDestroy.Enabled = True
250         cmdDispatch.Enabled = True
260   End Select

End Sub

Private Sub Fill_BatchList()

      Dim sql As String
      Dim sn As Recordset

10    On Error GoTo Fill_BatchList_Error

20    sql = "select * from batchproductlist where " & _
            "batchnumber = '" & Trim$(txtUnitNumber) & "'"

30    Set sn = New Recordset
40    RecOpenServerBB 0, sn, sql
50    If sn.EOF Then
60      lblProd = "No Such Batch"
70    Else
80      GroupRh = Trim$(sn!Group & "")
90      lblGroup = Trim$(sn!Group & "")
100     lblExpiry = (sn!DateExpiry & "")
110     lblProd = sn!Product
120     lblDose = Trim$(sn!UnitVolume & "")
130     lblCurrentStock = sn!CurrentStock & ""
140     If lblProd = "Anti-D" Then
150      lblDose = lblDose & " IU"
160     End If
170   End If

180   Exit Sub

Fill_BatchList_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmBatchMovement", "Fill_BatchList", intEL, strES, sql

End Sub


Private Sub Form_Load()

10    StatusBar.Panels(3) = UserName

20    NameSelected = False

End Sub

Private Sub grdBat_Click()

      Dim tb As Recordset
      Dim SID As String
      Dim strName As String
      Dim sql As String

10    On Error GoTo grdBat_Click_Error

20    If grdBat.Rows = 2 And grdBat.TextMatrix(1, 1) = "" Then
30      Exit Sub
40    End If

50    HighlightGridRow grdBat

60    If grdBat.Col = 2 And grdBat.TextMatrix(grdBat.Row, 2) = "Dispatched." Then
70      Set tb = New Recordset
80      sql = "select * from dispatch where number = '" & txtUnitNumber & "'"
90      RecOpenServerBB 0, tb, sql
100     Do While Not tb.EOF
110       If Format$(tb!DateTime, "yyyy/MM/dd mm:hh") = Format$(grdBat.TextMatrix(grdBat.Row, 1), "yyyy/MM/dd mm:hh") Then
120         If Trim$(tb!Details) <> "" Then
130           iMsg Trim$(tb!Details), vbInformation
140           If TimedOut Then Unload Me: Exit Sub
150           Exit Do
160         End If
170       End If
180       tb.MoveNext
190     Loop
        'Show extra details for destroyed
200   ElseIf grdBat.Col = 2 And grdBat.TextMatrix(grdBat.Row, 2) = "Destroyed." Then
210     Set tb = New Recordset
220     sql = "select * from destroy where unit = '" & txtUnitNumber & "'"
230     RecOpenServerBB 0, tb, sql
240     Do While Not tb.EOF
250       If Format(tb!Expiry, "dd/MMM/yyyy") = Format$(grdBat.TextMatrix(grdBat.Row, 1), "dd/MMM/yyyy") Then
260         iMsg tb!Reason & "", vbInformation
270         If TimedOut Then Unload Me: Exit Sub
280         Exit Do
290       End If
300       tb.MoveNext
310     Loop
        'message for transfused product
      '260   ElseIf grdBat.Col = 2 And grdBat.TextMatrix(grdBat.Row, 2) = "Transfused" Then
      '270     iMsg "A Transfusion is an event that has no Stock Implications!", vbInformation
320     If TimedOut Then Unload Me: Exit Sub
330   End If

340   If grdBat.TextMatrix(grdBat.Row, 2) = "Issued" Then
350     SID = grdBat.TextMatrix(grdBat.Row, 0)
360     If Trim$(SID) <> "" Then
370       strName = grdBat.TextMatrix(grdBat.Row, 4)
380       If Trim$(strName) <> "" Then
390         Patid = grdBat.TextMatrix(grdBat.Row, 3)
400         PatName = strName
410         pSampleID = SID
420         NameSelected = True
430         cmdTransfuse.Enabled = True
440       End If
450     End If
460   End If

470   SetButtons

480   Exit Sub

grdBat_Click_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "frmBatchMovement", "grdBat_Click", intEL, strES, sql

End Sub


Private Function QueryValidate(ByVal EventCode As String) As Boolean

      'get the event code of request
      Dim s As String

10    Select Case EventCode
         Case "T": s = "Return to Supplier."
20       Case "R": s = "Restock product."
30       Case "D": s = "Product has been destroyed."
40       Case "S": s = "Product transfused."
50       Case "F": s = "Inter-Hospital Transfer" & vbCrLf & "or Laboratory use."
60       Case "K": s = "Expired & Returned"
70       Case "J": s = "Expired & Destroyed"
80    End Select

90    Answer = iMsg(s, vbYesNo + vbQuestion)
100   If TimedOut Then Unload Me: Exit Function
110   QueryValidate = (IIf(Answer = vbYes, True, False))
  
End Function



Public Sub txtUnitNumber_LostFocus()

      'fill grid based on product number
10    FillGrid
  
End Sub

Private Sub Validate(ByVal EventCode As String, Optional ByVal strComment As String)

      Dim stock As String
      Dim tbLatest As Recordset
      Dim sql As String
      Dim lngCurrentStock As Long
      Dim strDetails As String
      Dim tb As Recordset
      Dim tbR As Recordset

      'check amount to change
10    On Error GoTo Validate_Error

20    If EventCode = "F" Or _
         EventCode = "S" Or _
         EventCode = "Z" Or _
         EventCode = "T" Or _
         EventCode = "M" Or _
         EventCode = "D" Or _
         EventCode = "R" Or _
         EventCode = "J" Then
30      stock = iBOX("Amount")
40      If TimedOut Then Unload Me: Exit Sub
50      If stock = "" Or Val(stock) = 0 Then
60        Exit Sub
70      End If
80    End If

90    If EventCode = "R" Then
100     If intAvailableForTransfusion - Val(stock) < 0 Then
110       iMsg "Issue stock first!"
120       If TimedOut Then Unload Me: Exit Sub
130       Exit Sub
140     End If
150   End If

      'Restock Product
160   If EventCode = "R" Then
170     lngCurrentStock = Restock_Prod(stock)
    
180     sql = "SELECT * FROM Reclaimed WHERE 0 = 1"
190     Set tbR = New Recordset
200     RecOpenServerBB 0, tbR, sql
210     With tbR
220       .AddNew
230       !Name = grdBat.TextMatrix(grdBat.Row, 4)
240       !Chart = grdBat.TextMatrix(grdBat.Row, 3)
250       !Unit = txtUnitNumber
260       !Group = lblGroup
270       !Product = lblProd
280       !xmdate = Null
290       !DateTime = Format(Now, "dd/mmm/yyyy HH:nn:ss")
300       !Operator = UserCode
310       !Ward = grdBat.TextMatrix(grdBat.Row, 7)
320       If IsDate(grdBat.TextMatrix(grdBat.Row, 5)) Then
330         !DoB = grdBat.TextMatrix(grdBat.Row, 5)
340       Else
350         !DoB = Null
360       End If
370       !Typenex = grdBat.TextMatrix(grdBat.Row, 6)
380       .Update
390     End With
      'Inter Hosp Transfer or Destroy
400   ElseIf EventCode = "J" Or _
             EventCode = "F" Or _
             EventCode = "I" Or _
             EventCode = "T" Or _
             (EventCode = "D" And grdBat.TextMatrix(grdBat.Row, 2) <> "Issued") Then
410     lngCurrentStock = Decrease_Stock(stock)
420     If lngCurrentStock = "-999" Then 'error
430       Exit Sub
440     End If
450   End If

460   If EventCode = "F" Then 'Dispatched
470     strDetails = iBOX("Enter Details.")
480     If TimedOut Then Unload Me: Exit Sub
490     strComment = strDetails
500     If Trim$(strDetails) = "" Then
510       iMsg "Cancelled", vbInformation
520       If TimedOut Then Unload Me: Exit Sub
530       Exit Sub
540     End If
  
550     sql = "Select * from Dispatch where Number = 'x'"
560     Set tb = New Recordset
570     RecOpenServerBB 0, tb, sql
580     tb.AddNew
590     tb!DateTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
600     tb!Number = txtUnitNumber
610     tb!Details = strDetails
620     tb.Update
630   End If

640   sql = "select * from batchdetails where " & _
            "batchnumber = '" & txtUnitNumber & "' " & _
            "Order by date desc"
650   Set tbLatest = New Recordset
660   RecOpenServerBB 0, tbLatest, sql
670   tbLatest.AddNew
680   tbLatest!Bottles = Val(stock)
690   tbLatest!Expiry = lblExpiry
      'tbLatest!UnitGroup = lblGroup
700   tbLatest!BatchNumber = txtUnitNumber
710   If EventCode <> "F" Then tbLatest!SampleID = pSampleID
720   If NameSelected Then
730     tbLatest!Name = grdBat.TextMatrix(grdBat.Row, 4)
740     tbLatest!Chart = grdBat.TextMatrix(grdBat.Row, 3)
750   End If
760   tbLatest!Product = lblProd
770   tbLatest!Event = EventCode
780   tbLatest!Date = Format$(Now, "dd/MM/yyyy hh:mm:ss")
790   tbLatest!UserCode = UserCode
800   If Len(strComment) > 0 Then
810     tbLatest!Comment = strComment
820   End If
830   tbLatest.Update

840   Exit Sub

Validate_Error:

      Dim strES As String
      Dim intEL As Integer

850   intEL = Erl
860   strES = Err.Description
870   LogError "frmBatchMovement", "Validate", intEL, strES, sql


End Sub


Private Function Restock_Prod(ByVal stock As String) As Long
      'returns new current stock

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Restock_Prod_Error

20    sql = "Select currentstock from batchproductlist where " & _
            "batchnumber = '" & txtUnitNumber & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    tb!CurrentStock = Val(tb!CurrentStock) + Val(stock)
60    Restock_Prod = tb!CurrentStock
70    tb.Update

80    Exit Function

Restock_Prod_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmBatchMovement", "Restock_Prod", intEL, strES, sql

End Function

Private Function Decrease_Stock(ByVal stock As String) As Long
      'returns new current stock

      Dim tb As Recordset
      Dim sql As String
      Dim lngNewStock As Long

10    On Error GoTo Decrease_Stock_Error

20    sql = "Select * from batchproductlist where " & _
            "batchnumber = '" & txtUnitNumber & "'"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql

50    lngNewStock = Val(tb!CurrentStock)

60    If lngNewStock < Val(stock) Then
70      iMsg "Current stock below level requested to move!", vbInformation
80      If TimedOut Then Unload Me:  Exit Function
90      Decrease_Stock = "-999" 'error
100     Exit Function
110   Else
120     lngNewStock = lngNewStock - Val(stock)
130     tb!CurrentStock = lngNewStock
140     tb.Update
150   End If
  
160   Decrease_Stock = lngNewStock

170   Exit Function

Decrease_Stock_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmBatchMovement", "Decrease_Stock", intEL, strES, sql


End Function




