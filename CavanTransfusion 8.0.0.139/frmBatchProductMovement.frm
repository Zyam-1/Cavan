VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBatchProductMovement 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Batch Product Movement"
   ClientHeight    =   8310
   ClientLeft      =   300
   ClientTop       =   525
   ClientWidth     =   16065
   ForeColor       =   &H80000008&
   Icon            =   "frmBatchProductMovement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   16065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIdentifier 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1860
      TabIndex        =   21
      Top             =   90
      Width           =   2205
   End
   Begin VB.CommandButton cmdExDes 
      Appearance      =   0  'Flat
      Caption         =   "Expire && Destroy"
      Enabled         =   0   'False
      Height          =   500
      Left            =   3675
      TabIndex        =   15
      ToolTipText     =   "Expire & Destroy"
      Top             =   7140
      Width           =   1500
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   7965
      Width           =   16065
      _ExtentX        =   28337
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "12/10/2020"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "16:08"
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
      Left            =   10800
      TabIndex        =   10
      ToolTipText     =   "Dispatch"
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdTransfuse 
      Appearance      =   0  'Flat
      Caption         =   "Transfuse"
      Enabled         =   0   'False
      Height          =   500
      Left            =   7230
      TabIndex        =   9
      ToolTipText     =   "Transfuse"
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdDestroy 
      Appearance      =   0  'Flat
      Caption         =   "Destroy"
      Enabled         =   0   'False
      Height          =   500
      Left            =   9015
      TabIndex        =   8
      ToolTipText     =   "Destroy"
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdReplace 
      Appearance      =   0  'Flat
      Caption         =   "Return to Stock"
      Enabled         =   0   'False
      Height          =   500
      Left            =   5460
      TabIndex        =   7
      ToolTipText     =   "Return to Stock"
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdReturn 
      Appearance      =   0  'Flat
      Caption         =   "Return to Supplier"
      Enabled         =   0   'False
      Height          =   500
      Left            =   1890
      TabIndex        =   6
      ToolTipText     =   "Return to Supplier"
      Top             =   7140
      Width           =   1500
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1065
      Left            =   14850
      Picture         =   "frmBatchProductMovement.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print History"
      Top             =   960
      Width           =   930
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   1065
      Left            =   14850
      Picture         =   "frmBatchProductMovement.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "cancel"
      ToolTipText     =   "Exit"
      Top             =   5940
      Width           =   930
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   1860
      TabIndex        =   19
      Top             =   7680
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6045
      Left            =   60
      TabIndex        =   23
      Top             =   960
      Width           =   14595
      _ExtentX        =   25744
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
      FormatString    =   $"frmBatchProductMovement.frx":265E
   End
   Begin VB.Label lblBatchNumber 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   22
      Top             =   570
      Width           =   2205
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Identifier"
      Height          =   195
      Left            =   1230
      TabIndex        =   20
      Top             =   188
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ml or IU"
      Height          =   195
      Left            =   12270
      TabIndex        =   18
      Top             =   195
      Width           =   540
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
      Left            =   11070
      TabIndex        =   17
      ToolTipText     =   "Volume/Dose"
      Top             =   150
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume/Dose"
      Height          =   195
      Left            =   9720
      TabIndex        =   16
      Top             =   195
      Width           =   1305
   End
   Begin VB.Label lblProduct 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5340
      TabIndex        =   14
      ToolTipText     =   "Product Name"
      Top             =   150
      Width           =   3915
   End
   Begin VB.Label lblGroup 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8370
      TabIndex        =   12
      ToolTipText     =   "Batch Group"
      Top             =   570
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   7890
      TabIndex        =   11
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblExpiry 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5370
      TabIndex        =   3
      ToolTipText     =   "Expiry Date"
      Top             =   570
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   4860
      TabIndex        =   5
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   4710
      TabIndex        =   4
      Top             =   188
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   600
      Width           =   1020
   End
End
Attribute VB_Name = "frmBatchProductMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BPs As BatchProducts
Private BP As BatchProduct
Private Sub DisableButtons()

10    cmdTransfuse.Enabled = False
20    cmdReturn.Enabled = False
30    cmdExDes.Enabled = False
40    cmdReplace.Enabled = False
50    cmdDestroy.Enabled = False
60    cmdDispatch.Enabled = False

End Sub

Private Sub cmdDispatch_Click()

    Dim s As String
    Dim Reason As String

10    On Error GoTo cmdDispatch_Click_Error

20    s = "Confirm Product to be Dispatched."
30    Answer = iMsg(s, vbYesNo + vbQuestion)
40    If TimedOut Then Unload Me: Exit Sub
50    If Answer = vbYes Then
60      Reason = iBOX("Why is this being dispatched?")
70      If TimedOut Then Unload Me: Exit Sub
80      If Trim$(Reason) <> "" Then
90          LogReasonWhy Reason, "F"
100         BP.EventCode = "F"
110         BP.Comment = Reason
120         BP.UserName = AddTicks(UserName)
130         BPs.Update BP
140     End If
150   End If

160   FillGrid

    '*******************************
    'Log unit fate activity in courier interface requests table to update
    'status in blood courier management system by Blood Courier Interface
    'Send FT signal with sample status as U (Unknow- Dispatched)

    Dim MSG As udtRS
170   With MSG
180     .UnitNumber = txtIdentifier
190     .ProductCode = lblProduct
200     .UnitExpiryDate = lblExpiry
210     .SampleStatus = "U"
220     .StorageLocation = ""
230     .ActionText = "Dispatch"
240     .UserName = UserName
250   End With
260   LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT


270   Exit Sub

cmdDispatch_Click_Error:

    Dim strES As String
    Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmBatchProductMovement", "cmdDispatch_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdDestroy_Click()

    Dim s As String
    Dim Reason As String

10    On Error GoTo cmdDestroy_Click_Error

20    s = "Confirm Product to be Destroyed."
30    Answer = iMsg(s, vbYesNo + vbQuestion)
40    If TimedOut Then Unload Me: Exit Sub
50    If Answer = vbYes Then
60      Reason = iBOX("Why is this being destroyed?")
70      If TimedOut Then Unload Me: Exit Sub
80      If Trim$(Reason) <> "" Then
90          LogReasonWhy Reason, "D"
100         BP.EventCode = "D"
110         BP.Comment = Reason
120         BP.UserName = AddTicks(UserName)
130         BPs.Update BP
140     End If
150   End If

160   FillGrid

    '*******************************
    'Log unit fate activity in courier interface requests table to update
    'status in blood courier management system by Blood Courier Interface
    'Send FT signal with sample status as D (Destroyed)

    Dim MSG As udtRS
170   With MSG
180     .UnitNumber = txtIdentifier
190     .ProductCode = lblProduct
200     .UnitExpiryDate = lblExpiry
210     .SampleStatus = "D"
220     .StorageLocation = ""
230     .ActionText = "Destroy"
240     .UserName = UserName
250   End With
260   LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT


270   Exit Sub

cmdDestroy_Click_Error:

    Dim strES As String
    Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmBatchProductMovement", "cmdDestroy_Click", intEL, strES

End Sub

Private Sub cmdExDes_Click()

    Dim s As String

10    On Error GoTo cmdExDes_Click_Error

20    If Not BP Is Nothing Then

30      s = "Confirm Product to be Destroyed."
40      Answer = iMsg(s, vbYesNo + vbQuestion)
50      If TimedOut Then Unload Me: Exit Sub
60      If Answer = vbYes Then
70          LogReasonWhy "Expired", "J"
80          BP.EventCode = "J"
90          BP.Comment = "Expired"
100         BP.UserName = AddTicks(UserName)
110         BPs.Update BP
120     End If

130   End If

140   FillGrid

    '*******************************
    'Log unit fate activity in courier interface requests table to update
    'status in blood courier management system by Blood Courier Interface
    'Send FT signal with sample status as D (Destroyed)

    Dim MSG As udtRS
150   With MSG
160     .UnitNumber = txtIdentifier
170     .ProductCode = lblProduct
180     .UnitExpiryDate = lblExpiry
190     .SampleStatus = "D"
200     .StorageLocation = ""
210     .ActionText = "Expire Destroy"
220     .UserName = UserName
230   End With
240   LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT

250   Exit Sub

cmdExDes_Click_Error:

    Dim strES As String
    Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmBatchProductMovement", "cmdExDes_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

    Dim Y As Integer
    Dim OriginalPrinter As String
    Dim Px As Printer

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName

30    If Not SetFormPrinter() Then Exit Sub

40    Printer.Orientation = vbPRORLandscape
    'FillGrid
50    Printer.Font.Name = "Courier New"
60    Printer.Print
70    Printer.Font.Size = 10
80    Printer.Font.Bold = True

90    Printer.Print "Identifier  : "; txtIdentifier;
100   Printer.Print "Unit Number : "; lblBatchNumber;
110   Printer.Print Tab(30); "Product : "; lblProduct
120   Printer.Print "Expiry Date : "; lblExpiry;
130   Printer.Print Tab(30); "Group : "; Left$(lblGroup, 2)
140   Printer.Print

150   Printer.Font.Size = 9
160   For Y = 1 To 152
170     Printer.Print "-";
180   Next Y
190   Printer.Print
200   Printer.Print FormatString("Lab#", 10, "|");
210   Printer.Print FormatString("Event", 16, "|");
220   Printer.Print FormatString("Date", 25, "|");
230   Printer.Print FormatString("Patient Name", 10, "|");
240   Printer.Print FormatString("Chart", 25, "|");
250   Printer.Print FormatString("User", 5, "|");
260   Printer.Print FormatString("Start Date/Time", 21, "|");
270   Printer.Print FormatString("End Date/Time", 21, "|")
280   For Y = 1 To 152
290     Printer.Print "-";
300   Next Y
310   Printer.Print
320   Printer.Font.Bold = False
330   For Y = 1 To g.Rows - 1
340     Printer.Print FormatString(g.TextMatrix(Y, 0), 10, "|");    'lab no
350     Printer.Print FormatString(g.TextMatrix(Y, 1), 16, "|");    'date
360     Printer.Print FormatString(g.TextMatrix(Y, 2), 25, "|");    'event
370     Printer.Print FormatString(g.TextMatrix(Y, 3), 10, "|");    'ID
380     Printer.Print FormatString(g.TextMatrix(Y, 4), 25, "|");    'Name
390     Printer.Print FormatString(g.TextMatrix(Y, 8), 5, "|");    'User
400     Printer.Print FormatString(g.TextMatrix(Y, 9), 10, "|");    'Units
410     Printer.Print FormatString(g.TextMatrix(Y, 10), 21, "|");    'Start date
420     Printer.Print FormatString(g.TextMatrix(Y, 10), 21, "|")    'End Date
430   Next

440   Printer.Print

450   Printer.EndDoc

460   For Each Px In Printers
470     If Px.DeviceName = OriginalPrinter Then
480         Set Printer = Px
490         Exit For
500     End If
510   Next

520   Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmBatchProductMovement", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdReplace_Click()

    Dim s As String

10    On Error GoTo cmdReplace_Click_Error

20    If Not BP Is Nothing Then

30      s = "Confirm Product to be Returned to Stock."
40      Answer = iMsg(s, vbYesNo + vbQuestion)
50      If TimedOut Then Unload Me: Exit Sub
60      If Answer = vbYes Then
70          BP.EventCode = "R"
80          BP.AandE = ""
90          BP.Addr0 = ""
100         BP.Addr1 = ""
110         BP.Addr2 = ""
120         BP.Age = ""
130         BP.Chart = ""
140         BP.Clinician = ""
150         BP.DoB = ""
160         BP.PatientGroup = ""
170         BP.PatName = ""
180         BP.SampleID = ""
190         BP.Sex = ""
200         BP.Typenex = ""
210         BP.Ward = ""
220         BP.Comment = ""
230         BP.LabelPrinted = 0
240         BP.UserName = AddTicks(UserName)
250         BPs.Update BP
260     End If

270   End If

280   FillGrid


    'Send RTS signal to Courier
    Dim Generic As String
    Dim MSG As udtRS
290   With MSG
300     .UnitNumber = txtIdentifier
310     .ProductCode = lblProduct
320     .UnitExpiryDate = lblExpiry
330     .StorageLocation = strBTCourier_StorageLocation_StockFridge
340     .ActionText = "Return to Stock"
350     .UserName = UserName
360   End With
370   LogCourierInterface "RTS", MSG

    'Cavan requested that "RTS" message followed with "SU"
380   With MSG
390     .UnitNumber = txtIdentifier
400     .ProductCode = lblProduct
410     Generic = ""    'ProductGenericFor(.ProductCode)
420     If Generic = "Platelets" Then
430         .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
440     Else
450         .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
460     End If
470     .UnitExpiryDate = Format(lblExpiry, "dd/mmm/yyyy HH:mm")
480     .UnitGroup = lblGroup
    'No real Patient details when unit Restocked
490     .StockComment = ""
500     .Chart = ""
510     .PatientHealthServiceNumber = ""
520     .ForeName = ""
530     .SurName = ""
540     .DoB = ""
550     .Sex = ""
560     .PatientGroup = ""
570     .DeReservationDateTime = Format(lblExpiry, "dd-MMM-yyyy hh:mm:ss")
580     .ActionText = "Stock Update"
590     .UserName = UserName
600   End With
610   LogCourierInterface "SU3", MSG


620   Exit Sub

cmdReplace_Click_Error:

    Dim strES As String
    Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmBatchProductMovement", "cmdReplace_Click", intEL, strES

End Sub

Private Sub cmdReturn_Click()

    Dim s As String
    Dim Reason As String

10    On Error GoTo cmdReturn_Click_Error

20    If Not BP Is Nothing Then

30      s = "Confirm Product to be Returned."
40      Answer = iMsg(s, vbYesNo + vbQuestion)
50      If TimedOut Then Unload Me: Exit Sub
60      If Answer = vbYes Then
70          Reason = iBOX("Why is this being Returned?")
80          If TimedOut Then Unload Me: Exit Sub
90          If Trim$(Reason) <> "" Then
100             LogReasonWhy Reason, "T"
110             BP.EventCode = "T"
120             BP.Comment = Reason
130             BP.UserName = AddTicks(UserName)
140             BPs.Update BP
150         End If
160     End If

170   End If

    'Send FT signal to Courier with sample status as U (Unknown - Returned to supplier)

    Dim MSG As udtRS
180   With MSG
190     .UnitNumber = txtIdentifier
200     .ProductCode = lblProduct
210     .UnitExpiryDate = lblExpiry
220     .SampleStatus = "US"
230     .StorageLocation = ""
240     .ActionText = "Return To Supplier"
250     .UserName = UserName
260   End With
270   LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT

280   FillGrid

290   Exit Sub

cmdReturn_Click_Error:

    Dim strES As String
    Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmBatchProductMovement", "cmdReturn_Click", intEL, strES

End Sub

Private Sub cmdTransfuse_Click()

          Dim s As String
          Dim PatName As String
          Dim SurName As String
          Dim ForeName As String
          Dim Y As Long
          Dim Chart As String
          Dim sql As String
          Dim tb As Recordset
          Dim Sex As String
          Dim DoB As String

10        On Error GoTo cmdTransfuse_Click_Error

20        If g.TextMatrix(g.row, 1) <> "Issued" Then
30            iMsg "Product must be Issued First!", vbInformation
40            If TimedOut Then Unload Me: Exit Sub
50            Exit Sub
60        End If


70        If g.TextMatrix(g.row, 2) = "Received" Or _
             g.TextMatrix(g.row, 2) = "Restocked" Or _
             g.TextMatrix(g.row, 2) = "Transfused" Then Exit Sub

80        s = "Confirm Product Transfused." & vbCrLf & _
              "Patient Chart : " & g.TextMatrix(g.row, 4) & vbCrLf & _
              "  Patient Name : " & g.TextMatrix(g.row, 3)
90        Answer = iMsg(s, vbYesNo + vbQuestion)
100       If TimedOut Then Unload Me: Exit Sub
110       If Answer = vbYes Then
120           BP.Addr0 = g.TextMatrix(g.row, 5)
130           BP.Chart = g.TextMatrix(g.row, 4)
140           Chart = g.TextMatrix(g.row, 4)
150           BP.EventCode = "S"
160           BP.PatName = g.TextMatrix(g.row, 3)
170           PatName = g.TextMatrix(g.row, 3)
180           BP.SampleID = g.TextMatrix(g.row, 0)
190           BP.UserName = AddTicks(UserName)
200           BPs.Update BP
210       End If

220       FillGrid

          '*******************************
          'Log unit fate activity in courier interface requests table to update
          'status in blood courier management system by Blood Courier Interface
          'Implemented Site: CAVAN General Hospital

230       Y = InStr(PatName, " ")
240       If Y <> 0 Then
250           SurName = Left$(PatName, Y - 1)
260           ForeName = Mid$(PatName, Y + 1)
270       Else
280           SurName = PatName
290           ForeName = ""
300       End If
310       SurName = UCase$(SurName)


320       sql = "SELECT PatSurName, PatForeName FROM PatientDetails WHERE " & _
                "LabNumber = '" & g.TextMatrix(g.row, 0) & "' " & _
                "ORDER BY DateTime DESC"
330       Set tb = New Recordset
340       RecOpenServerBB 0, tb, sql
350       If Not tb.EOF Then
            If Trim$(tb!PatSurName & "" & tb!PatForeName & "") <> "" Then
360           SurName = Trim$(tb!PatSurName & "")
370           ForeName = Trim$(tb!PatForeName & "")
380           SurName = UCase$(SurName)
          End If
390       End If


400       sql = "SELECT Sex, DoB, PatSurName, PatForeName FROM PatientDetails WHERE " & _
                "PatNum = '" & Chart & "' " & _
                "ORDER BY DateTime DESC"
410       Set tb = New Recordset
420       RecOpenServerBB 0, tb, sql
430       If tb.EOF Then
440           Sex = "U"
450           DoB = ""
460       Else
470           Sex = tb!Sex & ""
480           If Not IsNull(tb!DoB) Then
490               If IsDate(tb!DoB) Then
500                   DoB = Format$(tb!DoB, "dd-MMM-yyyy")
510               Else
520                   DoB = ""
530               End If
540           Else
550               DoB = ""
560           End If
570       End If

          'Send FT signal with sample status as U (Unknow- Returned to supplier)

          Dim MSG As udtRS
580       With MSG
590           .UnitNumber = txtIdentifier
600           .ProductCode = lblProduct
610           .UnitExpiryDate = lblExpiry
620           .Chart = Chart
630           .PatientHealthServiceNumber = ""
640           .SurName = SurName
650           .ForeName = ForeName
660           .DoB = DoB
670           .Sex = Sex
680           .SampleStatus = "T"
690           .StorageLocation = ""
700           .ActionText = "Transfuse"
710           .UserName = UserName
720       End With
730       LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT


740       Exit Sub

cmdTransfuse_Click_Error:

          Dim strES As String
          Dim intEL As Integer

750       intEL = Erl
760       strES = Err.Description
770       LogError "frmBatchProductMovement", "cmdTransfuse_Click", intEL, strES

End Sub

Private Sub FillGrid()

    Dim s As String

10    g.Rows = 2
20    g.AddItem ""
30    g.RemoveItem 1

40    Set BPs = New BatchProducts

50    If Trim$(txtIdentifier) <> "" Then
60      BPs.LoadSpecificIdentifier txtIdentifier
70      If BPs.Count = 0 Then
80          iMsg "Identifier not found", vbCritical
90          If TimedOut Then Unload Me: Exit Sub
100         Exit Sub
110     End If
120   End If

130   FillCommonDetails BPs.Item(1)

140   For Each BP In BPs

150     s = BP.SampleID & vbTab & _
            gEVENTCODES(BP.EventCode).Text & vbTab & _
            Format(BP.RecordDateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
            BP.PatName & vbTab & _
            BP.Chart & vbTab & _
            BP.Addr0 & vbTab
160     If Format$(BP.EventStart, "dd/MM/yyyy") <> "01/01/1900" Then
170         s = s & Format(BP.EventStart, "dd/MM/yyyy HH:nn")
180     End If
190     s = s & vbTab
200     If Format$(BP.EventEnd, "dd/MM/yyyy") <> "01/01/1900" Then
210         s = s & Format(BP.EventEnd, "dd/MM/yyyy HH:nn")
220     End If
230     s = s & vbTab & _
            BP.UserName & vbTab & _
            BP.Comment
240     g.AddItem s
250   Next

260   If g.Rows > 2 Then
270     g.RemoveItem 1
280   End If

290   cmdPrint.Enabled = True

300   DisableButtons
End Sub


Private Sub SetButtons()

    Dim E As String

10    DisableButtons

20    E = gEVENTCODES.CodeFor(g.TextMatrix(1, 1))
30    If E = "D" Or E = "F" Or E = "T" Then    'Destroyed, Despatched or Returned to Supplier
40      Exit Sub
50    End If

60    Select Case gEVENTCODES.CodeFor(g.TextMatrix(g.row, 1))
        Case "C", "R", "?????":    'last event was "Received" or "Restocked"
70          cmdReturn.Enabled = True
80          cmdExDes.Enabled = True
90          cmdDestroy.Enabled = True
100         cmdDispatch.Enabled = True
110     Case "X", "P", "I":    'last event was "Xmatched" or "Pending" or "Issued"
120         cmdTransfuse.Enabled = True
130         cmdReturn.Enabled = True
140         cmdExDes.Enabled = True
150         cmdReplace.Enabled = True
160         cmdDestroy.Enabled = True
170         cmdDispatch.Enabled = True
180     Case "Q":
190         cmdReturn.Enabled = True
200         cmdExDes.Enabled = True
210         cmdReplace.Enabled = True
220         cmdDestroy.Enabled = True
230         cmdDispatch.Enabled = True
240     Case "W", "Z", "M":
250         cmdReturn.Enabled = True
260         cmdExDes.Enabled = True
270         cmdReplace.Enabled = True
280         cmdDestroy.Enabled = True
290         cmdDispatch.Enabled = True
300   End Select

End Sub

Private Sub Form_Load()

10    StatusBar.Panels(3) = UserName

End Sub

Private Sub FillCommonDetails(ByVal BP As BatchProduct)

10    lblBatchNumber = BP.BatchNumber
20    lblGroup = BP.UnitGroup
30    lblExpiry = BP.DateExpiry
40    lblProduct = BP.Product
50    lblDose = BP.UnitVolume
60    If lblProduct = "Anti-D" Then
70      lblDose = lblDose & " IU"
80    End If

End Sub

Private Sub g_Click()

    Dim BPTemp As BatchProduct

10    On Error GoTo g_Click_Error

20    If g.Rows = 2 And g.TextMatrix(1, 1) = "" Then
30      Exit Sub
40    End If

50    HighlightGridRow g

60    For Each BPTemp In BPs
70      If Format$(BPTemp.RecordDateTime, "dd/MMM/yyyy HH:nn:ss") = Format$(g.TextMatrix(g.row, 2), "dd/MMM/yyyy HH:nn:ss") Then
80          Set BP = BPTemp
90          Exit For
100     End If
110   Next

120   If g.TextMatrix(g.row, 1) = "Issued" Then
130     cmdTransfuse.Enabled = True
140   End If

150   SetButtons

160   Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmBatchProductMovement", "g_Click", intEL, strES

End Sub


Private Sub txtIdentifier_LostFocus()

    Dim C As String

10    On Error GoTo txtIdentifier_LostFocus_Error

20    txtIdentifier = UCase$(txtIdentifier)

30    If Trim$(txtIdentifier) = "" Then
40      iMsg "Enter or scan Identifier", vbInformation
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70    End If

    '80    If Len(txtIdentifier) <> 9 Then
    '90      iMsg "Identifier should be 9 characters", vbInformation
    '100     If TimedOut Then Unload Me: Exit Sub
    '110     Exit Sub
    '120   End If

80    C = CheckCharacterForBatch(Left$(txtIdentifier, 8))
90    If Right$(txtIdentifier, 1) <> C Then
100     iMsg "Check character incorrect", vbCritical
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   Set BPs = New BatchProducts

150   If Trim$(txtIdentifier) <> "" Then
160     BPs.LoadSpecificIdentifier txtIdentifier
170     If BPs.Count = 0 Then
180         iMsg "Identifier not found", vbCritical
190         If TimedOut Then Unload Me: Exit Sub
200         Exit Sub
210     End If
220     FillGrid
230   End If

240   Exit Sub

txtIdentifier_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmBatchProductMovement", "txtIdentifier_LostFocus", intEL, strES

End Sub



Public Property Let Identifier(ByVal sNewValue As String)

10    txtIdentifier = sNewValue
20    FillGrid

End Property
