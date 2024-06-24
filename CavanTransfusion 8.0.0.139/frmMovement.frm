VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmMovement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Movement"
   ClientHeight    =   8490
   ClientLeft      =   300
   ClientTop       =   720
   ClientWidth     =   15750
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
   Icon            =   "frmMovement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   15750
   Begin VB.TextBox txtISBT128 
      Height          =   285
      Left            =   735
      TabIndex        =   25
      Top             =   15
      Width           =   2010
   End
   Begin VB.CommandButton cmdInterHospital 
      Caption         =   "Inter-Hospital Transfer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   13170
      TabIndex        =   24
      Top             =   7125
      Width           =   1545
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
      Height          =   675
      Left            =   9270
      Picture         =   "frmMovement.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdPendingTransfusion 
      Caption         =   "Remove from Lab Pending Transfusion"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11190
      TabIndex        =   21
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdTag 
      BackColor       =   &H000000FF&
      Caption         =   "Unit Tagged"
      Height          =   495
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   210
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6810
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5745
      Left            =   90
      TabIndex        =   15
      Top             =   1155
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   10134
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmMovement.frx":0BD4
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
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   288
      Left            =   0
      TabIndex        =   14
      Top             =   8208
      Width           =   15756
      _ExtentX        =   27781
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/09/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "17:01"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
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
      Caption         =   "Pack Dispatch"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9210
      TabIndex        =   11
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdTransfuse 
      Caption         =   "Transfuse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4965
      TabIndex        =   10
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7140
      TabIndex        =   9
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdReturnToStock 
      Caption         =   "Return to Stock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2820
      TabIndex        =   8
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdReturnToSupplier 
      Caption         =   "Return to Supplier"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   630
      TabIndex        =   7
      Top             =   7125
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print History"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11715
      Picture         =   "frmMovement.frx":0D4C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   675
      Left            =   13035
      Picture         =   "frmMovement.frx":13B6
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "cmdCancel"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtUnitNumber 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   14460
      MaxLength       =   14
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   75
      TabIndex        =   19
      Top             =   7950
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblVisionGroupCheckStatus 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      Height          =   270
      Left            =   7620
      TabIndex        =   28
      Top             =   390
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblEIsuitable 
      Caption         =   "EI suitable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5955
      TabIndex        =   27
      Top             =   135
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ISBT128"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   26
      Top             =   90
      Width           =   645
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   10515
      TabIndex        =   23
      Top             =   330
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblProduct 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   735
      TabIndex        =   17
      Top             =   705
      Width           =   5310
   End
   Begin VB.Label lChecked 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5955
      TabIndex        =   16
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label lgroup 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3675
      TabIndex        =   13
      Top             =   345
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4005
      TabIndex        =   12
      Top             =   120
      Width           =   525
   End
   Begin VB.Label lblExpiry 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   735
      TabIndex        =   4
      Top             =   360
      Width           =   2010
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   6
      Top             =   405
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   750
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13830
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GroupRh As String

Private Sub DisableButtons()

10    cmdReturnToSupplier.Enabled = False
20    cmdReturnToStock.Enabled = False
30    cmdTransfuse.Enabled = False
40    cmdDestroy.Enabled = False
50    cmdDispatch.Enabled = False
60    cmdPendingTransfusion.Enabled = False
70    cmdInterHospital.Enabled = False

End Sub


Private Sub FillGISBT()

      Dim s As String

10    If Trim$(txtISBT128) = "" Then Exit Sub

20    txtISBT128 = UCase(txtISBT128)

30    If Left$(txtISBT128, 1) = "=" Then    'Barcode scanning entry
40        s = ISOmod37_2(Mid$(txtISBT128, 2, 13))
50        txtISBT128 = Mid$(txtISBT128, 2, 13) & " " & s
60    End If

70    If Me.ActiveControl.Tag = "cmdCancel" Then
80        Unload Me
90    Else
100       FillGridISBT128
110   End If

120   If lblExpiry <> "" Then
130       If TagIsPresent(txtISBT128, CDate(lblExpiry)) Then
140           cmdTag.Visible = True
150       Else
160           cmdTag.Visible = False
170       End If
180   End If

185   DisplayVisionGroupCheckStatus (txtISBT128)
      
190   DisableButtons

End Sub

Private Sub DisplayVisionGroupCheckStatus(txtISBT128)
            Dim tb As Recordset
            Dim sql As String

10    On Error GoTo DisplayVisionGroupCheckStatus_Error

20       sql = "Select * from VisionUnitGroupCheck where " & _
                  "UnitNumber = '" & txtISBT128 & "' and Barcode = '" & ProductBarCodeFor(lblProduct) & "' "

30       Set tb = New Recordset
40       RecOpenClientBB 0, tb, sql
50       If tb.EOF Then
60            lblVisionGroupCheckStatus.Visible = False
80       ElseIf tb!ConfirmedGroup Then
90            lblVisionGroupCheckStatus.Visible = True
100           lblVisionGroupCheckStatus.BackColor = &HFF00&    'Green
110           lblVisionGroupCheckStatus.ForeColor = vbBlack
120           lblVisionGroupCheckStatus.ToolTipText = "Vision: - Group Check confirmed: " & Format(tb!DateTimeOfRecord, "dd/MMM/yyyy hh:mm:ss")
130      Else
140           lblVisionGroupCheckStatus.Visible = True
150           lblVisionGroupCheckStatus.BackColor = vbRed    'Red
160           lblVisionGroupCheckStatus.ForeColor = vbWhite
170           lblVisionGroupCheckStatus.ToolTipText = "Vision: - Group Check failed: " & Format(tb!DateTimeOfRecord, "dd/MMM/yyyy hh:mm:ss")
180      End If

190   Exit Sub

DisplayVisionGroupCheckStatus_Error:

       Dim strES As String
       Dim intEL As Integer

200    intEL = Erl
210    strES = Err.Description
220    LogError "frmMovement", "DisplayVisionGroupCheckStatus", intEL, strES, sql

End Sub


Private Sub FillG()

      Dim Check As String

10    If Trim$(txtUnitNumber) = "" Then Exit Sub
20    If Len(txtUnitNumber) = 9 And UCase$(Left$(txtUnitNumber, 1)) = "D" And UCase$(Right$(txtUnitNumber, 1)) = "D" Then
30        txtUnitNumber = Mid$(txtUnitNumber, 2, 7)
40    End If
50    If Len(txtUnitNumber) = 7 Then
60        txtUnitNumber = UCase$(Replace(txtUnitNumber, "+", "X"))
70        Check = ChkDig(Left$(txtUnitNumber, 6))
80        If Check <> Right$(txtUnitNumber, 1) Then
90            iMsg "Check Digit incorrect!", vbCritical
100           If TimedOut Then Unload Me
110           Exit Sub
120       End If
130   End If

140   If Me.ActiveControl.Tag = "cmdCancel" Then
150       Unload Me
160   Else
170       FillGrid
180   End If

190   If lblExpiry <> "" Then
200       If TagIsPresent(txtUnitNumber, CDate(lblExpiry)) Then
210           cmdTag.Visible = True
220       Else
230           cmdTag.Visible = False
240       End If
250   End If
260   DisableButtons

End Sub

Private Sub cmdDispatch_Click()

      Dim sql As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim Reason As String
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdDispatch_Click_Error

20    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
30    If Ps.Count = 0 Then
40        iMsg "Unit Number not found."
50        If TimedOut Then Unload Me: Exit Sub
60        txtISBT128 = ""
70        Exit Sub
80    ElseIf Ps.Count > 1 Then    'multiple products found
90        Set f = New frmSelectFromMultiple
100       f.ProductList = Ps
110       f.Show 1
120       Set p = f.SelectedProduct
130       Unload f
140       Set f = Nothing
150   Else
160       Set p = Ps.Item(1)
170   End If

180   If p.PackEvent = "Z" Then
190       iMsg "This unit cannot be dispatched." & vbCrLf & "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220   End If

230   Set f = New frmQueryValidate
240   With f
250       .ShowDatePicker = False
260       .Prompt = "Enter Details of Despatch."
270       .Options = New Collection
280       .Options.Add "Unit Transferred to Drogheda"
290       .Options.Add "Inter-Hospital Transfer"
300       .Options.Add "Laboratory Use"
310       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
320       .NewStatus = gEVENTCODES("F").Text
330       .UnitNumber = txtISBT128
340       .Chart = ""
350       .PatientName = ""
360       .Show 1
370       Accepted = .retval
380       Reason = .Reason
390       Comment = .Comment
          ' StartDateTime = .DateTimeReturn
400   End With
410   Set f = Nothing

420   If Accepted Then

430       If Trim$(Reason) = "" Then
440           iMsg "Cancelled", vbInformation
450           If TimedOut Then Unload Me: Exit Sub
460           Exit Sub
470       End If

480       Validate "F", "", "", "", "", "", Comment

490       sql = "Insert into Dispatch " & _
                "(DateTime, Number, Details) VALUES " & _
                "('" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "', " & _
                " '" & Trim$(txtISBT128) & "', " & _
                " '" & Reason & "') "
500       CnxnBB(0).Execute sql

510       FillGISBT

          '*******************************
          'Log unit fate activity in courier interface requests table to update
          'status in blood courier management system by Blood Courier Interface
          'Send FT signal with sample status as U (Unknow- Dispatched)
520       If InStr(1, lblProduct, "Red") <> 0 Or _
             InStr(1, lblProduct, "Platelet") <> 0 Or _
             InStr(1, lblProduct, "Octaplas") <> 0 Then
              Dim MSG As udtRS
530           With MSG
540               .UnitNumber = txtISBT128
550               .ProductCode = ProductBarCodeFor(lblProduct)
560               .UnitExpiryDate = lblExpiry
570               .SampleStatus = "U"
580               .StorageLocation = ""
590               .ActionText = "Dispatch"
600               .UserName = UserName
610           End With
620           LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT
630       End If
640   End If

650   Exit Sub

cmdDispatch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

660   intEL = Erl
670   strES = Err.Description
680   LogError "frmMovement", "cmdDispatch_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdDestroy_Click()

      Dim Reason As String
      Dim f As Form
      Dim sql As String
      Dim Accepted As Boolean
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdDestroy_Click_Error

20    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
30    If Ps.Count = 0 Then
40        iMsg "Unit Number not found."
50        If TimedOut Then Unload Me: Exit Sub
60        txtISBT128 = ""
70        Exit Sub
80    ElseIf Ps.Count > 1 Then    'multiple products found
90        Set f = New frmSelectFromMultiple
100       f.ProductList = Ps
110       f.Show 1
120       Set p = f.SelectedProduct
130       Unload f
140       Set f = Nothing
150   Else
160       Set p = Ps.Item(1)
170   End If

180   If p.PackEvent = "Z" Then
190       iMsg "This unit cannot be destroyed." & vbCrLf & _
               "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220   End If

230   Set f = New frmQueryValidate
240   With f
250       .ShowDatePicker = False
260       .Prompt = "Enter Reason for Destroying."
270       .Options = New Collection
280       .Options.Add "Product Expired"
290       .Options.Add "Out of Fridge > 30 min"
300       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
310       .NewStatus = gEVENTCODES("D").Text
320       .UnitNumber = txtISBT128
330       .Chart = ""
340       .PatientName = ""
350       .Show 1
360       Accepted = .retval
370       Reason = .Reason
380       Comment = .Comment
          ' StartDateTime = .DateTimeReturn
390   End With
400   Set f = Nothing

410   If Accepted Then

420       If Trim$(Reason) = "" Then
430           iMsg "Cancelled", vbInformation
440           If TimedOut Then Unload Me: Exit Sub
450           Exit Sub
460       End If

470       Validate "D", "", "", Reason, "", "", Comment

480       sql = "Insert into Destroy " & _
                "(Unit, Reason, Expiry)  VALUES " & _
                "('" & txtISBT128 & "', " & _
                "'" & AddTicks(Reason) & "', " & _
                "'" & Format$(lblExpiry, "dd/MMM/yyyy HH:mm") & "')"
490       CnxnBB(0).Execute sql

500   End If

510   FillGISBT

      '*******************************
      'Log unit fate activity in courier interface requests table to update
      'status in blood courier management system by Blood Courier Interface
      'Send FT signal with sample status as D (Destroyed)
520   If InStr(1, lblProduct, "Red") <> 0 Or _
         InStr(1, lblProduct, "Platelet") <> 0 Or _
        InStr(1, lblProduct, "Octaplas") <> 0 Then
          Dim MSG As udtRS
530       With MSG
540           .UnitNumber = txtISBT128
550           .ProductCode = ProductBarCodeFor(lblProduct)
560           .UnitExpiryDate = lblExpiry
570           .SampleStatus = "D"
580           .StorageLocation = ""
590           .ActionText = "Destroy"
600           .UserName = UserName
610       End With
620       LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT
630   End If

640   Exit Sub

cmdDestroy_Click_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "frmMovement", "cmdDestroy_Click", intEL, strES, sql

End Sub

Private Sub cmdInterHospital_Click()

      Dim sql As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim Reason As String
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdInterHospital_Click_Error

20    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
30    If Ps.Count = 0 Then
40        iMsg "Unit Number not found."
50        If TimedOut Then Unload Me: Exit Sub
60        txtISBT128 = ""
70        Exit Sub
80    ElseIf Ps.Count > 1 Then    'multiple products found
90        Set f = New frmSelectFromMultiple
100       f.ProductList = Ps
110       f.Show 1
120       Set p = f.SelectedProduct
130       Unload f
140       Set f = Nothing
150   Else
160       Set p = Ps.Item(1)
170   End If

180   If p.PackEvent = "Z" Then
190       iMsg "This unit cannot be dispatched." & vbCrLf & _
               "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
200       If TimedOut Then Unload Me: Exit Sub
210       Exit Sub
220   End If

230   Set f = New frmQueryValidate
240   With f
250       .ShowDatePicker = False
260       .Prompt = "Reason for Transfer"
270       .Options = New Collection
280       .Options.Add "Patient sent to Drogheda."
290       .Options.Add "Unit transferred with Patient."
300       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
310       .NewStatus = gEVENTCODES("N").Text
320       .UnitNumber = txtISBT128
330       .Chart = ""
340       .PatientName = ""
350       .Show 1
360       Accepted = .retval
370       Reason = .Reason
380       Comment = .Comment
          ' StartDateTime = .DateTimeReturn
390   End With
400   Set f = Nothing

410   If Accepted Then

420       If Trim$(Reason) = "" Then
430           iMsg "Cancelled", vbInformation
440           If TimedOut Then Unload Me: Exit Sub
450           Exit Sub
460       End If

470       Validate "N", "", "", "", "", "", Comment


480       sql = "Insert into Dispatch " & _
                "(DateTime, Number, Details) VALUES " & _
                "('" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "', " & _
                " '" & Trim$(txtISBT128) & "', " & _
                " '" & Reason & "') "
490       CnxnBB(0).Execute sql

500       FillGISBT

          '*******************************
          'Log unit fate activity in courier interface requests table to update
          'status in blood courier management system by Blood Courier Interface
          'Send FT signal with sample status as U (Unknow- Dispatched)
510       If InStr(1, lblProduct, "Red") <> 0 Or _
             InStr(1, lblProduct, "Platelet") <> 0 Or _
             InStr(1, lblProduct, "Octaplas") <> 0 Then
              Dim MSG As udtRS
520           With MSG
530               .UnitNumber = txtISBT128
540               .ProductCode = ProductBarCodeFor(lblProduct)
550               .UnitExpiryDate = lblExpiry
560               .SampleStatus = "U"
570               .StorageLocation = ""
580               .ActionText = "Dispatch"
590               .UserName = UserName
600           End With
610           LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT
620       End If
630   End If

640   Exit Sub

cmdInterHospital_Click_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "frmMovement", "cmdInterHospital_Click", intEL, strES, sql

End Sub

Private Sub cmdPendingTransfusion_Click()

      Dim Y As Long
      Dim PatName As String
      Dim Chart As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim StartDateTime As String
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdPendingTransfusion_Click_Error

20    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
30    If Ps.Count = 0 Then
40        iMsg "Unit Number not found."
50        If TimedOut Then Unload Me: Exit Sub
60        txtISBT128 = ""
70        Exit Sub
80    ElseIf Ps.Count > 1 Then    'multiple products found
90        Set f = New frmSelectFromMultiple
100       f.ProductList = Ps
110       f.Show 1
120       Set p = f.SelectedProduct
130       Unload f
140       Set f = Nothing
150   Else
160       Set p = Ps.Item(1)
170   End If

180   For Y = 1 To g.Rows - 1
190       g.row = Y
200       If g.CellBackColor = vbYellow Then
210           PatName = g.TextMatrix(Y, 8)
220           Chart = g.TextMatrix(Y, 6)
230           Exit For
240       End If
250   Next

260   If p.PackEvent = "Z" Then
270       iMsg "This unit cannot be removed from lab." & vbCrLf & _
               "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
280       If TimedOut Then Unload Me: Exit Sub
290       Exit Sub
300   End If

310   If Trim$(Chart) = "" And Trim$(PatName) = "" Then
320       iMsg "Select Patient!", vbCritical
330       If TimedOut Then Unload Me: Exit Sub
340       Exit Sub
350   End If

360   Set f = New frmQueryValidate
370   With f
380       .ShowDatePicker = True
390       .DatePickerCaption = "Date/Time of Removal"
400       .Prompt = ""
410       .Options = New Collection
420       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
430       .NewStatus = gEVENTCODES("Y").Text
440       .UnitNumber = txtISBT128
450       .Chart = Chart
460       .PatientName = PatName
470       .Show 1
480       Accepted = .retval
490       StartDateTime = .DateTimeReturn
500       Comment = .Comment
510   End With
520   Set f = Nothing

530   If Accepted Then
540       Validate "Y", PatName, Chart, "", StartDateTime, "", Comment
550   End If

560   FillGISBT

570   Exit Sub

cmdPendingTransfusion_Click_Error:

      Dim strES As String
      Dim intEL As Integer

580   intEL = Erl
590   strES = Err.Description
600   LogError "frmMovement", "cmdPendingTransfusion_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim Rhesus As String

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    If txtISBT128.Visible Then
40        If Trim$(txtISBT128) = "" Then
50            iMsg "Unit Number?", vbQuestion
60            If TimedOut Then Unload Me: Exit Sub
70            txtUnitNumber.SetFocus
80            Exit Sub
90        End If
100       FillGISBT
110   Else
120       If Trim$(txtUnitNumber) = "" Then
130           iMsg "Unit Number?", vbQuestion
140           If TimedOut Then Unload Me: Exit Sub
150           txtUnitNumber.SetFocus
160           Exit Sub
170       End If
180       FillGrid
190   End If

200   If InStr(UCase$(lgroup), "NEG") Then
210       Rhesus = "Negative"
220   ElseIf InStr(UCase$(lgroup), "POS") Then
230       Rhesus = "Positive"
240   End If

250   Printer.FontName = "Courier New"
260   Printer.Orientation = vbPRORPortrait

270   Printer.Print
280   Printer.FontSize = 10

290   Printer.FontBold = True
300   Printer.Print FormatString("Stock Movement Report", 99, , AlignCenter)
310   Printer.FontBold = False
320   Printer.FontSize = 4
330   Printer.Print String$(248, "-")
340   Printer.Font.Size = 10
350   If txtISBT128.Visible Then    'ISBT128
360       Printer.Print FormatString("Unit: " & txtISBT128 & "    Product: " & lblProduct, 99, , Alignleft)
370   Else
380       Printer.Print FormatString("Unit: " & txtUnitNumber & "    Product: " & lblProduct, 99, , Alignleft)
390   End If
400   Printer.Print FormatString("Expiry: " & lblExpiry & "    Group: " & Left$(lgroup, 2) & "    Rhesus: " & Rhesus, 99, , Alignleft)
410   Printer.FontSize = 4
420   Printer.Print String$(248, "-")
430   Printer.Font.Size = 9
440   Printer.FontBold = True
450   Printer.Print FormatString("", 0, "|", AlignCenter);
460   Printer.Print FormatString("Date", 14, "|", AlignCenter);
470   Printer.Print FormatString("Event", 55, "|", AlignCenter);
480   Printer.Print FormatString("ID", 10, "|", AlignCenter);
490   Printer.Print FormatString("Name", 20, "|", AlignCenter);
500   Printer.Print FormatString("User", 5, "|", AlignCenter)
510   Printer.FontBold = False
520   Printer.FontSize = 4
530   Printer.Print String$(248, "-")
540   Printer.Font.Size = 9

550   For Y = 1 To g.Rows - 1
560       Printer.Print FormatString("", 0, "|", Alignleft);
570       Printer.Print FormatString(g.TextMatrix(Y, 0), 14, "|", Alignleft);  'date
580       Printer.Print FormatString(g.TextMatrix(Y, 1), 55, "|", Alignleft);    'event
590       Printer.Print FormatString(g.TextMatrix(Y, 6), 10, "|", Alignleft);    'ID
600       Printer.Print FormatString(g.TextMatrix(Y, 8), 20, "|", Alignleft);    'Name
610       Printer.Print FormatString(g.TextMatrix(Y, 5), 5, "|", Alignleft)    'User

620   Next
630   Printer.FontSize = 4
640   Printer.Print String$(400, "-")
650   Printer.Font.Size = 9

660   Printer.EndDoc

670   For Each Px In Printers
680       If Px.DeviceName = OriginalPrinter Then
690           Set Printer = Px
700           Exit For
710       End If
720   Next

End Sub

Private Sub cmdReturnToStock_Click()

      Dim sql As String
      Dim DoB As String
      Dim Generic As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdReturnToStock_Click_Error

20    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
30    If Ps.Count > 0 Then
40        Set p = Ps(1)
50        If p.PackEvent = "Z" Then
60            iMsg "This unit cannot be returned to stock." & vbCrLf & _
                   "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
70            If TimedOut Then Unload Me: Exit Sub
80            Exit Sub
90        End If
100   End If

110   If IsDate(g.TextMatrix(1, 9)) Then
120       DoB = Format(g.TextMatrix(1, 9), "dd/mmm/yyyy")
130   Else
140       DoB = ""
150   End If

160   If DateDiff("d", lblExpiry, Now) > 0 Then
170       iMsg "Product has expired and can now only be destroyed or returned to the Supplier.", vbInformation
180       If TimedOut Then Unload Me: Exit Sub
190       Exit Sub
200   End If

210   Set f = New frmQueryValidate
220   With f
230       .ShowDatePicker = False
240       .Prompt = ""
250       .Options = New Collection
          '.Options.Add "Returned for credit"
          '.Options.Add "Returned not for credit"
          '.Options.Add "Returned reason not specified"
260       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
270       .NewStatus = gEVENTCODES("R").Text
280       .UnitNumber = txtISBT128
290       .Chart = ""
300       .PatientName = ""
310       .Show 1
320       Accepted = .retval
330       Comment = .Comment
340   End With
350   Set f = Nothing

360   If Accepted Then
370       Validate "R", "", "", "", "", "", Comment

380       sql = "Insert into Reclaimed " & _
                "( Name, Chart, Unit, [Group], Product, xmDate, " & _
                "  DateTime, Operator, Ward, DoB, Typenex ) VALUES " & _
                "( '" & AddTicks(g.TextMatrix(1, 8)) & "', " & _
                "  '" & g.TextMatrix(1, 6) & "', " & _
                "  '" & txtISBT128 & "', " & _
                "  '" & lgroup & "', " & _
                "  '" & lblProduct & "', " & _
                "  '" & Format(g.TextMatrix(1, 0), "dd/mmm/yyyy") & "', " & _
                "  '" & Format(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
                "  '" & UserCode & "', " & _
                "  '" & g.TextMatrix(1, 9) & "', " & _
                "  '" & DoB & "', " & _
                "  '" & g.TextMatrix(1, 7) & "' )"
390       CnxnBB(0).Execute sql

400       FillGISBT

          'Send RTS signal to Courier
410       If InStr(1, lblProduct, "Red") <> 0 Or _
             InStr(1, lblProduct, "Platelet") <> 0 Or _
             InStr(1, lblProduct, "Octaplas") <> 0 Then
              Dim MSG As udtRS
420           With MSG
430               .UnitNumber = txtISBT128
440               .ProductCode = ProductBarCodeFor(lblProduct)
450               .UnitExpiryDate = lblExpiry
460               .StorageLocation = strBTCourier_StorageLocation_StockFridge
470               .ActionText = "Return to Stock"
480               .UserName = UserName
490           End With
500           LogCourierInterface "RTS", MSG

              'Cavan requested that "RTS" message followed with "SU"
510           With MSG
520               .UnitNumber = txtISBT128
530               .ProductCode = ProductBarCodeFor(lblProduct)
540               Generic = ProductGenericFor(.ProductCode)
550               If Generic = "Platelets" Then
560                   .StorageLocation = strBTCourier_StorageLocation_RoomTempIssueFridge
570               Else
580                   .StorageLocation = strBTCourier_StorageLocation_HemoSafeFridge
590               End If
600               .UnitExpiryDate = Format(lblExpiry, "dd/mmm/yyyy HH:mm")
610               .UnitGroup = lgroup
                  'No real Patient details when unit Restocked
620               .StockComment = ""
630               .Chart = ""
640               .PatientHealthServiceNumber = ""
650               .ForeName = ""
660               .SurName = ""
670               .DoB = ""
680               .Sex = ""
690               .PatientGroup = ""
700               .DeReservationDateTime = Format(lblExpiry, "dd-MMM-yyyy hh:mm:ss")
710               .ActionText = "Stock Update"
720               .UserName = UserName
730           End With
740           LogCourierInterface "SU3", MSG
750       End If
760   End If

770   Exit Sub

cmdReturnToStock_Click_Error:

      Dim strES As String
      Dim intEL As Integer

780   intEL = Erl
790   strES = Err.Description
800   LogError "frmmovement", "cmdReturnToStock_Click", intEL, strES, sql

End Sub

Private Sub cmdTag_Click()

10    If txtISBT128.Visible Then
20        frmUnitNotes.txtUnitNumber = txtISBT128
30        frmUnitNotes.txtUnitNumber.Tag = "ISBT128"
40    Else
50        frmUnitNotes.txtUnitNumber = txtUnitNumber
60        frmUnitNotes.txtUnitNumber.Tag = ""
70    End If
80    frmUnitNotes.txtExpiry = lblExpiry
90    frmUnitNotes.Show 1
100   If txtISBT128.Visible Then
110       FillGISBT
120   Else
130       FillG
140   End If

End Sub

Private Sub cmdTransfuse_Click()

      Dim Y As Long
      Dim PatName As String
      Dim SurName As String
      Dim ForeName As String
      Dim Chart As String
      Dim StartDateTime As String
      Dim sql As String
      Dim tb As Recordset
      Dim Sex As String
      Dim DoB As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdTransfuse_Click_Error

20    For Y = 1 To g.Rows - 1
30        g.row = Y
40        If g.CellBackColor = vbYellow Then
50            PatName = g.TextMatrix(Y, 8)
60            SurName = g.TextMatrix(Y, 14)
70            ForeName = g.TextMatrix(Y, 15)
80            Chart = g.TextMatrix(Y, 6)
90            Exit For
100       End If
110   Next

120   Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
130   If Ps.Count = 0 Then
140       iMsg "Unit Number not found."
150       If TimedOut Then Unload Me: Exit Sub
160       txtISBT128 = ""
170       Exit Sub
180   ElseIf Ps.Count > 1 Then    'multiple products found
190       Set f = New frmSelectFromMultiple
200       f.ProductList = Ps
210       f.Show 1
220       Set p = f.SelectedProduct
230       Unload f
240       Set f = Nothing
250   Else
260       Set p = Ps.Item(1)
270   End If

280   If p.PackEvent = "Z" Then
290       iMsg "This unit cannot be transfused." & vbCrLf & _
               "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
300       If TimedOut Then Unload Me: Exit Sub
310       Exit Sub
320   End If

330   If Trim$(Chart) = "" And Trim$(PatName) = "" Then
340       iMsg "Select Patient!", vbCritical
350       If TimedOut Then Unload Me: Exit Sub
360       Exit Sub
370   End If

380   Set f = New frmQueryValidate
390   With f
400       .ShowDatePicker = True
410       .DatePickerCaption = "Date/Time of Transfusion START"
420       .Prompt = ""
430       .Options = New Collection
440       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
450       If p.PackEvent = "Y" Then
460           .DateTimeRemovedFromLab = p.RecordDateTime
470       Else
480           .DateTimeRemovedFromLab = ""
490       End If
500       .NewStatus = gEVENTCODES("S").Text
510       .UnitNumber = txtISBT128
520       .Chart = Chart
530       .PatientName = PatName
540       .Show 1
550       Accepted = .retval
560       StartDateTime = .DateTimeReturn
570       Comment = .Comment
580   End With
590   Set f = Nothing

600   If Accepted Then
610       Validate "S", PatName, Chart, "", "", StartDateTime, Comment
620       FillGISBT

          '*******************************
          'Log unit fate activity in courier interface requests table to update
          'status in blood courier management system by Blood Courier Interface
          'Implemented Site: CAVAN General Hospital
630       If SurName & ForeName = "" Then
640           Y = InStr(PatName, " ")
650           If Y <> 0 Then
660               SurName = Left$(PatName, Y - 1)
670               ForeName = Mid$(PatName, Y + 1)
680           Else
690               SurName = PatName
700               ForeName = ""
710           End If
                  
720       End If
730       SurName = UCase$(SurName)

740       sql = "SELECT Sex, DoB FROM PatientDetails WHERE " & _
                "PatNum = '" & Chart & "' " & _
                "ORDER BY DateTime DESC"
750       Set tb = New Recordset
760       RecOpenServerBB 0, tb, sql
770       If tb.EOF Then
780           Sex = "U"
790           DoB = ""
800       Else
810           Sex = tb!Sex & ""
820           If Not IsNull(tb!DoB) Then
830               If IsDate(tb!DoB) Then
840                   DoB = Format$(tb!DoB, "dd-MMM-yyyy")
850               Else
860                   DoB = ""
870               End If
880           Else
890               DoB = ""
900           End If
910       End If

          'Send FT signal with sample status as U (Unknow- Returned to supplier)
920       If InStr(1, lblProduct, "Red") <> 0 Or _
             InStr(1, lblProduct, "Platelet") <> 0 Or _
             InStr(1, lblProduct, "Octaplas") <> 0 Then
              Dim MSG As udtRS
930           With MSG
940               .UnitNumber = txtISBT128
950               .ProductCode = ProductBarCodeFor(lblProduct)
960               .UnitExpiryDate = lblExpiry
970               .Chart = g.TextMatrix(g.row, 6)
980               .PatientHealthServiceNumber = ""
990               .SurName = SurName
1000              .ForeName = ForeName
1010              .DoB = DoB
1020              .Sex = Sex
1030              .SampleStatus = "T"
1040              .StorageLocation = ""
1050              .ActionText = "Transfuse"
1060              .UserName = UserName
1070          End With
1080          LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT
1090      End If

1100  End If

1110  Exit Sub

cmdTransfuse_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1120  intEL = Erl
1130  strES = Err.Description
1140  LogError "frmMovement", "cmdTransfuse_Click", intEL, strES

End Sub

Private Sub FillGrid()

      Dim tbProd As Recordset
      Dim tbDem As Recordset
      Dim tbDis As Recordset
      Dim n As Integer
      Dim sql As String
      Dim s As String
      Dim Typenex As String
      Dim DoB As String
      Dim Ward As String
      Dim DispatchDetails As String
      Dim EventDetails As String
      Dim DDFound As Boolean
      Dim PatName As String
      Dim Chart As String
      Dim SurName As String
      Dim ForeName As String
      Dim Transfused As Boolean
      Dim Ps As New Products
      Dim f As Form
      Dim p As Product

10    On Error GoTo FillGrid_Error

20    Ps.LoadLatestByUnitNumber (txtUnitNumber)

30    If Ps.Count = 0 Then
40        iMsg "Unit Number not found."
50        If TimedOut Then Unload Me: Exit Sub
60        txtUnitNumber = ""
70        Exit Sub
80    ElseIf Ps.Count > 1 Then    'multiple products found
90        Set f = New frmSelectFromMultiple
100       f.ProductList = Ps
110       f.Show 1
120       Set p = f.SelectedProduct
130       Unload f
140       Set f = Nothing
150   Else
160       Set p = Ps.Item(1)
170   End If

180   g.Rows = 2
190   g.AddItem ""
200   g.RemoveItem 1

210   If p Is Nothing Then Exit Sub
220   lblProduct = ProductWordingFor(p.BarCode)
230   lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy")

240   If Trim$(lblProduct) = "" Then Exit Sub

250   sql = "Select * from Product where " & _
            "Number = '" & txtUnitNumber & "' " & _
            "and BarCode = '" & p.BarCode & "' " & _
            "and DateExpiry = '" & Format(lblExpiry, "dd/mmm/yyyy") & "' " & _
            "order by Counter desc"

260   Set tbProd = New Recordset
270   RecOpenClientBB 0, tbProd, sql
280   If tbProd.EOF Then
290       Exit Sub
300   End If

310   GroupRh = tbProd!GroupRh & ""
320   lgroup = Bar2Group(GroupRh)

330   Transfused = False
340   Do While Not tbProd.EOF
350       sql = "Select PatNum, Name, PatSurName, PatForeName, Ward, DoB, Typenex from PatientDetails where " & _
                "LabNumber = '" & tbProd!LabNumber & "'"
360       Set tbDem = New Recordset
370       RecOpenServerBB 0, tbDem, sql
380       If tbDem.EOF Then
390           Typenex = ""
400           DoB = ""
410           Ward = ""
420           PatName = ""
430           Chart = ""
440       Else
450           PatName = tbDem!Name & ""
460           SurName = tbDem!PatSurName & ""
470           ForeName = tbDem!PatForeName & ""
480           Chart = tbDem!Patnum & ""
490           Typenex = tbDem!Typenex & ""
500           DoB = tbDem!DoB & ""
510           Ward = tbDem!Ward & ""
520       End If
530       DispatchDetails = ""
540       If tbProd!Event & "" = "S" Then Transfused = True
550       EventDetails = gEVENTCODES(tbProd!Event & "").Text
560       If EventDetails = "Dispatched" Or EventDetails = "Transferred to Drogheda" Then
570           sql = "Select * from Dispatch where " & _
                    "Number = '" & txtUnitNumber & "' AND [DateTime] = '" & Format(tbProd!DateTime, "dd/MMM/yyyy HH:nn:ss") & "'"
580           Set tbDis = New Recordset
590           RecOpenServerBB 0, tbDis, sql
600           If Not tbDis.EOF Then
610               DispatchDetails = tbDis!Details & ""
620           End If
630       End If
640       s = Format(tbProd!DateTime, "dd/MM/yy HH:nn:ss") & vbTab & _
              EventDetails & vbTab
650       If IsDate(tbProd!EventStart & "") Then
660           s = s & Format$(tbProd!EventStart, "dd/MM/yy HH:nn:ss") & vbTab
670       Else
680           s = s & vbTab
690       End If

700       If IsDate(tbProd!EventEnd & "") Then
710           s = s & Format$(tbProd!EventEnd, "dd/MM/yy HH:nn:ss") & vbTab
720       Else
730           s = s & vbTab
740       End If

750       s = s & DispatchDetails & vbTab & _
              tbProd!Operator & vbTab & _
              tbProd!Patid & vbTab & _
              Typenex & vbTab & _
              tbProd!PatName & vbTab & _
              DoB & vbTab & _
              Ward & vbTab

          '860     s = s & vbTab

          's = s & vbTab
760       s = s & tbProd!Notes & vbTab & _
              tbProd!Counter & vbTab & _
              tbProd!ISBT128 & "" & _
              SurName & vbTab & _
              ForeName
770       g.AddItem s
780       If PatName <> tbProd!PatName And tbProd!PatName <> "" And PatName <> "" Then
790           iMsg "Name Conflict Detected." & vbCrLf & "Demographic Name = " & PatName & "." & vbCrLf & _
                   "Product Record = " & tbProd!PatName & ".", , , vbRed, 10
800           If TimedOut Then Unload Me: Exit Sub
810       End If
820       If Chart <> tbProd!Patid And tbProd!Patid <> "" And Chart <> "" Then
830           iMsg "Patient ID Conflict Detected." & vbCrLf & "Demographic ID = " & Chart & "." & vbCrLf & _
                   "Product Record = " & tbProd!Patid & ".", , , vbRed, 10
840           If TimedOut Then Unload Me: Exit Sub
850       End If

860       tbProd.MoveNext
870   Loop

880   tbProd.MoveFirst


890   cmdReturnToSupplier.Enabled = False
900   cmdDestroy.Enabled = False
910   cmdDispatch.Enabled = False
920   cmdInterHospital.Enabled = False
930   cmdTransfuse.Enabled = False
940   cmdReturnToStock.Enabled = False
950   cmdPendingTransfusion.Enabled = False

960   cmdPrint.Enabled = True

970   If Not Transfused Then
980       Select Case tbProd!Event & ""

              Case "C", "R", "P":    'last event was "Received" or "Restocked" or "Pending"
990               cmdReturnToSupplier.Enabled = True
1000              cmdDestroy.Enabled = True
1010              cmdDispatch.Enabled = True
1020              cmdInterHospital.Enabled = True
1030          Case "X":    'last event was "Xmatched"

1040              cmdDestroy.Enabled = True
1050              cmdPendingTransfusion.Enabled = True
1060              cmdTransfuse.Enabled = True
1070              cmdReturnToStock.Enabled = True
1080              cmdDispatch.Enabled = True
1090              cmdInterHospital.Enabled = True

1100          Case "P", "I", "Y", "V":   'last event was "Pending" or "Issued" or Pending transfusion or Electronic Issue
1110              cmdDestroy.Enabled = True
1120              cmdTransfuse.Enabled = True
1130              cmdReturnToStock.Enabled = True
1140              cmdDispatch.Enabled = True
1150              cmdInterHospital.Enabled = True


1160          Case "N":    'Last event was dispatched
1170              cmdReturnToStock.Enabled = True

1180      End Select
1190  End If

1200  If Not tbProd!Checked Then
1210      lChecked = "NO" ' was-- "Group NOT Checked"
1220  Else
1230      lChecked = "YES" ' was -- "Group Checked"
1240  End If

1250  If g.Rows > 2 Then
1260      g.RemoveItem 1
1270  End If


1280  DDFound = False
1290  For n = 1 To g.Rows - 1
1300      If Trim$(g.TextMatrix(n, 4)) <> "" Then
1310          DDFound = True
1320          Exit For
1330      End If
1340  Next
1350  If Not DDFound Then
1360      g.ColWidth(4) = 0
1370  Else
1380      g.ColWidth(4) = 735
1390  End If

1400  Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

1410  intEL = Erl
1420  strES = Err.Description
1430  LogError "frmmovement", "FillGrid", intEL, strES, sql

End Sub


Private Sub FillGridISBT128()

      Dim tbProd As Recordset
      Dim tbDem As Recordset
      Dim tbDis As Recordset
      Dim n As Integer
      Dim sql As String
      Dim s As String
      Dim Typenex As String
      Dim DoB As String
      Dim Ward As String
      Dim DispatchDetails As String
      Dim EventDetails As String
      Dim DDFound As Boolean
      Dim PatName As String
      Dim SurName As String
      Dim ForeName As String
      Dim Chart As String
      Dim Transfused As Boolean
      Dim Ps As New Products
      Dim f As Form
      Dim p As Product


10    Ps.LoadLatestByUnitNumberISBT128 (txtISBT128)

20    If Ps.Count = 0 Then
30        iMsg "Unit Number not found."
40        If TimedOut Then Unload Me: Exit Sub
50        txtISBT128 = ""
60        Exit Sub
70    ElseIf Ps.Count > 1 Then    'multiple products found
80        Set f = New frmSelectFromMultiple
90        f.ProductList = Ps
100       f.Show 1
110       Set p = f.SelectedProduct
120       Unload f
130       Set f = Nothing
140   Else
150       Set p = Ps.Item(1)
160   End If

170   g.Rows = 2
180   g.AddItem ""
190   g.RemoveItem 1

200   If p Is Nothing Then Exit Sub
210   lblProduct = ProductWordingFor(p.BarCode)
220   lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy HH:mm")

230   If Trim$(lblProduct) = "" Then Exit Sub

240   sql = "Select * from Product where " & _
            "ISBT128 = '" & txtISBT128 & "' " & _
            "and BarCode = '" & p.BarCode & "' " & _
            "" & _
            "order by Counter desc"

250   Set tbProd = New Recordset
260   RecOpenClientBB 0, tbProd, sql
270   If tbProd.EOF Then
280       Exit Sub
290   End If

300   GroupRh = tbProd!GroupRh & ""
310   lgroup = Bar2Group(GroupRh)

320   Transfused = False
330   Do While Not tbProd.EOF
340       sql = "Select PatNum, Name, PatSurName, PatForeName, Ward, DoB, Typenex from PatientDetails where " & _
                "LabNumber = '" & tbProd!LabNumber & "'"
350       Set tbDem = New Recordset
360       RecOpenServerBB 0, tbDem, sql
370       If tbDem.EOF Then
380           Typenex = ""
390           DoB = ""
400           Ward = ""
410           PatName = ""
420           Chart = ""
430       Else
440           PatName = tbDem!Name & ""
450           SurName = tbDem!PatSurName & ""
460           ForeName = tbDem!PatForeName & ""
470           Chart = tbDem!Patnum & ""
480           Typenex = tbDem!Typenex & ""
490           DoB = tbDem!DoB & ""
500           Ward = tbDem!Ward & ""
510       End If
520       DispatchDetails = ""
530       If tbProd!Event & "" = "S" Then Transfused = True
540       EventDetails = gEVENTCODES(tbProd!Event & "").Text
550       If EventDetails = "Dispatched" Or EventDetails = "Transferred to Drogheda" Then
560           sql = "Select * from Dispatch where " & _
                    "Number = '" & txtUnitNumber & "' AND [DateTime] = '" & Format(tbProd!DateTime, "dd/MMM/yyyy HH:nn:ss") & "'"
570           Set tbDis = New Recordset
580           RecOpenServerBB 0, tbDis, sql
590           If Not tbDis.EOF Then
600               DispatchDetails = tbDis!Details & ""
610           End If
620       End If
630       s = Format(tbProd!DateTime, "dd/MM/yy HH:nn:ss") & vbTab & _
              EventDetails & vbTab
640       If IsDate(tbProd!EventStart & "") Then
650           s = s & Format$(tbProd!EventStart, "dd/MM/yy HH:nn:ss") & vbTab
660       Else
670           s = s & vbTab
680       End If

690       If IsDate(tbProd!EventEnd & "") Then
700           s = s & Format$(tbProd!EventEnd, "dd/MM/yy HH:nn:ss") & vbTab
710       Else
720           s = s & vbTab
730       End If

740       s = s & DispatchDetails & vbTab & _
              tbProd!Operator & vbTab & _
              tbProd!Patid & vbTab & _
              Typenex & vbTab & _
              tbProd!PatName & vbTab & _
              DoB & vbTab & _
              Ward & vbTab

750       s = s & tbProd!Notes & vbTab & _
              tbProd!Counter & vbTab & _
              tbProd!ISBT128 & "" & vbTab & _
              SurName & vbTab & _
              ForeName
760       g.AddItem s
770       If PatName <> tbProd!PatName And tbProd!PatName <> "" And PatName <> "" Then
780           iMsg "Name Conflict Detected." & vbCrLf & "Demographic Name = " & PatName & "." & vbCrLf & _
                   "Product Record = " & tbProd!PatName & ".", , , vbRed, 10
790           If TimedOut Then Unload Me: Exit Sub
800       End If
810       If Chart <> tbProd!Patid And tbProd!Patid <> "" And Chart <> "" Then
820           iMsg "Patient ID Conflict Detected." & vbCrLf & "Demographic ID = " & Chart & "." & vbCrLf & _
                   "Product Record = " & tbProd!Patid & ".", , , vbRed, 10
830           If TimedOut Then Unload Me: Exit Sub
840       End If

850       tbProd.MoveNext
860   Loop

870   tbProd.MoveFirst


880   cmdReturnToSupplier.Enabled = False
890   cmdDestroy.Enabled = False
900   cmdDispatch.Enabled = False
910   cmdInterHospital.Enabled = False
920   cmdTransfuse.Enabled = False
930   cmdReturnToStock.Enabled = False
940   cmdPendingTransfusion.Enabled = False

950   cmdPrint.Enabled = True

960   If Not Transfused Then
970       Select Case tbProd!Event & ""

              Case "C", "R", "P":    'last event was "Received" or "Restocked" or "Pending"
980               cmdReturnToSupplier.Enabled = True
990               cmdDestroy.Enabled = True
1000              cmdDispatch.Enabled = True
1010              cmdInterHospital.Enabled = True
1020          Case "X":    'last event was "Xmatched"

1030              cmdDestroy.Enabled = True
1040              cmdPendingTransfusion.Enabled = True
1050              cmdTransfuse.Enabled = True
1060              cmdReturnToStock.Enabled = True
1070              cmdDispatch.Enabled = True
1080              cmdInterHospital.Enabled = True

1090          Case "P", "I", "Y", "V":    'last event was "Pending" or "Issued" or Pending transfusion or Electronic Issue
1100              cmdDestroy.Enabled = True
1110              cmdTransfuse.Enabled = True
1120              cmdReturnToStock.Enabled = True
1130              cmdDispatch.Enabled = True
1140              cmdInterHospital.Enabled = True


1150          Case "N":    'Last event was dispatched
1160              cmdReturnToStock.Enabled = True

1170      End Select
1180  End If

1190  If Not tbProd!Checked Then
1200      lChecked = "NO" ' was-- "Group NOT Checked"
1210      cmdLog.Visible = True
1220  Else
1230      lChecked = "YES" ' was -- "Group Checked"
1240      cmdLog.Visible = False
1250  End If

1260  If g.Rows > 2 Then
1270      g.RemoveItem 1
1280  End If


1290  DDFound = False
1300  For n = 1 To g.Rows - 1
1310      If Trim$(g.TextMatrix(n, 4)) <> "" Then
1320          DDFound = True
1330          Exit For
1340      End If
1350  Next
1360  If Not DDFound Then
1370      g.ColWidth(4) = 0
1380  Else
1390      g.ColWidth(4) = 735
1400  End If


End Sub


Private Sub cmdLog_Click()

     Dim sql As String
     
10    On Error GoTo cmdLog_Click_Error

15  If iMsg("Are you sure you want to mark this unit as suitable for Electronic Issue?", vbYesNo) = vbYes Then
 
20        sql = "UPDATE Product SET Checked = '1'  WHERE " & _
                    "ISBT128 = '" & txtISBT128 & "' AND BarCode = '" & ProductBarCodeFor(lblProduct) & "' "
30         CnxnBB(0).Execute sql
40         sql = "UPDATE Latest SET Checked = '1'  WHERE " & _
                    "ISBT128 = '" & txtISBT128 & "' AND BarCode = '" & ProductBarCodeFor(lblProduct) & "' "
50         CnxnBB(0).Execute sql

60    cmdLog.Visible = False
70    lChecked = "YES"  'lChecked = "Group Checked"

80    FillGridISBT128
85  End If

90    Exit Sub

cmdLog_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmmovement", "cmdLog_Click", intEL, strES

End Sub


Private Sub cmdReturnToSupplier_Click()

      Dim f As Form
      Dim Accepted As Boolean
      Dim Reason As String
      Dim Comment As String
      Dim Ps As New Products
      Dim p As Product

10    Ps.LoadLatestISBT128 txtISBT128, ProductBarCodeFor(lblProduct)
20    If Ps.Count = 0 Then
30        iMsg "Unit Number not found."
40        If TimedOut Then Unload Me: Exit Sub
50        txtISBT128 = ""
60        Exit Sub
70    ElseIf Ps.Count > 1 Then    'multiple products found
80        Set f = New frmSelectFromMultiple
90        f.ProductList = Ps
100       f.Show 1
110       Set p = f.SelectedProduct
120       Unload f
130       Set f = Nothing
140   Else
150       Set p = Ps.Item(1)
160   End If

170   If p.PackEvent = "Z" Then
180       iMsg "This unit cannot be returned to supplier." & vbCrLf & _
               "Unit current status : " & gEVENTCODES(p.PackEvent).Text, vbInformation
190       If TimedOut Then Unload Me: Exit Sub
200       Exit Sub
210   End If

220   Set f = New frmQueryValidate
230   With f
240       .ShowDatePicker = False
250       .Prompt = "Enter reason for return."
260       .Options = New Collection
270       .Options.Add "Returned for credit"
280       .Options.Add "Returned not for credit"
290       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
300       .NewStatus = gEVENTCODES("T").Text
310       .UnitNumber = txtISBT128
320       .Chart = ""
330       .PatientName = ""
340       .Show 1
350       Accepted = .retval
360       Reason = .Reason
370       Comment = .Comment
380   End With
390   Set f = Nothing

400   If Accepted = True Then
410       LogReasonWhy "(" & txtISBT128 & ") " & Reason, "M"
420       Validate "T", "", "", "", "", "", Comment
430       FillGISBT

          'Send FT signal to Courier with sample status as U (Unknown - Returned to supplier)
440       If InStr(1, lblProduct, "Red") <> 0 Or _
             InStr(1, lblProduct, "Platelet") <> 0 Or _
             InStr(1, lblProduct, "Octaplas") <> 0 Then
              Dim MSG As udtRS
450           With MSG
460               .UnitNumber = txtISBT128
470               .ProductCode = ProductBarCodeFor(lblProduct)
480               .UnitExpiryDate = lblExpiry
490               .SampleStatus = "US"
500               .StorageLocation = ""
510               .ActionText = "Return To Supplier"
520               .UserName = UserName
530           End With
540           LogCourierInterface "FT", MSG    'PUT IN LABNUMBER IF YOU HAVE IT
550       End If
560   End If

End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String
      Dim Rhesus As String

10    If InStr(UCase$(lgroup), "NEG") Then
20        Rhesus = "Negative"
30    ElseIf InStr(UCase$(lgroup), "POS") Then
40        Rhesus = "Positive"
50    End If
60    strHeading = "Stock Movement Report" & vbCr
70    strHeading = strHeading & "Unit: " & txtUnitNumber & "    Product: " & lblProduct & vbCr
80    strHeading = strHeading & "Expiry: " & lblExpiry & "    Group: " & Left$(lgroup, 2) & "    Rhesus: " & Rhesus & vbCr
90    strHeading = strHeading & vbCr

100   ExportFlexGrid g, Me, strHeading

End Sub

Private Sub Form_Load()

5     frmMain.timAnalyserHeartBeat.Enabled = False

10    StatusBar.Panels(3) = UserName

20    g.ColWidth(3) = 0    'End dateTime
30    g.ColWidth(7) = 0    'Typenex
40    g.ColWidth(12) = 0    'Counter
50    g.ColWidth(14) = 0    'Surname
60    g.ColWidth(15) = 0    'Forename
70    g.ColWidth(2) = 1300

80    g.ColAlignment(2) = flexAlignCenterCenter
90    g.RowHeight(0) = 600
100   g.WordWrap = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
10    frmMain.timAnalyserHeartBeat.Enabled = True
End Sub

Private Sub g_Click()

10    If txtISBT128.Visible Then
20        gridClickISBT128
30    Else    'Codabar
40        gridClickCodabar
50    End If

End Sub

Private Sub gridClickCodabar()

      Dim sql As String
      Dim f As Form

10    On Error GoTo gridClickCodabar_Error

20    If g.MouseRow = 0 Then Exit Sub

30    If g.col = 4 Then
40        If g.ColWidth(4) = 4005 Then
50            g.ColWidth(4) = TextWidth("Details      ")
60        Else
70            g.ColWidth(4) = 4005
80        End If
90    End If

100   If g.col = 11 Then
110       Set f = New frmRemarks
120       f.Comment = g.TextMatrix(g.row, 11)
130       f.Heading = "Notes Entry"
140       f.Show 1
150       Unload f
160       Set f = Nothing
170   End If

180   DisableButtons

190   Exit Sub

gridClickCodabar_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmMovement", "gridClickCodabar", intEL, strES, sql

End Sub

Private Sub gridClickISBT128()
      Dim Y As Integer
      Dim X As Integer
      Dim ySave As Integer
      Dim Fated As Boolean
      Dim Notes As String
      Dim sql As String
      Dim f As Form

10    On Error GoTo gridClickISBT128_Error

20    If g.MouseRow = 0 Then Exit Sub

30    If g.col = 4 Then
40        If g.ColWidth(4) = 4005 Then
50            g.ColWidth(4) = TextWidth("Details      ")
60        Else
70            g.ColWidth(4) = 4005
80        End If
90    End If

100   If g.col = 11 Then

110       Set f = New frmRemarks
120       f.Comment = g.TextMatrix(g.row, 11)
130       f.Heading = "Notes Entry"
140       f.Show 1
150       Notes = f.Comment
160       Unload f
170       Set f = Nothing

180       If g.TextMatrix(g.row, 11) <> Notes Then
190           g.TextMatrix(g.row, 11) = Notes
200           Notes = Left$(AddTicks(Notes), 1000)
210           sql = "UPDATE Product SET Notes = '" & Notes & "' WHERE " & _
                    "Counter = '" & g.TextMatrix(g.row, 12) & "'"
220           CnxnBB(0).Execute sql
230           sql = "UPDATE Latest SET Notes = '" & Notes & "' WHERE " & _
                    "ISBT128 = '" & txtISBT128 & "' " & _
                    "AND BarCode = '" & ProductBarCodeFor(lblProduct) & "' " & _
                    "AND DateTime = '" & Format$(g.TextMatrix(g.row, 0), "yyyyMMdd HH:nn") & "' " & _
                    "AND Event = '" & gEVENTCODES.CodeFor(g.TextMatrix(g.row, 1)) & "' "
240           CnxnBB(0).Execute sql
250       End If
260   End If

270   ySave = g.row
280   g.col = 0
290   For Y = 1 To g.Rows - 1
300       g.row = Y
310       If g.CellBackColor = vbYellow Then
320           For X = 0 To g.Cols - 1
330               g.col = X
340               g.CellBackColor = 0
350           Next
360       End If
370   Next
380   g.row = ySave
390   For X = 0 To g.Cols - 1
400       g.col = X
410       g.CellBackColor = vbYellow
420   Next

430   DisableButtons

440   Fated = False
450   If g.TextMatrix(1, 1) = "Destroyed" Then
460       Fated = True
470   ElseIf g.TextMatrix(1, 1) = "Returned to Supplier" Then
480       Fated = True
490   ElseIf g.TextMatrix(1, 1) = "Transfused" Then
500       Fated = True
510   End If
520   If Fated Then Exit Sub

530   Select Case g.TextMatrix(g.row, 1)
          Case "Received into Stock", "Restocked"
540           cmdReturnToSupplier.Enabled = True
550           cmdDestroy.Enabled = True
560           cmdPendingTransfusion.Enabled = True
570           cmdDispatch.Enabled = True
580           cmdInterHospital.Enabled = True
590       Case "Cross matched", "Issued"
600           cmdDispatch.Enabled = True
610           cmdInterHospital.Enabled = True
620           cmdTransfuse.Enabled = True
630           cmdReturnToStock.Enabled = True
640           cmdDestroy.Enabled = True
650           cmdPendingTransfusion.Enabled = True
660       Case "Pending"
670           cmdReturnToStock.Enabled = True
680           cmdDestroy.Enabled = True
690       Case "Removed Pending Transfusion"
700           cmdDestroy.Enabled = True
710           cmdDispatch.Enabled = True
720           cmdInterHospital.Enabled = True
730           cmdTransfuse.Enabled = True
740           cmdReturnToStock.Enabled = True
750       Case "Unit Dispatched"
760       Case "Transferred"
770           cmdDestroy.Enabled = True
780           cmdReturnToStock.Enabled = True
790           cmdTransfuse.Enabled = True
800           cmdDispatch.Enabled = True
801       Case "Blocked - Group Check failed"
802           cmdDestroy.Enabled = True
803           cmdReturnToStock.Enabled = True
810   End Select

820   Exit Sub

gridClickISBT128_Error:

      Dim strES As String
      Dim intEL As Integer

830   intEL = Erl
840   strES = Err.Description
850   LogError "frmMovement", "gridClickISBT128", intEL, strES, sql

End Sub

Private Sub Validate(ByVal EventCode As String, _
                     ByVal PatName As String, _
                     ByVal Chart As String, _
                     ByVal Reason As String, _
                     ByVal EndDateTime As String, _
                     ByVal StartDateTime As String, _
                     ByVal Comment As String)

      Dim ComponentCode As String
      Dim TimeNow As String
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo Validate_Error

20    ComponentCode = ProductBarCodeFor(lblProduct)
30    TimeNow = Format(Now, "dd/mmm/yyyy hh:mm:ss")

40    Ps.LoadLatestISBT128 txtISBT128, ComponentCode
50    If Ps.Count > 0 Then
60        Set p = Ps.Item(1)
70        p.RecordDateTime = TimeNow
80        If EndDateTime = "" Then
90            p.EventEnd = ""
100       Else
110           p.EventEnd = Format(EndDateTime, "dd/mmm/yyyy hh:mm:ss")
120       End If
130       If StartDateTime <> "" Then
140           p.EventStart = Format(StartDateTime, "dd/mmm/yyyy hh:mm:ss")
150       End If
160       p.PackEvent = EventCode
161       If UCase(EventCode) = "R" Or UCase(EventCode) = "D" Or UCase(EventCode) = "T" Then 'R - Restocked, D -Destroyed, T - Return to Supplier
162            p.cco = False
163            p.ccor = False
164            p.cen = False
165            p.cenr = False
166            p.crt = False
167            p.crtr = False
168       End If
170       p.Chart = Chart
180       p.PatName = PatName
190       p.UserName = UserCode
200       p.Reason = Reason
210       p.Notes = Trim$(Comment)
220       p.Save
230   End If

240   Exit Sub

Validate_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmmovement", "Validate", intEL, strES

End Sub


Private Sub txtISBT128_LostFocus()

10    FillGISBT

End Sub

Private Sub txtUnitNumber_LostFocus()

10    txtUnitNumber = Replace(txtUnitNumber, "'", "")

20    FillG

End Sub

