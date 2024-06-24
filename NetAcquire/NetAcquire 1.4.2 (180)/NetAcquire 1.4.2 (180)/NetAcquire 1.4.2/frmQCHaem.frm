VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQCHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Haematology Controls"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   645
      Left            =   8370
      Picture         =   "frmQCHaem.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6210
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   11130
      Picture         =   "frmQCHaem.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6210
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid gControl 
      Height          =   5385
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9499
      _Version        =   393216
      Rows            =   21
      FixedRows       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
   End
   Begin VB.ComboBox cmbLotNumber 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "cmbLotNumber"
      Top             =   300
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "frmQCHaem.frx":0974
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Parameter to show graph"
      Height          =   255
      Left            =   810
      TabIndex        =   6
      Top             =   6300
      Width           =   2715
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
      Left            =   9660
      TabIndex        =   5
      Top             =   6330
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lot Number"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmQCHaem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CleargControl()

      Dim Y As Integer

40270 With gControl
40280   .Cols = 2
40290   For Y = 0 To .Rows - 1
40300     .TextMatrix(Y, 1) = ""
40310   Next
40320   .TextMatrix(0, 0) = "Run Date"
40330   .TextMatrix(1, 0) = "Run Time"
40340   .TextMatrix(2, 0) = "RBC"
40350   .TextMatrix(3, 0) = "WBC"
40360   .TextMatrix(4, 0) = "Hgb"
40370   .TextMatrix(5, 0) = "MCV"
40380   .TextMatrix(6, 0) = "Hct"
40390   .TextMatrix(7, 0) = "MCH"
40400   .TextMatrix(8, 0) = "MCHC"
40410   .TextMatrix(9, 0) = "RDWCV"
40420   .TextMatrix(10, 0) = "Plt"
40430   .TextMatrix(11, 0) = "MPV"
40440   .TextMatrix(12, 0) = "Plcr"
40450   .TextMatrix(13, 0) = "PDW"
40460   .TextMatrix(14, 0) = "RDWCV"
40470   .TextMatrix(15, 0) = "LymA"
40480   .TextMatrix(16, 0) = "LymP"
40490   .TextMatrix(17, 0) = "MonoA"
40500   .TextMatrix(18, 0) = "MonoP"
40510   .TextMatrix(19, 0) = "NeutA"
40520   .TextMatrix(20, 0) = "NeutP"
40530 End With

End Sub

Private Sub FillLotNumbers()

      Dim sql As String
      Dim tb As Recordset

40540 On Error GoTo FillLotNumbers_Error

40550 cmbLotNumber.Clear

40560 sql = "SELECT SampleID FROM HaemControls " & _
            "GROUP BY SampleID " & _
            "ORDER BY max(RunDate) DESC"
40570 Set tb = New Recordset
40580 RecOpenServer 0, tb, sql
40590 Do While Not tb.EOF
40600   cmbLotNumber.AddItem tb!SampleID
40610   tb.MoveNext
40620 Loop

40630 Exit Sub

FillLotNumbers_Error:

      Dim strES As String
      Dim intEL As Integer

40640 intEL = Erl
40650 strES = Err.Description
40660 LogError "frmQCHaem", "FillLotNumbers", intEL, strES, sql


End Sub

Private Sub cmbLotNumber_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim ThisCol As Integer
      Dim n As Integer

40670 On Error GoTo cmbLotNumber_Click_Error

40680 CleargControl
        
40690 sql = "SELECT * FROM HaemControls WHERE " & _
            "SampleID = '" & cmbLotNumber & "' " & _
            "ORDER BY RunDateTime Desc"
40700 Set tb = New ADODB.Recordset
40710 RecOpenClient 0, tb, sql
40720 If tb.EOF Then
40730   iMsg "No QC information found for that Lot Number", vbExclamation
40740   Exit Sub
40750 End If

40760 gControl.Cols = tb.RecordCount + 1
40770 For n = 1 To gControl.Cols - 1
40780   gControl.ColWidth(n) = 850
40790 Next

40800 ThisCol = 0
40810 Do While Not tb.EOF
40820   ThisCol = ThisCol + 1
40830   gControl.TextMatrix(0, ThisCol) = Format$(tb!Rundate, "dd/MM/yy")
40840   gControl.TextMatrix(1, ThisCol) = Format$(tb!RunDateTime, "HH:nn")
40850   gControl.TextMatrix(2, ThisCol) = Format$(tb!rbc & "", "0.00")
40860   gControl.TextMatrix(3, ThisCol) = Format$(tb!WBC & "", "0.0")
40870   gControl.TextMatrix(4, ThisCol) = Format$(tb!Hgb & "", "0.0")
40880   gControl.TextMatrix(5, ThisCol) = Format$(tb!MCV & "", "0.0")
40890   gControl.TextMatrix(6, ThisCol) = Format$(tb!hct & "", "0.0")
40900   gControl.TextMatrix(7, ThisCol) = Format$(tb!mch & "", "0.0")
40910   gControl.TextMatrix(8, ThisCol) = Format$(tb!mchc & "", "0.0")
40920   gControl.TextMatrix(9, ThisCol) = Format$(tb!RDWCV & "", "0.0")
40930   gControl.TextMatrix(10, ThisCol) = Format$(tb!plt & "", "0")
40940   gControl.TextMatrix(11, ThisCol) = Format$(tb!mpv & "", "0.0")
40950   gControl.TextMatrix(12, ThisCol) = Format$(tb!plcr & "", "0.0")
40960   gControl.TextMatrix(13, ThisCol) = Format$(tb!pdw & "", "0.0")
40970   gControl.TextMatrix(14, ThisCol) = Format$(tb!RDWCV & "", "0.0")
40980   gControl.TextMatrix(15, ThisCol) = Format$(tb!LymA & "", "0.0")
40990   gControl.TextMatrix(16, ThisCol) = Format$(tb!LymP & "", "0.0")
41000   gControl.TextMatrix(17, ThisCol) = Format$(tb!MonoA & "", "0.0")
41010   gControl.TextMatrix(18, ThisCol) = Format$(tb!MonoP & "", "0.0")
41020   gControl.TextMatrix(19, ThisCol) = Format$(tb!NeutA & "", "0.0")
41030   gControl.TextMatrix(20, ThisCol) = Format$(tb!NeutP & "", "0.0")
        
41040   tb.MoveNext
41050 Loop

41060 Exit Sub

cmbLotNumber_Click_Error:

      Dim strES As String
      Dim intEL As Integer

41070 intEL = Erl
41080 strES = Err.Description
41090 LogError "frmQCHaem", "cmbLotNumber_Click", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

41100 Unload Me

End Sub


Private Sub cmdXL_Click()

41110 ExportFlexGrid gControl, Me

End Sub


Private Sub Form_Load()
        
41120 CleargControl
41130 FillLotNumbers

End Sub

Private Sub gControl_Click()

      Dim f As Form

41140 With gControl
41150   If .Cols < 3 Then Exit Sub
41160   If .MouseRow > 1 Then
          
41170     Set f = frmQCHaemGraph
          
41180     .row = 0
41190     .RowSel = 0
41200     .Col = 0
41210     .ColSel = .Cols - 1
41220     f.RunDates = .Clip
          
41230     .row = 1
41240     .RowSel = 1
41250     .Col = 0
41260     .ColSel = .Cols - 1
41270     f.RunTimes = .Clip
          
41280     .row = .MouseRow
41290     .RowSel = 1
41300     .Col = 0
41310     .ColSel = .Cols - 1
41320     f.Values = .Clip
          
41330     f.LotNumber = cmbLotNumber
          
41340     f.Show 1
41350     Set f = Nothing
          
41360   End If
41370 End With

End Sub


