VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCopyTo 
   Caption         =   "NetAcquire"
   ClientHeight    =   4620
   ClientLeft      =   75
   ClientTop       =   1275
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   9300
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      Left            =   1380
      TabIndex        =   15
      Top             =   960
      Width           =   1875
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   7080
      TabIndex        =   14
      Text            =   "cmbPrinter"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   7500
      Picture         =   "frmCopyTo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3510
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   6150
      Picture         =   "frmCopyTo.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3510
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   2
      Left            =   4440
      Picture         =   "frmCopyTo.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   1
      Left            =   2640
      Picture         =   "frmCopyTo.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   0
      Left            =   840
      Picture         =   "frmCopyTo.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   825
   End
   Begin VB.ComboBox cmbGP 
      Height          =   315
      Left            =   4050
      TabIndex        =   1
      Text            =   "cmbGP"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   330
      TabIndex        =   2
      Text            =   "cmbWard"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbClinician 
      Height          =   315
      Left            =   2190
      TabIndex        =   0
      Text            =   "cmbClinician"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   1965
      Left            =   300
      TabIndex        =   3
      Top             =   1260
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   5
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
      ScrollBars      =   2
      FormatString    =   $"frmCopyTo.frx":199A
   End
   Begin VB.Label lblSampleID 
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
      Left            =   6630
      TabIndex        =   12
      Top             =   510
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   6960
      TabIndex        =   11
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Send Copy To"
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label lblOriginal 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   510
      Width           =   5535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Send Original To"
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   300
      Width           =   1185
   End
End
Attribute VB_Name = "frmCopyTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private DataChanged As Boolean

Private pDept As String

Private Sub AdjustBlanks()

          Dim Y As Long
          Dim X As Long
          Dim IsBlank As Boolean

          'remove all blank lines
24400     For Y = g.Rows - 1 To 1 Step -1
24410         IsBlank = True
24420         For X = 0 To 2
24430             If g.TextMatrix(Y, X) <> "" Then
24440                 IsBlank = False
24450                 Exit For
24460             End If
24470         Next
24480         If IsBlank Then
24490             If g.Rows > 2 Then
24500                 g.RemoveItem Y
24510             Else
24520                 g.AddItem ""
24530                 g.RemoveItem 1
24540             End If
24550         End If
24560     Next

          'add a blank line to the bottom
24570     Y = g.Rows - 1
24580     IsBlank = True
24590     For X = 0 To 2
24600         If g.TextMatrix(Y, X) <> "" Then
24610             IsBlank = False
24620         End If
24630     Next
24640     If Not IsBlank Then
24650         g.AddItem ""
24660     End If

End Sub

Private Sub FillWardList(ByVal HospitalName As String)

          Dim tb As Recordset
          Dim sql As String
          Dim strHospitalCode As String

24670     On Error GoTo FillWardList_Error

24680     strHospitalCode = ListCodeFor("HO", HospitalName)

          '30    If HospName(0) = "Monaghan" Then
          '40      sql = "Select * from Wards where " & _
          '              "InUse = 1 " & _
          '              "Order by ListOrder"
          '50    Else
24690     sql = "Select * from Wards where " & _
              "HospitalCode = '" & strHospitalCode & "' " & _
              "and InUse = 1 " & _
              "Order by ListOrder"
          '70    End If
24700     Set tb = New Recordset
24710     RecOpenServer 0, tb, sql

24720     With cmbWard
24730         .Clear
24740         Do While Not tb.EOF
24750             If .ListCount > 0 Then
24760                 If Trim$(UCase$(.List(.ListCount - 1))) <> Trim$(UCase$(tb!Text & "")) Then
24770                     .AddItem tb!Text & ""
24780                 End If
24790             Else
24800                 .AddItem tb!Text & ""
24810             End If
24820             tb.MoveNext
24830         Loop
24840     End With

24850     Exit Sub

FillWardList_Error:

          Dim strES As String
          Dim intEL As Integer

24860     intEL = Erl
24870     strES = Err.Description
24880     LogError "frmCopyTo", "FillWardList", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

24890     Unload Me

End Sub

Private Sub cmbClinician_Click()

24900     DataChanged = True

24910     g.SetFocus

End Sub

Private Sub cmbGP_LostFocus()

          Dim Gx As New GP
          Dim strOrig As String

24920     On Error GoTo cmbGP_LostFocus_Error

24930     strOrig = cmbGP

24940     cmbGP = ""

24950     Gx.LoadCodeOrText strOrig
24960     cmbGP = Gx.Text
24970     If sysOptAllowGPFreeText(0) And cmbGP = "" Then
24980         cmbGP = strOrig
24990     End If

25000     Exit Sub
cmbGP_LostFocus_Error:

25010     LogError "frmCopyTo", "cmbGP_LostFocus", Erl, Err.Description
End Sub



Private Sub cmbHospital_Click()

25020     FillLists

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

25030     KeyAscii = 0

End Sub


Private Sub cmbPrinter_Click()

25040     DataChanged = True

25050     g.SetFocus

End Sub


Private Sub cmbGP_Click()

25060     DataChanged = True

25070     g.SetFocus

End Sub


Private Sub cmbWard_Click()

25080     DataChanged = True

25090     g.SetFocus

End Sub


Private Sub cmdClear_Click(Index As Integer)

          Dim Y As Long

25100     For Y = 1 To g.Rows - 1
25110         g.TextMatrix(Y, Index) = ""
25120     Next

25130     Select Case Index
              Case 0: cmbWard.Visible = False
25140         Case 1: cmbClinician.Visible = False
25150         Case 2: cmbGP.Visible = False
25160     End Select

25170     AdjustBlanks

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Long
          Dim SID As Long

25180     On Error GoTo cmdSave_Click_Error

25190     SID = Val(lblSampleID)
25200     If pDept = "M" Then
25210         SID = SID ' + sysOptMicroOffset(0)
25220     End If
25230     If pDept = "Z" Then
25240         SID = SID ' + sysOptSemenOffset(0)
25250     End If

25260     sql = "Delete from SendCopyTo where " & _
              "SampleID = " & SID
25270     Cnxn(0).Execute sql

25280     For Y = 1 To g.Rows - 2
25290         If g.TextMatrix(Y, 4) = "Use Default" Then
25300             g.TextMatrix(Y, 4) = ""
25310         End If
25320         sql = "Insert into SendCopyTo " & _
                  "(SampleID, Ward, Clinician, GP, Device, Destination) VALUES " & _
                  "('" & SID & "', " & _
                  " '" & AddTicks(g.TextMatrix(Y, 0)) & "', " & _
                  " '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                  " '" & AddTicks(g.TextMatrix(Y, 2)) & "', " & _
                  " '" & g.TextMatrix(Y, 3) & "', " & _
                  " '" & IIf(g.TextMatrix(Y, 4) = "", "", g.TextMatrix(Y, 4)) & "')"

25330         Cnxn(0).Execute sql
25340     Next

25350     Unload Me

25360     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

25370     intEL = Erl
25380     strES = Err.Description
25390     LogError "frmCopyTo", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String
          Dim SID As Long

25400     On Error GoTo Form_Activate_Error

25410     g.Rows = 2
25420     g.AddItem ""
25430     g.RemoveItem 1

25440     SID = Val(lblSampleID)
25450     If pDept = "M" Then
25460         SID = SID ' + sysOptMicroOffset(0)
25470     End If

25480     sql = "Select * from SendCopyTo where " & _
              "SampleID = " & SID
25490     Set tb = New Recordset
25500     RecOpenServer 0, tb, sql
25510     Do While Not tb.EOF
25520         s = tb!Ward & vbTab & _
                  tb!Clinician & vbTab & _
                  tb!GP & vbTab & _
                  tb!Device & vbTab & _
                  tb!Destination & ""
25530         g.AddItem s
25540         tb.MoveNext
25550     Loop

25560     If g.Rows > 2 Then
25570         g.RemoveItem 1
25580     End If

25590     Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

25600     intEL = Erl
25610     strES = Err.Description
25620     LogError "frmCopyTo", "Form_Activate", intEL, strES, sql


End Sub

Private Sub Form_Load()

25630     cmbHospital.Clear
25640     cmbHospital.AddItem "Cavan"
          '30    cmbHospital.AddItem "Monaghan"
25650     cmbHospital = HospName(0)

25660     FillLists

End Sub

Private Sub FillLists()

25670     FillWardList cmbHospital
25680     FillClinicians cmbClinician, cmbHospital
25690     FillGPs cmbGP, cmbHospital

25700     cmbWard.AddItem "", 0
25710     cmbClinician.AddItem "", 0
25720     cmbGP.AddItem "", 0

25730     FillPrinterList

End Sub

Private Sub g_Click()

          Dim Found As Boolean
          Dim FaxNumber As String
          Dim WCG As String

25740     If g.MouseRow = 0 Then Exit Sub

25750     Select Case g.Col
              Case 0:
25760             cmbWard.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
25770             cmbWard.Visible = True
25780             cmbWard.SetFocus
25790         Case 1:
25800             cmbClinician.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
25810             cmbClinician.Visible = True
25820             cmbClinician.SetFocus
25830         Case 2:
25840             cmbGP.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
25850             cmbGP.Visible = True
25860             cmbGP.SetFocus
25870         Case 3:
25880             Found = False
25890             FaxNumber = ""
25900             WCG = ""
25910             If g.TextMatrix(g.row, 0) <> "" Then
25920                 Found = True
25930                 FaxNumber = IsFaxable("Wards", g.TextMatrix(g.row, 0))
25940                 If FaxNumber <> "" Then
25950                     WCG = "W"
25960                 End If
25970             End If
25980             If g.TextMatrix(g.row, 1) <> "" Then
25990                 Found = True
26000             End If
26010             If g.TextMatrix(g.row, 2) <> "" Then
26020                 Found = True
                      Dim Gx As New GP
26030                 Gx.LoadName g.TextMatrix(g.row, 2)
26040                 If Gx.FAX <> "" Then
26050                     FaxNumber = Gx.FAX
26060                     WCG = "G"
26070                 End If
26080             End If
          
26090             If Found Then
26100                 If g.TextMatrix(g.row, 3) = "Printer" Then
26110                     If FaxNumber <> "" Then
26120                         g.TextMatrix(g.row, 3) = "FAX"
26130                         g.TextMatrix(g.row, 4) = FaxNumber
26140                     End If
26150                 Else
26160                     g.TextMatrix(g.row, 3) = "Printer"
26170                     g.TextMatrix(g.row, 4) = ""
26180                 End If
26190             End If
        
26200         Case 4:
26210             If g.TextMatrix(g.row, 3) = "Printer" Then
26220                 cmbPrinter.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
26230                 cmbPrinter.Visible = True
26240                 cmbPrinter.SetFocus
26250             End If
26260     End Select

26270     DataChanged = True

End Sub

Private Sub FillPrinterList()

          Dim tb As Recordset
          Dim sql As String

26280     On Error GoTo FillPrinterList_Error

26290     cmbPrinter.Clear
        
26300     cmbPrinter.AddItem "Use Default"
26310     cmbPrinter.AddItem ""

26320     sql = "Select * from InstalledPrinters"
26330     Set tb = New Recordset
26340     RecOpenServer 0, tb, sql

26350     Do While Not tb.EOF
26360         cmbPrinter.AddItem tb!PrinterName & ""
26370         tb.MoveNext
26380     Loop

26390     Exit Sub

FillPrinterList_Error:

          Dim strES As String
          Dim intEL As Integer

26400     intEL = Erl
26410     strES = Err.Description
26420     LogError "frmCopyTo", "FillPrinterList", intEL, strES, sql

End Sub

Private Sub g_GotFocus()

26430     Select Case g.Col
              Case 0:
26440             If cmbWard.Visible Then
26450                 g = cmbWard
26460                 cmbWard.Visible = False
26470                 g.TextMatrix(g.row, 1) = ""
26480                 g.TextMatrix(g.row, 2) = ""
26490             End If
26500         Case 1:
26510             If cmbClinician.Visible Then
26520                 g = cmbClinician
26530                 cmbClinician.Visible = False
26540                 g.TextMatrix(g.row, 0) = ""
26550                 g.TextMatrix(g.row, 2) = ""
26560             End If
26570         Case 2:
26580             If cmbGP.Visible Then
26590                 g = cmbGP
26600                 cmbGP.Visible = False
26610                 g.TextMatrix(g.row, 0) = ""
26620                 g.TextMatrix(g.row, 1) = ""
26630             End If
26640         Case 4:
26650             If cmbPrinter.Visible Then
26660                 g = cmbPrinter
26670                 cmbPrinter.Visible = False
26680             End If
26690     End Select

26700     If g.TextMatrix(g.row, 3) = "" Then
26710         g.TextMatrix(g.row, 3) = "Printer"
26720     End If

26730     cmbWard.Visible = False
26740     cmbClinician.Visible = False
26750     cmbGP.Visible = False
26760     cmbPrinter.Visible = False

26770     AdjustBlanks

End Sub


Public Property Let Dept(ByVal strNewValue As String)

26780     pDept = strNewValue

End Property

