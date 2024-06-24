VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExternalTests 
   Caption         =   "NetAcquire - Test List"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   15330
   ForeColor       =   &H80000008&
   Icon            =   "frmExternalTests.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   15330
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   14010
      TabIndex        =   30
      Top             =   6210
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1095
      Left            =   14130
      Picture         =   "frmExternalTests.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   14130
      Picture         =   "frmExternalTests.frx":5C1E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5100
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1095
      Left            =   14130
      Picture         =   "frmExternalTests.frx":6AE8
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6930
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   1110
      TabIndex        =   17
      Top             =   120
      Width           =   11805
      Begin VB.TextBox txtBiomnisCode 
         Height          =   315
         Left            =   4380
         TabIndex        =   2
         Top             =   720
         Width           =   1605
      End
      Begin VB.ComboBox cmbDepartment 
         Height          =   315
         Left            =   4380
         TabIndex        =   4
         Text            =   "cmbDepartment"
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1950
         Width           =   4635
      End
      Begin VB.TextBox txtUnits 
         Height          =   285
         Left            =   1590
         TabIndex        =   5
         Top             =   1530
         Width           =   1815
      End
      Begin VB.TextBox txtMBCode 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   720
         Width           =   1605
      End
      Begin VB.Frame Frame3 
         Caption         =   "Send To Address"
         Height          =   735
         Left            =   6900
         TabIndex        =   26
         Top             =   240
         Width           =   4605
         Begin VB.CommandButton bAddToAddress 
            Caption         =   "Add to Addresses"
            Height          =   465
            Left            =   3570
            TabIndex        =   27
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox cmbAddress 
            Height          =   315
            ItemData        =   "frmExternalTests.frx":79B2
            Left            =   120
            List            =   "frmExternalTests.frx":79B4
            TabIndex        =   7
            Text            =   "cmbAddress"
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Normal Ranges"
         Height          =   1215
         Left            =   6870
         TabIndex        =   20
         Top             =   1050
         Width           =   2235
         Begin VB.TextBox txtFemaleLow 
            Height          =   285
            Left            =   1290
            TabIndex        =   11
            Top             =   750
            Width           =   705
         End
         Begin VB.TextBox txtFemaleHigh 
            Height          =   285
            Left            =   1290
            TabIndex        =   10
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtMaleLow 
            Height          =   285
            Left            =   510
            TabIndex        =   9
            Top             =   750
            Width           =   705
         End
         Begin VB.TextBox txtMaleHigh 
            Height          =   285
            Left            =   510
            TabIndex        =   8
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Female"
            Height          =   195
            Left            =   1350
            TabIndex        =   24
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Male"
            Height          =   195
            Left            =   600
            TabIndex        =   23
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Low"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "High"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   330
         End
      End
      Begin VB.TextBox txtAnalyteName 
         Height          =   315
         Left            =   1590
         TabIndex        =   0
         Top             =   330
         Width           =   4605
      End
      Begin VB.ComboBox cmbSampleType 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Text            =   "cmbSampleType"
         Top             =   1140
         Width           =   1815
      End
      Begin VB.CommandButton bAddToList 
         Caption         =   "Add To List"
         Height          =   1125
         Left            =   10650
         Picture         =   "frmExternalTests.frx":79B6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Biomnis Code"
         Height          =   195
         Left            =   3360
         TabIndex        =   35
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   195
         Left            =   3510
         TabIndex        =   32
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   810
         TabIndex        =   31
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   1110
         TabIndex        =   29
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Medibridge Code"
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sample Type"
         Height          =   195
         Left            =   540
         TabIndex        =   25
         Top             =   1170
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analyte Name"
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   390
         Width           =   990
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5325
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   9393
      _Version        =   393216
      Cols            =   12
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
      AllowUserResizing=   1
      FormatString    =   $"frmExternalTests.frx":9338
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Click on analyte name to remove/edit"
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   8160
      Width           =   2640
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   33
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   13965
      TabIndex        =   18
      Top             =   4410
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmExternalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private grd As MSFlexGrid

Private Sub LoadControls()
37050     On Error GoTo LoadControls_Error

37060     txtText.Visible = False
37070     txtText = ""
          'gRD.SetFocus

37080     Select Case grd.Col
              Case 1, 2, 3, 4, 5, 8, 9, 10, 11:
37090             txtText.Move grd.Left + grd.CellLeft + 5, _
                      grd.Top + grd.CellTop + 5, _
                      grd.CellWidth - 20, grd.CellHeight - 20
37100             txtText.Text = grd.TextMatrix(grd.row, grd.Col)
37110             txtText.Visible = True
37120             txtText.SelStart = 0
37130             txtText.SelLength = Len(txtText)
37140             txtText.SetFocus

37150     End Select

37160     Exit Sub

LoadControls_Error:

          Dim strES As String
          Dim intEL As Integer

37170     intEL = Erl
37180     strES = Err.Description
37190     LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub txtText_LostFocus()

37200     On Error GoTo txtText_LostFocus_Error



37210     txtText.Visible = False

37220     Exit Sub

txtText_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

37230     intEL = Erl
37240     strES = Err.Description
37250     LogError "frmTestCodeMapping", "txtText_LostFocus", intEL, strES

End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
37260     If KeyCode = vbKeyUp Then
              'GoOneRowUp
37270     ElseIf KeyCode = vbKeyDown Then
              'GoOneRowDown
37280     ElseIf KeyCode = 13 Then
37290         txtText.Visible = False
37300         If cmdSave.Enabled = True Then cmdSave.SetFocus
37310     Else
37320         g.TextMatrix(grd.row, grd.Col) = txtText
37330         cmdSave.Enabled = True
          
37340     End If
End Sub

Private Sub baddtoaddress_Click()

37350     frmExtAddress.Show 1

37360     FillAddress

End Sub

Private Sub baddtolist_Click()

          Dim s As String

37370     On Error GoTo baddtolist_Click_Error

37380     txtAnalyteName = Trim(txtAnalyteName)

37390     If txtAnalyteName = "" Then
37400         iMsg "Enter Test Name"
37410         Exit Sub
37420     End If

37430     If cmbAddress = "" Then
37440         iMsg "Enter Address"
37450         Exit Sub
37460     End If
37470     If cmbSampleType = "" Then
37480         iMsg "Enter SampleType"
37490         Exit Sub
37500     End If
37510     If cmbDepartment = "" Then
37520         iMsg "Enter Department"
37530         Exit Sub
37540     End If
37550     s = txtAnalyteName & vbTab & _
              Format$(Val(txtMaleLow)) & vbTab & _
              Format$(Val(txtMaleHigh)) & vbTab & _
              Format$(Val(txtFemaleLow)) & vbTab & _
              Format$(Val(txtFemaleHigh)) & vbTab & _
              txtUnits & vbTab & _
              cmbAddress & vbTab & _
              cmbSampleType.Text & vbTab & _
              txtMBCode & vbTab & _
              txtBiomnisCode & vbTab & _
              cmbDepartment & vbTab & _
              txtComment
37560     g.AddItem s

37570     txtAnalyteName = ""
37580     cmbAddress = ""
37590     txtUnits = ""
37600     txtMaleHigh = ""
37610     txtMaleLow = ""
37620     txtFemaleHigh = ""
37630     txtFemaleLow = ""
37640     txtMBCode = ""
37650     txtBiomnisCode = ""
37660     cmbDepartment = ""
37670     txtComment = ""

37680     cmdSave.Enabled = True

37690     Exit Sub

baddtolist_Click_Error:

          Dim strES As String
          Dim intEL As Integer

37700     intEL = Erl
37710     strES = Err.Description
37720     LogError "frmExtTests", "baddtolist_Click", intEL, strES

End Sub

Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

37730     KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

37740     If cmdSave.Enabled Then
37750         If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
37760             Exit Sub
37770         End If
37780     End If

37790     Unload Me

End Sub

Private Sub FillAddress()

          Dim tb As Recordset
          Dim sql As String

37800     On Error GoTo FillAddress_Error

37810     sql = "Select * from eaddress"
37820     Set tb = New Recordset
37830     RecOpenServer 0, tb, sql

37840     cmbAddress.Clear
37850     Do While Not tb.EOF
37860         cmbAddress.AddItem tb!Addr0 & ""
37870         tb.MoveNext
37880     Loop

37890     Exit Sub

FillAddress_Error:

          Dim strES As String
          Dim intEL As Integer

37900     intEL = Erl
37910     strES = Err.Description
37920     LogError "frmExtTests", "FillAddress", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

37930     On Error GoTo FillG_Error

37940     g.Rows = 2
37950     g.AddItem ""
37960     g.RemoveItem 1

37970     sql = "Select * from ExternalDefinitions " & _
              "Order by AnalyteName"
37980     Set tb = New Recordset
37990     RecOpenServer 0, tb, sql
38000     Do While Not tb.EOF
38010         With tb
38020             s = !AnalyteName & vbTab & _
                      !MaleLow & vbTab & _
                      !MaleHigh & vbTab & _
                      !FemaleLow & vbTab & _
                      !FemaleHigh & vbTab & _
                      !Units & vbTab & _
                      !SendTo & vbTab & _
                      !SampleType & vbTab & _
                      !MBCode & vbTab & _
                      !BiomnisCode & vbTab & _
                      !Department & vbTab & _
                      !Comment & ""
38030             g.AddItem s
38040         End With
38050         tb.MoveNext
38060     Loop

38070     If g.Rows > 2 Then
38080         g.RemoveItem 1
38090     End If

38100     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

38110     intEL = Erl
38120     strES = Err.Description
38130     LogError "frmExtTests", "FillG", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

38140     On Error GoTo cmdSave_Click_Error

38150     pb.Min = 0
38160     pb = 0
38170     pb.max = g.Rows - 1
38180     pb.Visible = True

38190     For n = 1 To g.Rows - 1
38200         pb = n
38210         If g.TextMatrix(n, 0) <> "" Then
38220             sql = "Select * from ExternalDefinitions where " & _
                      "AnalyteName  = '" & g.TextMatrix(n, 0) & "'"
38230             Set tb = New Recordset
38240             RecOpenServer 0, tb, sql
38250             If tb.EOF Then
38260                 tb.AddNew
38270             End If
38280             tb!AnalyteName = g.TextMatrix(n, 0)
38290             tb!MaleLow = Val(g.TextMatrix(n, 1))
38300             tb!MaleHigh = Val(g.TextMatrix(n, 2))
38310             tb!FemaleLow = Val(g.TextMatrix(n, 3))
38320             tb!FemaleHigh = Val(g.TextMatrix(n, 4))
38330             tb!Units = g.TextMatrix(n, 5)
38340             tb!SendTo = g.TextMatrix(n, 6)
38350             tb!SampleType = g.TextMatrix(n, 7)
38360             tb!MBCode = g.TextMatrix(n, 8)
38370             tb!BiomnisCode = g.TextMatrix(n, 9)
38380             tb!Department = g.TextMatrix(n, 10)
38390             tb!Comment = g.TextMatrix(n, 11)
38400             tb!PrintPriority = n
38410             tb.Update
38420         End If
38430     Next

38440     cmdSave.Enabled = False

38450     pb.Visible = False

38460     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

38470     intEL = Erl
38480     strES = Err.Description
38490     LogError "frmExtTests", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

38500     ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

38510     On Error GoTo Form_Load_Error

38520     sql = "Select * from Lists where " & _
              "ListType = 'ST' and InUse = 1 " & _
              "order by ListOrder"
38530     Set tb = New Recordset
38540     RecOpenServer 0, tb, sql

38550     cmbSampleType.Clear

38560     Do While Not tb.EOF
38570         cmbSampleType.AddItem tb!Text & ""
38580         tb.MoveNext
38590     Loop

38600     FillAddress
38610     FillG

38620     cmbDepartment.Clear
38630     cmbDepartment.AddItem ""
38640     cmbDepartment.AddItem "Haematology"
38650     cmbDepartment.AddItem "Biochemistry"
38660     cmbDepartment.AddItem "Immunology"
38670     cmbDepartment.AddItem "Endocrinology"
38680     cmbDepartment.AddItem "MicroBiology"


38690     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

38700     intEL = Erl
38710     strES = Err.Description
38720     LogError "frmExtTests", "Form_Load", intEL, strES, sql


End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim s As String
          Dim sql As String

38730     On Error GoTo g_Click_Error

38740     If g.MouseRow = 0 Then
38750         If SortOrder Then
38760             g.Sort = flexSortGenericAscending
38770         Else
38780             g.Sort = flexSortGenericDescending
38790         End If
38800         SortOrder = Not SortOrder
38810         Exit Sub
38820     End If

          'If g.Col = 8 Then
          '    g.Enabled = False
          '    S = iBOX("Medibridge Code?", , g.TextMatrix(g.Row, 8))
          '    g.TextMatrix(g.Row, 8) = S
          '    g.Enabled = True
          '    cmdSave.Enabled = True
          '    Exit Sub
          'End If

38830     If g.Col = 0 Then
38840         s = "Remove " & g.TextMatrix(g.row, 0) & " from list?"
38850         If iMsg(s, vbQuestion + vbYesNo) = vbNo Then
38860             Exit Sub
38870         End If
          
38880         txtAnalyteName = g.TextMatrix(g.row, 0)
38890         txtMaleLow = g.TextMatrix(g.row, 1)
38900         txtMaleHigh = g.TextMatrix(g.row, 2)
38910         txtFemaleLow = g.TextMatrix(g.row, 3)
38920         txtFemaleHigh = g.TextMatrix(g.row, 4)
38930         txtUnits = g.TextMatrix(g.row, 5)
38940         cmbAddress = g.TextMatrix(g.row, 6) & ":"
38950         cmbSampleType = g.TextMatrix(g.row, 7)
38960         txtMBCode = g.TextMatrix(g.row, 8)
38970         cmbDepartment = g.TextMatrix(g.row, 9)
38980         txtComment = g.TextMatrix(g.row, 10)
          
38990         sql = "Delete from ExternalDefinitions where " & _
                  "AnalyteName = '" & g.TextMatrix(g.row, 0) & "'"
39000         Cnxn(0).Execute sql
          
39010         g.RemoveItem g.row
39020     Else
39030         If g.MouseRow > 0 Then
39040             Set grd = g
39050             grd.row = grd.MouseRow
39060             grd.Col = grd.MouseCol
39070             LoadControls
39080         End If
39090         Exit Sub
39100     End If


39110     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

39120     intEL = Erl
39130     strES = Err.Description
39140     LogError "frmExtTests", "g_Click", intEL, strES, sql

End Sub

Private Sub cmbaddress_KeyPress(KeyAscii As Integer)

39150     KeyAscii = 0

End Sub

Private Sub g_Scroll()
39160     Label11.Caption = UCase(Left(g.TextMatrix(g.TopRow, 0), 1))
End Sub


