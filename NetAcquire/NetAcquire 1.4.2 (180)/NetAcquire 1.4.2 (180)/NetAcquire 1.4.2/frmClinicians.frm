VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicians 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Clinician List"
   ClientHeight    =   8355
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   12705
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8355
   ScaleWidth      =   12705
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   11190
      Picture         =   "frmClinicians.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
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
      Height          =   705
      Left            =   11190
      Picture         =   "frmClinicians.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   150
      Width           =   975
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   10980
      Top             =   4290
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   11010
      Top             =   3750
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   10560
      Picture         =   "frmClinicians.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   10560
      Picture         =   "frmClinicians.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4260
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   11190
      Picture         =   "frmClinicians.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Clinician"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   10365
      Begin VB.ComboBox cmbListItems 
         Height          =   315
         ItemData        =   "frmClinicians.frx":189A
         Left            =   7890
         List            =   "frmClinicians.frx":189C
         TabIndex        =   26
         Top             =   990
         Width           =   2055
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   7890
         TabIndex        =   20
         Text            =   "cmbHospital"
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   3660
         MaxLength       =   30
         TabIndex        =   16
         Top             =   990
         Width           =   2145
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   15
         Top             =   990
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5910
         TabIndex        =   10
         Top             =   960
         Width           =   1155
      End
      Begin VB.ComboBox cmbWard 
         Height          =   315
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   300
         Width           =   2145
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   750
         MaxLength       =   12
         TabIndex        =   5
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   780
         MaxLength       =   30
         TabIndex        =   4
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Click on Code to Edit/Remove record."
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1350
         Width           =   3255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Items to be displayed in list"
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
         Left            =   7950
         TabIndex        =   25
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
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
         Left            =   7920
         TabIndex        =   24
         Top             =   30
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
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
         Left            =   3660
         TabIndex        =   19
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ForeName"
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
         Left            =   1800
         TabIndex        =   18
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title"
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
         Left            =   780
         TabIndex        =   17
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default Ward"
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
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Left            =   300
         TabIndex        =   7
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
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
         TabIndex        =   6
         Top             =   1020
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   11190
      Picture         =   "frmClinicians.frx":189E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   11190
      Picture         =   "frmClinicians.frx":1F08
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7140
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6165
      Left            =   150
      TabIndex        =   0
      Top             =   1710
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   8
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
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmClinicians.frx":2572
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   10395
      _ExtentX        =   18336
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
      Left            =   11070
      TabIndex        =   22
      Top             =   870
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmClinicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer





Private Sub cmbListItems_Click()
13750     cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
13760     KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()

          Dim ClinFullName As String

13770     On Error GoTo cmdAdd_Click_Error

13780     txtCode = UCase$(Trim$(txtCode))
13790     If txtCode = "" Then
13800         iMsg "Enter Code.", vbCritical
13810         Exit Sub
13820     End If

13830     txtSurName = Trim$(txtSurName)
13840     If txtSurName = "" Then
13850         iMsg "Enter Clinicians Surname.", vbCritical
13860         Exit Sub
13870     End If

13880     ClinFullName = Trim$(txtTitle) & " " & Trim$(txtForeName) & " " & Trim$(txtSurName)

13890     g.AddItem "Yes" & vbTab & _
              txtCode & vbTab & _
              txtTitle & vbTab & _
              txtForeName & vbTab & _
              txtSurName & vbTab & _
              ClinFullName & vbTab & _
              cmbWard

13900     txtCode = ""
13910     txtTitle = ""
13920     txtForeName = ""
13930     txtSurName = ""
13940     cmbWard.ListIndex = -1
        
13950     cmdSave.Visible = True

13960     Exit Sub

cmdAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

13970     intEL = Erl
13980     strES = Err.Description
13990     LogError "fclinicians", "cmdadd_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

14000     Unload Me

End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

14010     If g.row = g.Rows - 1 Then Exit Sub
14020     n = g.row

14030     VisibleRows = g.height \ g.RowHeight(1) - 1

14040     FireCounter = FireCounter + 1
14050     If FireCounter > 5 Then
14060         tmrDown.Interval = 100
14070     End If

14080     g.Visible = False

14090     s = ""
14100     For X = 0 To g.Cols - 1
14110         s = s & g.TextMatrix(n, X) & vbTab
14120     Next
14130     s = Left$(s, Len(s) - 1)

14140     g.RemoveItem n
14150     If n < g.Rows Then
14160         g.AddItem s, n + 1
14170         g.row = n + 1
14180     Else
14190         g.AddItem s
14200         g.row = g.Rows - 1
14210     End If

14220     For X = 0 To g.Cols - 1
14230         g.Col = X
14240         g.CellBackColor = vbYellow
14250     Next

14260     If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
14270         If g.row - VisibleRows + 1 > 0 Then
14280             g.TopRow = g.row - VisibleRows + 1
14290         End If
14300     End If

14310     g.Visible = True

14320     cmdSave.Visible = True

End Sub
Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

14330     If g.row = 1 Then Exit Sub

14340     FireCounter = FireCounter + 1
14350     If FireCounter > 5 Then
14360         tmrUp.Interval = 100
14370     End If

14380     n = g.row

14390     g.Visible = False

14400     s = ""
14410     For X = 0 To g.Cols - 1
14420         s = s & g.TextMatrix(n, X) & vbTab
14430     Next
14440     s = Left$(s, Len(s) - 1)

14450     g.RemoveItem n
14460     g.AddItem s, n - 1

14470     g.row = n - 1
14480     For X = 0 To g.Cols - 1
14490         g.Col = X
14500         g.CellBackColor = vbYellow
14510     Next

14520     If Not g.RowIsVisible(g.row) Then
14530         g.TopRow = g.row
14540     End If

14550     g.Visible = True

14560     cmdSave.Visible = True

End Sub



Private Sub cmdDelete_Click()
          'check if record doesn't exist in demographics
          Dim tb As Recordset
          Dim sql As String
          Dim strFullName As String

14570     On Error GoTo cmdDelete_Click_Error

14580     If g.row = 0 Or g.Rows <= 2 Then Exit Sub

14590     strFullName = g.TextMatrix(g.row, 2) & " " & g.TextMatrix(g.row, 3) & " " & g.TextMatrix(g.row, 4)

14600     sql = "SELECT Count(*) as RC FROM Demographics WHERE Clinician = '" & strFullName & "'"
14610     Set tb = New Recordset
14620     RecOpenClient 0, tb, sql
14630     If tb!rc > 0 Then
14640         iMsg "Reference to " & strFullName & " is in use so cannot be edited/removed"
14650         Exit Sub
14660     Else
14670         If iMsg("Are you sure you want to delete " & strFullName & "?", vbYesNo) = vbYes Then
14680             Cnxn(0).Execute "DELETE FROM Clinicians WHERE " & _
                      "Code = '" & g.TextMatrix(g.row, 1) & "' AND Text = '" & strFullName & "'"
14690             FillG
14700         End If
          
14710     End If

14720     Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

14730     intEL = Erl
14740     strES = Err.Description
14750     LogError "fclinicians", "cmdDelete_Click", intEL, strES, sql

          
End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14760     FireDown

14770     tmrDown.Interval = 250
14780     FireCounter = 0

14790     tmrDown.Enabled = True

End Sub

Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14800     tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14810     FireUp

14820     tmrUp.Interval = 250
14830     FireCounter = 0

14840     tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

14850     tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

          Dim Y As Integer
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim i As Integer

14860     Screen.MousePointer = vbHourglass

14870     OriginalPrinter = Printer.DeviceName

14880     If Not SetFormPrinter() Then Exit Sub

14890     Printer.FontName = "Courier New"
14900     Printer.Orientation = vbPRORPortrait

          '****Report heading
14910     Printer.FontSize = 10
14920     Printer.Font.Bold = True
14930     Printer.Print
14940     Printer.Print FormatString("List Of Clinicians", 99, , AlignCenter)

          '****Report body heading

14950     Printer.Font.size = 9
14960     For i = 1 To 108
14970         Printer.Print "-";
14980     Next i
14990     Printer.Print


15000     Printer.Print FormatString("", 0, "|");
15010     Printer.Print FormatString("In Use", 6, "|", AlignCenter);
15020     Printer.Print FormatString("Code", 10, "|", AlignCenter);
15030     Printer.Print FormatString("Name", 70, "|", AlignCenter);
15040     Printer.Print FormatString("Ward", 17, "|", AlignCenter)
          '****Report body

15050     Printer.Font.Bold = False

15060     For i = 1 To 108
15070         Printer.Print "-";
15080     Next i
15090     Printer.Print
15100     For Y = 1 To g.Rows - 1
15110         Printer.Print FormatString("", 0, "|");
15120         Printer.Print FormatString(g.TextMatrix(Y, 0), 6, "|", AlignCenter);
15130         Printer.Print FormatString(g.TextMatrix(Y, 1), 10, "|", AlignLeft);
15140         Printer.Print FormatString(g.TextMatrix(Y, 5), 70, "|", AlignLeft);
15150         Printer.Print FormatString(g.TextMatrix(Y, 6), 17, "|", AlignLeft)
15160         Printer.Print FormatString(g.TextMatrix(Y, 7), 3, "|", AlignLeft)
       
15170     Next

15180     Printer.EndDoc

15190     Screen.MousePointer = vbDefault

15200     For Each Px In Printers
15210         If Px.DeviceName = OriginalPrinter Then
15220             Set Printer = Px
15230             Exit For
15240         End If
15250     Next
End Sub

Private Sub cmdSave_Click()

          Dim HospCode As String
          Dim Y As Integer
          Dim sql As String
          Dim tb As Recordset

15260     On Error GoTo cmdSave_Click_Error

15270     HospCode = ListCodeFor("HO", cmbHospital)

15280     pb.max = g.Rows - 1
15290     pb.Visible = True
15300     cmdSave.Caption = "Saving..."

15310     For Y = 1 To g.Rows - 1
15320         pb = Y
15330         sql = "Select * from Clinicians where " & _
                  "HospitalCode = '" & HospCode & "' " & _
                  "and Code = '" & g.TextMatrix(Y, 1) & "'"
15340         Set tb = New Recordset
15350         RecOpenServer 0, tb, sql
15360         If tb.EOF Then
15370             tb.AddNew
15380         End If
15390         With tb
15400             !HospitalCode = HospCode
15410             !InUse = g.TextMatrix(Y, 0) = "Yes"
15420             !Code = g.TextMatrix(Y, 1)
15430             !Title = g.TextMatrix(Y, 2)
15440             !ForeName = g.TextMatrix(Y, 3)
15450             !SurName = g.TextMatrix(Y, 4)
15460             !Text = g.TextMatrix(Y, 5)
15470             !Ward = g.TextMatrix(Y, 6)
15480             !ListOrder = Y
15490             .Update

15500             sql = "IF EXISTS (SELECT * FROM IncludeEGFR " & _
                      "           WHERE SourceType = 'Clinician' " & _
                      "           AND Hospital = '" & cmbHospital & "' " & _
                      "           AND SourceName = '" & AddTicks(g.TextMatrix(Y, 5)) & "' ) " & _
                      "    UPDATE IncludeEGFR " & _
                      "    SET Include = '" & IIf(g.TextMatrix(Y, 7) = "Yes", 1, 0) & "' " & _
                      "    WHERE SourceType = 'Clinician' " & _
                      "    AND SourceName = '" & AddTicks(g.TextMatrix(Y, 5)) & "' " & _
                      "ELSE " & _
                      "    INSERT INTO IncludeEGFR (SourceType, Hospital, SourceName, Include) " & _
                      "    VALUES ('Clinician', " & _
                      "            '" & cmbHospital & "', " & _
                      "            '" & AddTicks(g.TextMatrix(Y, 5)) & "', " & _
                      "            '" & IIf(g.TextMatrix(Y, 7) = "Yes", 1, 0) & "')"
15510             Cnxn(0).Execute sql

15520         End With
15530     Next

15540     Call SaveOptionSetting("ClinicianListLength", cmbListItems)

15550     pb.Visible = False
15560     cmdSave.Visible = False
15570     cmdSave.Caption = "Save"

15580     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

15590     intEL = Erl
15600     strES = Err.Description
15610     LogError "fclinicians", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmbHospital_Click()

15620     FillWards cmbWard, cmbHospital
15630     FillG

End Sub

Private Sub cmdXL_Click()

15640     ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

15650     If Activated Then
15660         Exit Sub
15670     End If

15680     Activated = True

15690     FillG

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

15700     On Error GoTo Form_Load_Error

15710     sql = "Select * from Lists where " & _
              "ListType = 'HO' and InUse = 1 " & _
              "order by ListOrder"
15720     Set tb = New Recordset
15730     RecOpenServer 0, tb, sql
15740     Do While Not tb.EOF
15750         cmbHospital.AddItem tb!Text & ""
15760         tb.MoveNext
15770     Loop

15780     If cmbHospital.ListCount > 0 Then
15790         cmbHospital.ListIndex = 0
15800     End If

15810     FillWards cmbWard, cmbHospital

15820     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

15830     intEL = Erl
15840     strES = Err.Description
15850     LogError "fclinicians", "Form_Load", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String

15860     On Error GoTo FillG_Error

15870     g.Rows = 2
15880     g.AddItem ""
15890     g.RemoveItem 1

15900     sql = "SELECT C.*, COALESCE(E.Include, 0) EGFR " & _
              "FROM Clinicians C JOIN Lists L " & _
              "ON C.HospitalCode = L.Code " & _
              "LEFT JOIN IncludeEGFR E " & _
              "ON C.Text = E.SourceName " & _
              "WHERE L.Text = '" & cmbHospital & "' " & _
              "AND L.ListType = 'HO'"
          'sql = "Select * from Clinicians order by ListOrder"
          'If cmbHospital <> "" Then
          '  Hosp = ListCodeFor("HO", cmbHospital)
          '  sql = "Select * from Clinicians where " & _
          '        "HospitalCode = '" & Hosp & "' " & _
          '        "order  by ListOrder"
          '  End If

15910     Set tb = New Recordset
15920     RecOpenServer 0, tb, sql

15930     Do While Not tb.EOF
15940         s = IIf(tb!InUse, "Yes", "No") & vbTab & _
                  tb!Code & vbTab & _
                  tb!Title & vbTab & _
                  tb!ForeName & vbTab & _
                  tb!SurName & vbTab & _
                  tb!Text & vbTab & _
                  tb!Ward & vbTab & _
                  IIf(tb!EGFR, "Yes", "No")
15950         g.AddItem s
15960         tb.MoveNext
15970     Loop

15980     If g.Rows > 2 Then
15990         g.RemoveItem 1
16000     End If

16010     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

16020     intEL = Erl
16030     strES = Err.Description
16040     LogError "fclinicians", "FillG", intEL, strES, sql

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

16050     If cmdSave.Visible Then
16060         If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
16070             Cancel = True
16080         End If
16090     End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

16100     Activated = False

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim strFullName As String

16110     On Error GoTo g_Click_Error

16120     ySave = g.row

16130     If g.MouseRow = 0 Then
16140         If SortOrder Then
16150             g.Sort = flexSortGenericAscending
16160         Else
16170             g.Sort = flexSortGenericDescending
16180         End If
16190         SortOrder = Not SortOrder
16200         cmdMoveUp.Enabled = False
16210         cmdMoveDown.Enabled = False
16220         cmdSave.Visible = True
16230         Exit Sub
16240     End If

16250     If g.Col = 0 Or g.Col = 7 Then
16260         g = IIf(g = "No", "Yes", "No")
16270         cmdSave.Visible = True
16280         Exit Sub
16290     End If

16300     If g.Col = 1 Then
16310         strFullName = g.TextMatrix(g.row, 2) & " " & g.TextMatrix(g.row, 3) & " " & g.TextMatrix(g.row, 4)

16320         sql = "SELECT Count(*) as RC FROM Demographics WHERE Clinician = '" & strFullName & "'"
16330         Set tb = New Recordset
16340         RecOpenClient 0, tb, sql
16350         If tb!rc > 0 Then
16360             iMsg "Reference to " & strFullName & " is in use so cannot be edited/removed"
16370             Exit Sub
16380         End If
16390         g.Enabled = False
16400         If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then
16410             txtCode = g.TextMatrix(g.row, 1)
16420             txtTitle = g.TextMatrix(g.row, 2)
16430             txtForeName = g.TextMatrix(g.row, 3)
16440             txtSurName = g.TextMatrix(g.row, 4)
16450             g.RemoveItem g.row
16460             cmdSave.Visible = True
16470         End If
16480         g.Enabled = True
16490         Exit Sub
16500     End If
          
16510     g.Visible = False
16520     g.Col = 0
16530     For Y = 1 To g.Rows - 1
16540         g.row = Y
16550         If g.CellBackColor = vbYellow Then
16560             For X = 0 To g.Cols - 1
16570                 g.Col = X
16580                 g.CellBackColor = 0
16590             Next
16600             Exit For
16610         End If
16620     Next
16630     g.row = ySave
16640     g.Visible = True

16650     For X = 0 To g.Cols - 1
16660         g.Col = X
16670         g.CellBackColor = vbYellow
16680     Next

16690     cmdMoveUp.Enabled = True
16700     cmdMoveDown.Enabled = True

16710     Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

16720     intEL = Erl
16730     strES = Err.Description
16740     LogError "fclinicians", "g_Click", intEL, strES, sql

End Sub

Private Sub g_GotFocus()
    'cmdDelete.Enabled = (g.Rows > 2)
End Sub

Private Sub g_LostFocus()
    'cmdDelete.Enabled = False
End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16750     If g.MouseRow = 0 Then
16760         g.ToolTipText = ""
16770     ElseIf g.MouseCol = 0 Then
16780         g.ToolTipText = "Click to Toggle"
16790     ElseIf g.MouseCol = 1 Then
16800         g.ToolTipText = "Click to Edit"
16810     Else
16820         g.ToolTipText = "Click to Move"
16830     End If

End Sub


Private Sub tmrDown_Timer()

16840     FireDown

End Sub


Private Sub tmrUp_Timer()

16850     FireUp

End Sub


