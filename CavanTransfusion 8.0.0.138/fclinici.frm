VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fclinicians 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Clinician List"
   ClientHeight    =   8535
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   9885
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
   ScaleHeight     =   8535
   ScaleWidth      =   9885
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   28
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
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
      Left            =   2730
      Picture         =   "fclinici.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7740
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
      Left            =   150
      Picture         =   "fclinici.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7770
      Width           =   1245
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   9360
      Top             =   7830
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7950
      Top             =   7830
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8430
      Picture         =   "fclinici.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8880
      Picture         =   "fclinici.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   6450
      Picture         =   "fclinici.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7770
      Visible         =   0   'False
      Width           =   1365
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
      Top             =   240
      Width           =   9645
      Begin VB.ComboBox cmbListItems 
         Height          =   315
         ItemData        =   "fclinici.frx":189A
         Left            =   7410
         List            =   "fclinici.frx":189C
         TabIndex        =   27
         Top             =   1110
         Width           =   1875
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   7380
         TabIndex        =   20
         Text            =   "cmbHospital"
         Top             =   480
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
         Left            =   7410
         TabIndex        =   26
         Top             =   870
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
         Left            =   7380
         TabIndex        =   25
         Top             =   240
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
      Left            =   3990
      Picture         =   "fclinici.frx":189E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7770
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
      Left            =   5250
      Picture         =   "fclinici.frx":1F08
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7770
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5355
      Left            =   150
      TabIndex        =   0
      Top             =   1890
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9446
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   $"fclinici.frx":2572
   End
   Begin MSComctlLib.ProgressBar pB 
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   7530
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Click on Code to Edit/Remove record."
      Height          =   195
      Left            =   150
      TabIndex        =   24
      Top             =   7290
      Width           =   3255
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   1410
      TabIndex        =   22
      Top             =   7920
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "fclinicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer





Private Sub cmbListItems_Click()
10    cmdSave.Visible = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
10    KeyAscii = 0
End Sub

Private Sub cmdadd_Click()

      Dim ClinFullName As String

10    On Error GoTo cmdAdd_Click_Error

20    txtCode = UCase$(Trim$(txtCode))
30    If txtCode = "" Then
40      iMsg "Enter Code.", vbCritical
50      Exit Sub
60    End If

70    txtSurname = Trim$(txtSurname)
80    If txtSurname = "" Then
90      iMsg "Enter Clinicians Surname.", vbCritical
100     Exit Sub
110   End If

120   ClinFullName = Trim$(txtTitle) & " " & Trim$(txtForeName) & " " & Trim$(txtSurname)

130   g.AddItem "Yes" & vbTab & _
                txtCode & vbTab & _
                txtTitle & vbTab & _
                txtForeName & vbTab & _
                txtSurname & vbTab & _
                ClinFullName & vbTab & _
                cmbWard

140   txtCode = ""
150   txtTitle = ""
160   txtForeName = ""
170   txtSurname = ""
180   cmbWard.ListIndex = -1
  
190   cmdSave.Visible = True

200   Exit Sub

cmdAdd_Click_Error:

Dim strES As String
Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "fclinicians", "cmdadd_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub FireDown()

      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim VisibleRows As Integer

10    If g.Row = g.Rows - 1 Then Exit Sub
20    n = g.Row

30    VisibleRows = g.Height \ g.RowHeight(1) - 1

40    FireCounter = FireCounter + 1
50    If FireCounter > 5 Then
60      tmrDown.Interval = 100
70    End If

80    g.Visible = False

90    s = ""
100   For X = 0 To g.Cols - 1
110     s = s & g.TextMatrix(n, X) & vbTab
120   Next
130   s = Left$(s, Len(s) - 1)

140   g.RemoveItem n
150   If n < g.Rows Then
160     g.AddItem s, n + 1
170     g.Row = n + 1
180   Else
190     g.AddItem s
200     g.Row = g.Rows - 1
210   End If

220   For X = 0 To g.Cols - 1
230     g.Col = X
240     g.CellBackColor = vbYellow
250   Next

260   If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
270     If g.Row - VisibleRows + 1 > 0 Then
280       g.TopRow = g.Row - VisibleRows + 1
290     End If
300   End If

310   g.Visible = True

320   cmdSave.Visible = True

End Sub
Private Sub FireUp()

      Dim n As Integer
      Dim s As String
      Dim X As Integer

10    If g.Row = 1 Then Exit Sub

20    FireCounter = FireCounter + 1
30    If FireCounter > 5 Then
40      tmrUp.Interval = 100
50    End If

60    n = g.Row

70    g.Visible = False

80    s = ""
90    For X = 0 To g.Cols - 1
100     s = s & g.TextMatrix(n, X) & vbTab
110   Next
120   s = Left$(s, Len(s) - 1)

130   g.RemoveItem n
140   g.AddItem s, n - 1

150   g.Row = n - 1
160   For X = 0 To g.Cols - 1
170     g.Col = X
180     g.CellBackColor = vbYellow
190   Next

200   If Not g.RowIsVisible(g.Row) Then
210     g.TopRow = g.Row
220   End If

230   g.Visible = True

240   cmdSave.Visible = True

End Sub



Private Sub cmdDelete_Click()
      'check if record doesn't exist in demographics
      Dim tb As Recordset
      Dim sql As String
      Dim strFullName As String

10    On Error GoTo cmdDelete_Click_Error

20    If g.Row = 0 Or g.Rows <= 2 Then Exit Sub

30    strFullName = g.TextMatrix(g.Row, 2) & " " & g.TextMatrix(g.Row, 3) & " " & g.TextMatrix(g.Row, 4)

40    sql = "SELECT Count(*) as RC FROM Demographics WHERE Clinician = '" & strFullName & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If tb!rc > 0 Then
80        iMsg "Reference to " & strFullName & " is in use so cannot be edited/removed"
90        Exit Sub
100   Else
110       If iMsg("Are you sure you want to delete " & strFullName & "?", vbYesNo) = vbYes Then
120           Cnxn(0).Execute "DELETE FROM Clinicians WHERE " & _
                      "Code = '" & g.TextMatrix(g.Row, 1) & "' AND Text = '" & strFullName & "'"
130           FillG
140       End If
    
150   End If

160   Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fclinicians", "cmdDelete_Click", intEL, strES, sql

    
End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireDown

20    tmrDown.Interval = 250
30    FireCounter = 0

40    tmrDown.Enabled = True

End Sub

Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    FireUp

20    tmrUp.Interval = 250
30    FireCounter = 0

40    tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    On Error GoTo cmdPrint_Click_Error

20    OriginalPrinter = Printer.DeviceName

30    If Not SetFormPrinter() Then Exit Sub

40    Printer.FontName = "Courier New"
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.FontSize = 10
70    Printer.Font.Bold = True
80    Printer.Print
90    Printer.Print FormatString("List Of Clinicians", 99, , AlignCenter)

      '****Report body heading

100   Printer.Font.Size = 9
110   For i = 1 To 108
120       Printer.Print "-";
130   Next i
140   Printer.Print


150   Printer.Print FormatString("", 0, "|");
160   Printer.Print FormatString("In Use", 6, "|", AlignCenter);
170   Printer.Print FormatString("Code", 10, "|", AlignCenter);
180   Printer.Print FormatString("Name", 70, "|", AlignCenter);
190   Printer.Print FormatString("Ward", 17, "|", AlignCenter)
      '****Report body

200   Printer.Font.Bold = False

210   For i = 1 To 108
220       Printer.Print "-";
230   Next i
240   Printer.Print
250   For Y = 1 To g.Rows - 1
260       Printer.Print FormatString("", 0, "|");
270       Printer.Print FormatString(g.TextMatrix(Y, 0), 6, "|", AlignCenter);
280       Printer.Print FormatString(g.TextMatrix(Y, 1), 10, "|", AlignLeft);
290       Printer.Print FormatString(g.TextMatrix(Y, 5), 70, "|", AlignLeft);
300       Printer.Print FormatString(g.TextMatrix(Y, 6), 17, "|", AlignLeft)
310   Next

320   Printer.EndDoc

330   For Each Px In Printers
340     If Px.DeviceName = OriginalPrinter Then
350       Set Printer = Px
360       Exit For
370     End If
380   Next

390   Exit Sub

cmdPrint_Click_Error:

Dim strES As String
Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "fclinicians", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

      Dim HospCode As String
      Dim Y As Integer
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmdSave_Click_Error

20    HospCode = ListCodeFor("HO", cmbHospital)

30    PB.max = g.Rows - 1
40    PB.Visible = True
50    cmdSave.Caption = "Saving..."

60    For Y = 1 To g.Rows - 1
70      PB = Y
80      sql = "Select * from Clinicians where " & _
              "HospitalCode = '" & HospCode & "' " & _
              "and Code = '" & g.TextMatrix(Y, 1) & "'"
90      Set tb = New Recordset
100     RecOpenServer 0, tb, sql
110     If tb.EOF Then
120       tb.AddNew
130     End If
140     With tb
150       !HospitalCode = HospCode
160       !InUse = g.TextMatrix(Y, 0) = "Yes"
170       !code = g.TextMatrix(Y, 1)
180       !Title = g.TextMatrix(Y, 2)
190       !ForeName = g.TextMatrix(Y, 3)
200       !SurName = g.TextMatrix(Y, 4)
210       !Text = g.TextMatrix(Y, 5)
220       !Ward = g.TextMatrix(Y, 6)
230       !ListOrder = Y
240       .Update
250     End With
260   Next

270   Call SaveOptionSetting("ClinicianListLength", cmbListItems)

280   PB.Visible = False
290   cmdSave.Visible = False
300   cmdSave.Caption = "Save"

310   Exit Sub

cmdSave_Click_Error:

Dim strES As String
Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "fclinicians", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmbHospital_Click()

10    FillWards cmbWard, cmbHospital
20    FillG

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

10    If Activated Then
20      Exit Sub
30    End If

40    Activated = True

50    FillG

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Form_Load_Error

20    sql = "Select * from Lists where " & _
            "ListType = 'HO' and InUse = 1 " & _
            "order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    Do While Not tb.EOF
60      cmbHospital.AddItem tb!Text & ""
70      tb.MoveNext
80    Loop

90    If cmbHospital.ListCount > 0 Then
100     cmbHospital.ListIndex = 0
110   End If

120   FillWards cmbWard, cmbHospital

130   Exit Sub

Form_Load_Error:

Dim strES As String
Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "fclinicians", "Form_Load", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim s As String
      Dim Hosp As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Clinicians order by ListOrder"
60    If cmbHospital <> "" Then
70      Hosp = ListCodeFor("HO", cmbHospital)
80      sql = "Select * from Clinicians where " & _
              "HospitalCode = '" & Hosp & "' " & _
              "order  by ListOrder"
90      End If

100   Set tb = New Recordset
110   RecOpenServer 0, tb, sql

120   Do While Not tb.EOF
130     s = IIf(tb!InUse, "Yes", "No") & vbTab & _
            tb!code & vbTab & _
            tb!Title & vbTab & _
            tb!ForeName & vbTab & _
            tb!SurName & vbTab & _
            tb!Text & vbTab & _
            tb!Ward & ""
140     g.AddItem s
150     tb.MoveNext
160   Loop

170   If g.Rows > 2 Then
180     g.RemoveItem 1
190   End If

200   Exit Sub

FillG_Error:

Dim strES As String
Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "fclinicians", "FillG", intEL, strES, sql

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If cmdSave.Visible Then
20      If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
30        Cancel = True
40      End If
50    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim strFullName As String

10    On Error GoTo g_Click_Error

20    ySave = g.Row

30    If g.MouseRow = 0 Then
40      If SortOrder Then
50        g.Sort = flexSortGenericAscending
60      Else
70        g.Sort = flexSortGenericDescending
80      End If
90      SortOrder = Not SortOrder
100     cmdMoveUp.Enabled = False
110     cmdMoveDown.Enabled = False
120     cmdSave.Visible = True
130     Exit Sub
140   End If

150   If g.Col = 0 Then
160     g = IIf(g = "No", "Yes", "No")
170     cmdSave.Visible = True
180     Exit Sub
190   End If

200   If g.Col = 1 Then
210       strFullName = g.TextMatrix(g.Row, 2) & " " & g.TextMatrix(g.Row, 3) & " " & g.TextMatrix(g.Row, 4)

220       sql = "SELECT Count(*) as RC FROM Demographics WHERE Clinician = '" & strFullName & "'"
230       Set tb = New Recordset
240       RecOpenClient 0, tb, sql
250       If tb!rc > 0 Then
260           iMsg "Reference to " & strFullName & " is in use so cannot be edited/removed"
270           Exit Sub
280       End If
290     g.Enabled = False
300     If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then
310       txtCode = g.TextMatrix(g.Row, 1)
320       txtTitle = g.TextMatrix(g.Row, 2)
330       txtForeName = g.TextMatrix(g.Row, 3)
340       txtSurname = g.TextMatrix(g.Row, 4)
350       g.RemoveItem g.Row
360       cmdSave.Visible = True
370     End If
380     g.Enabled = True
390     Exit Sub
400   End If
    
410   g.Visible = False
420   g.Col = 0
430   For Y = 1 To g.Rows - 1
440     g.Row = Y
450     If g.CellBackColor = vbYellow Then
460       For X = 0 To g.Cols - 1
470         g.Col = X
480         g.CellBackColor = 0
490       Next
500       Exit For
510     End If
520   Next
530   g.Row = ySave
540   g.Visible = True

550   For X = 0 To g.Cols - 1
560     g.Col = X
570     g.CellBackColor = vbYellow
580   Next

590   cmdMoveUp.Enabled = True
600   cmdMoveDown.Enabled = True

610   Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "fclinicians", "g_Click", intEL, strES, sql

End Sub

Private Sub g_GotFocus()
                                                            'cmdDelete.Enabled = (g.Rows > 2)
End Sub

Private Sub g_LostFocus()
                                                            'cmdDelete.Enabled = False
End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    If g.MouseRow = 0 Then
20      g.ToolTipText = ""
30    ElseIf g.MouseCol = 0 Then
40      g.ToolTipText = "Click to Toggle"
50    ElseIf g.MouseCol = 1 Then
60      g.ToolTipText = "Click to Edit"
70    Else
80      g.ToolTipText = "Click to Move"
90    End If

End Sub


Private Sub tmrDown_Timer()

10    FireDown

End Sub


Private Sub tmrUp_Timer()

10    FireUp

End Sub


