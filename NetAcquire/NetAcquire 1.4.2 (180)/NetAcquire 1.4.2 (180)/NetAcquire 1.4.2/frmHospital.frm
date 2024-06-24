VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHospital 
   Caption         =   "NetAcquire - Hospitals"
   ClientHeight    =   7050
   ClientLeft      =   1830
   ClientTop       =   1110
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7230
   Begin VB.ComboBox cmbListItems 
      Height          =   315
      ItemData        =   "frmHospital.frx":0000
      Left            =   4860
      List            =   "frmHospital.frx":0002
      TabIndex        =   12
      Top             =   420
      Width           =   1875
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   6150
      Picture         =   "frmHospital.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1350
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   6180
      Picture         =   "frmHospital.frx":066E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "frmHospital.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3630
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "frmHospital.frx":265A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4470
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6150
      Picture         =   "frmHospital.frx":3FDC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Hospital"
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   4575
      Begin VB.TextBox tCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox tText 
         Height          =   285
         Left            =   810
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3645
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   3750
         TabIndex        =   1
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   630
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   90
      TabIndex        =   6
      Top             =   1350
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   9604
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Code       |<Text                                                                     "
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Items to be displayed in list"
      Height          =   195
      Left            =   4860
      TabIndex        =   13
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean







Private Sub cmbListItems_Click()
6040      bsave.Enabled = True
End Sub

Private Sub cmbListItems_KeyPress(KeyAscii As Integer)
6050      KeyAscii = 0
End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

6060      On Error GoTo FillG_Error

6070      g.Rows = 2
6080      g.AddItem ""
6090      g.RemoveItem 1

6100      sql = "Select * from Lists where " & _
              "ListType = 'HO' and InUse = 1 " & _
              "order by ListOrder"
6110      Set tb = New Recordset
6120      RecOpenServer 0, tb, sql
6130      Do While Not tb.EOF
6140          s = tb!Code & vbTab & tb!Text & ""
6150          g.AddItem s
6160          tb.MoveNext
6170      Loop

6180      If g.Rows > 2 Then
6190          g.RemoveItem 1
6200      End If

6210      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

6220      intEL = Erl
6230      strES = Err.Description
6240      LogError "fHospital", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

6250      tCode = Trim$(UCase$(tCode))
6260      tText = Trim$(tText)

6270      If tCode = "" Then
6280          Exit Sub
6290      End If

6300      If tText = "" Then Exit Sub

6310      g.AddItem tCode & vbTab & tText

6320      tCode = ""
6330      tText = ""

6340      bsave.Enabled = True

End Sub


Private Sub bcancel_Click()

6350      Unload Me

End Sub

Private Sub bMoveDown_Click()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

6360      If g.row = g.Rows - 1 Then Exit Sub
6370      n = g.row

6380      s = ""
6390      For X = 0 To g.Cols - 1
6400          s = s & g.TextMatrix(n, X) & vbTab
6410      Next
6420      s = Left$(s, Len(s) - 1)

6430      g.RemoveItem n
6440      If n < g.Rows Then
6450          g.AddItem s, n + 1
6460          g.row = n + 1
6470      Else
6480          g.AddItem s
6490          g.row = g.Rows - 1
6500      End If

6510      For X = 0 To g.Cols - 1
6520          g.Col = X
6530          g.CellBackColor = vbYellow
6540      Next

6550      bsave.Enabled = True

End Sub


Private Sub bMoveUp_Click()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

6560      If g.row = 1 Then Exit Sub

6570      n = g.row

6580      s = ""
6590      For X = 0 To g.Cols - 1
6600          s = s & g.TextMatrix(n, X) & vbTab
6610      Next
6620      s = Left$(s, Len(s) - 1)

6630      g.RemoveItem n
6640      g.AddItem s, n - 1

6650      g.row = n - 1
6660      For X = 0 To g.Cols - 1
6670          g.Col = X
6680          g.CellBackColor = vbYellow
6690      Next

6700      bsave.Enabled = True

End Sub


Private Sub bPrint_Click()

          Dim Y As Integer
          Dim Px As Printer
          Dim OriginalPrinter As String
          Dim i As Integer

6710      Screen.MousePointer = 11

6720      OriginalPrinter = Printer.DeviceName

6730      If Not SetFormPrinter() Then Exit Sub

6740      Printer.FontName = "Courier New"
6750      Printer.Orientation = vbPRORPortrait


          '****Report heading
6760      Printer.FontSize = 10
6770      Printer.Font.Bold = True
6780      Printer.Print
6790      Printer.Print FormatString("List Of Hospitals", 99, , AlignCenter)

          '****Report body heading

6800      Printer.Font.size = 9
6810      For i = 1 To 108
6820          Printer.Print "-";
6830      Next i
6840      Printer.Print


6850      Printer.Print FormatString("", 0, "|");
6860      Printer.Print FormatString("Code", 10, "|", AlignCenter);
6870      Printer.Print FormatString("Description", 95, "|", AlignCenter)
          '****Report body

6880      Printer.Font.Bold = False

6890      For i = 1 To 108
6900          Printer.Print "-";
6910      Next i
6920      Printer.Print
6930      For Y = 1 To g.Rows - 1
6940          Printer.Print FormatString("", 0, "|");
6950          Printer.Print FormatString(g.TextMatrix(Y, 0), 10, "|", AlignLeft);
6960          Printer.Print FormatString(g.TextMatrix(Y, 1), 95, "|", AlignLeft)

       
6970      Next



6980      Printer.EndDoc

6990      Screen.MousePointer = 0

7000      For Each Px In Printers
7010          If Px.DeviceName = OriginalPrinter Then
7020              Set Printer = Px
7030              Exit For
7040          End If
7050      Next




End Sub


Private Sub bSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer

7060      On Error GoTo bSave_Click_Error

7070      For Y = 1 To g.Rows - 1
7080          sql = "Select * from Lists where " & _
                  "ListType = 'HO' and InUse = 1 " & _
                  "and Code = '" & g.TextMatrix(Y, 0) & "'"
7090          Set tb = New Recordset
7100          RecOpenServer 0, tb, sql
7110          If tb.EOF Then
7120              tb.AddNew
7130          End If
7140          tb!Code = g.TextMatrix(Y, 0)
7150          tb!ListType = "HO"
7160          tb!Text = g.TextMatrix(Y, 1)
7170          tb!ListOrder = Y
7180          tb!InUse = 1
7190          tb.Update
        
7200      Next

7210      Call SaveOptionSetting("HospitalListLength", cmbListItems)

7220      FillG

7230      tCode = ""
7240      tText = ""
7250      tCode.SetFocus
7260      bMoveUp.Enabled = False
7270      bMoveDown.Enabled = False
7280      bsave.Enabled = False

7290      Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

7300      intEL = Erl
7310      strES = Err.Description
7320      LogError "fHospital", "bSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

7330      If Activated Then Exit Sub

7340      Activated = True

7350      FillG

End Sub

Private Sub Form_Load()

7360      g.Font.Bold = True

7370      Activated = False

          Dim i  As Integer
7380      cmbListItems.Clear
7390      For i = 8 To 32 Step 8
7400          cmbListItems.AddItem i
7410      Next i
7420      cmbListItems.Text = GetOptionSetting("HospitalListLength", 8)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

7430      If bsave.Enabled Then
7440          If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
7450              Cancel = True
7460              Exit Sub
7470          End If
7480      End If

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

7490      ySave = g.row

7500      g.Visible = False
7510      g.Col = 0
7520      For Y = 1 To g.Rows - 1
7530          g.row = Y
7540          If g.CellBackColor = vbYellow Then
7550              For X = 0 To g.Cols - 1
7560                  g.Col = X
7570                  g.CellBackColor = 0
7580              Next
7590              Exit For
7600          End If
7610      Next
7620      g.row = ySave
7630      g.Visible = True

7640      If g.MouseRow = 0 Then
7650          If SortOrder Then
7660              g.Sort = flexSortGenericAscending
7670          Else
7680              g.Sort = flexSortGenericDescending
7690          End If
7700          SortOrder = Not SortOrder
7710          Exit Sub
7720      End If

7730      For X = 0 To g.Cols - 1
7740          g.Col = X
7750          g.CellBackColor = vbYellow
7760      Next

7770      bMoveUp.Enabled = True
7780      bMoveDown.Enabled = True

End Sub


