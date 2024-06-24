VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestCodeMapping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Biochemistry Test Code Mapping"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   8640
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   8220
      Picture         =   "frmTestCodeMapping.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   1100
      Left            =   8220
      Picture         =   "frmTestCodeMapping.frx":049C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cmbAnalyser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2835
   End
   Begin VB.ComboBox cmbTestName 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox cmbTestCode 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6060
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4575
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8070
      _Version        =   393216
      RowHeightMin    =   315
      ScrollTrack     =   -1  'True
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSelectAnalyser 
      AutoSize        =   -1  'True
      Caption         =   "Select Analyser"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   1095
   End
End
Attribute VB_Name = "frmTestCodeMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrevAnalyser As String
Private grd As MSFlexGrid
Private m_sDiscipline As String

Public Property Get Discipline() As String

6450      Discipline = m_sDiscipline

End Property

Public Property Let Discipline(ByVal sDiscipline As String)

6460      m_sDiscipline = sDiscipline

End Property


Private Sub InitializeGrid()
      Dim i As Integer
6470  With g
6480      .Rows = 2: .FixedRows = 1
6490      .Cols = 4: .FixedCols = 0
6500      .Rows = 1
          '.Font.Size = 10         'fgcFontSize
          '.Font.Name = fgcFontName
          '.ForeColor = fgcForeColor
          '.BackColor = fgcBackColor
          '.ForeColorFixed = fgcForeColorFixed
          '.BackColorFixed = fgcBackColorFixed
6510      .ScrollBars = flexScrollBarBoth
          'Name                                                                      |Code
6520      .TextMatrix(0, 0) = "Code": .ColWidth(0) = 1500: .ColAlignment(0) = flexAlignLeftCenter
6530      .TextMatrix(0, 1) = "Test Name": .ColWidth(1) = 3000: .ColAlignment(1) = flexAlignLeftCenter
6540      .TextMatrix(0, 2) = "Short Name": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
6550      .TextMatrix(0, 3) = "Analyser Code": .ColWidth(3) = 1500: .ColAlignment(3) = flexAlignLeftCenter
6560      For i = 0 To .Cols - 1
6570          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
6580              .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + 150    'fgcExtraSpace
6590          End If
6600      Next i
6610  End With
End Sub

Private Sub FillGrid()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

6620  On Error GoTo FillGrid_Error

6630  sql = "Select Distinct D.Code, D.LongName, D.ShortName, M.EquipmentAnalyserCode " & _
            "From " & Discipline & "TestDefinitions D Left Outer Join " & _
            "(Select * From AnalyserTestCodeMapping Where AnalyserName = '" & cmbAnalyser & "' AND Department = '" & Discipline & "') M " & _
            "On D.Code = M.NetAcquireTestCode " & _
            "Order By M.EquipmentAnalyserCode"
6640  Set tb = New Recordset
6650  RecOpenClient 0, tb, sql
6660  cmdSave.Visible = False
6670  InitializeGrid
6680  If Not tb.EOF Then
6690      While Not tb.EOF
6700          s = tb!Code & "" & vbTab & _
                  tb!LongName & "" & vbTab & _
                  tb!ShortName & "" & vbTab & _
                  tb!EquipmentAnalyserCode
6710          g.AddItem s
6720          tb.MoveNext
6730      Wend

6740  End If
6750  g.TextMatrix(0, 3) = cmbAnalyser & " Code"

6760  Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

6770  intEL = Erl
6780  strES = Err.Description
6790  LogError "frmTestCodeMapping", "FillGrid", intEL, strES

End Sub

Private Sub cmbAnalyser_Click()

6800  If cmdSave.Visible = True Then
6810      If iMsg("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then
6820          Save PrevAnalyser
6830      End If
6840  End If

6850  If cmbAnalyser <> "" Then
6860      FillGrid
6870  End If
End Sub

Private Sub cmbAnalyser_DropDown()
6880  PrevAnalyser = cmbAnalyser
End Sub

Private Sub cmbTestCode_Click()
      Dim Code As String
6890  Code = cmbTestCode.Text
6900  g.TextMatrix(g.row, g.Col) = Code
6910  cmbTestCode.Visible = False
6920  g.TextMatrix(g.row, 1) = LongNamebyCode(Code, "Bio")
End Sub

Private Sub cmbTestName_Click()
      Dim LongName As String
6930  LongName = cmbTestName.Text
6940  g.TextMatrix(g.row, g.Col) = LongName
6950  cmbTestName.Visible = False
6960  g.TextMatrix(g.row, 0) = CodebyLongName(LongName, "Bio")
End Sub

Private Sub cmdCancel_Click()
6970  Unload Me
End Sub

Private Sub Save(AnalyserName As String)

      Dim tb As Recordset
      Dim sql As String
      Dim i As Integer

6980  On Error GoTo Save_Error

6990  For i = 1 To g.Rows - 1
7000      If g.TextMatrix(i, 3) = "" Then
7010          sql = "Delete From AnalyserTestCodeMapping " & _
                    "Where AnalyserName = '" & AnalyserName & "' " & _
                    "And NetAcquireTestCode = '" & g.TextMatrix(i, 0) & "' " & _
                    "AND Department = '" & Discipline & "'"
7020          Cnxn(0).Execute sql
7030      Else
7040          sql = "Select * From AnalyserTestCodeMapping " & _
                    "Where AnalyserName = '" & AnalyserName & "' " & _
                    "And NetAcquireTestCode = '" & g.TextMatrix(i, 0) & "' " & _
                    "AND Department = '" & Discipline & "'"

7050          Set tb = New Recordset
7060          RecOpenClient 0, tb, sql
7070          If tb.EOF Then
7080              tb.AddNew
7090          End If

7100          tb!NetAcquireTestCode = g.TextMatrix(i, 0)
7110          tb!EquipmentAnalyserCode = g.TextMatrix(i, 3)
7120          tb!TestName = g.TextMatrix(i, 1)
7130          tb!AnalyserName = AnalyserName
7140          tb!Department = Discipline
7150          tb!DateTimeOfRecord = Format(Now, "YYYY-MM-DD hh:mm:ss")

7160          tb.Update
7170      End If

7180  Next i

7190  Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

7200  intEL = Erl
7210  strES = Err.Description
7220  LogError "frmTestCodeMapping", "Save", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()



7230  On Error GoTo cmdSave_Click_Error

7240  Save cmbAnalyser
7250  FillGrid
7260  cmdSave.Visible = False
7270  Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

7280  intEL = Erl
7290  strES = Err.Description
7300  LogError "frmTestCodeMapping", "cmdSave_Click", intEL, strES

End Sub

Private Sub Form_Load()
7310  InitializeGrid
7320  Me.Caption = "NetAcquire --- " & Discipline & " Test Code Mapping"
      'Fill all combos
7330  FillGenericList cmbAnalyser, "AC"

End Sub

Private Sub FillTestCombos()

      Dim tb As Recordset
      Dim sql As String

7340  On Error GoTo FillTestCombos_Error

7350  sql = "Select Distinct Code, LongName From BioTestDefinitions"
7360  Set tb = New Recordset
7370  RecOpenClient 0, tb, sql

7380  cmbTestCode.Clear

7390  cmbTestName.Clear
7400  If Not tb.EOF Then
7410      While Not tb.EOF
7420          cmbTestCode.AddItem tb!Code & ""
7430          cmbTestName.AddItem tb!LongName & ""
7440          tb.MoveNext

7450      Wend
7460  End If

7470  Exit Sub

FillTestCombos_Error:

      Dim strES As String
      Dim intEL As Integer

7480  intEL = Erl
7490  strES = Err.Description
7500  LogError "frmTestCodeMapping", "FillTestCombos", intEL, strES, sql

End Sub





Private Function LongNamebyCode(ByVal Code As String, _
                                ByVal Department As String) As String

      Dim tb As New Recordset
      Dim sql As String

7510  On Error GoTo LongNamebyCode_Error

7520  LongNamebyCode = "???"

7530  sql = "SELECT LongName FROM " & Department & "TestDefinitions WHERE " & _
            "Code = '" & Code & "'"

7540  Set tb = New Recordset
7550  RecOpenServer 0, tb, sql
7560  If Not tb.EOF Then
7570      LongNamebyCode = Trim(tb!LongName & "")
7580  End If

7590  Exit Function

LongNamebyCode_Error:

      Dim strES As String
      Dim intEL As Integer

7600  intEL = Erl
7610  strES = Err.Description
7620  LogError "basDisciplineFunctions", "LongNamebyCode", intEL, strES, sql

End Function

Private Function CodebyLongName(ByVal LongName As String, _
                                ByVal Department As String) As String

      Dim sql As String
      Dim tb As New Recordset

7630  On Error GoTo CodebyLongName_Error

7640  CodebyLongName = "???"

7650  sql = "SELECT Code FROM " & Department & "TestDefinitions WHERE " & _
            "LongName = '" & LongName & "'"
7660  Set tb = New Recordset
7670  RecOpenServer 0, tb, sql
7680  If Not tb.EOF Then
7690      CodebyLongName = Trim(tb!Code & "")
7700  End If

7710  Exit Function

CodebyLongName_Error:

      Dim strES As String
      Dim intEL As Integer

7720  intEL = Erl
7730  strES = Err.Description
7740  LogError "basDisciplineFunctions", "CodebyLongName", intEL, strES, sql

End Function




Private Sub g_Click()

      Static SortOrder As Boolean


7750  If g.MouseRow = 0 Then
7760      If SortOrder Then
7770          g.Sort = flexSortGenericAscending
7780      Else
7790          g.Sort = flexSortGenericDescending
7800      End If
7810      SortOrder = Not SortOrder
7820      Exit Sub
7830  End If

7840  If g.ColSel = 3 Then
7850      If g.MouseRow > 0 Then
7860          Set grd = g
7870          grd.row = grd.MouseRow
7880          grd.Col = grd.MouseCol
7890          LoadControls
7900      End If
7910      Exit Sub
7920  End If

End Sub

Private Sub g_KeyUp(KeyCode As Integer, Shift As Integer)
7930  If g.Col = 3 Then
7940      If EditGrid(g, KeyCode, Shift) Then
7950          cmdSave.Visible = True
7960      End If
7970  End If
End Sub

Private Sub g_LeaveCell()

7980  On Error GoTo g_LeaveCell_Error

7990  txtText.Visible = False

8000  Exit Sub

g_LeaveCell_Error:

      Dim strES As String
      Dim intEL As Integer

8010  intEL = Erl
8020  strES = Err.Description
8030  LogError "frmTestCodeMapping", "g_LeaveCell", intEL, strES

End Sub

Private Sub g_Scroll()

8040  On Error GoTo g_Scroll_Error

8050  txtText.Visible = False

8060  Exit Sub

g_Scroll_Error:

      Dim strES As String
      Dim intEL As Integer

8070  intEL = Erl
8080  strES = Err.Description
8090  LogError "frmTestCodeMapping", "g_Scroll", intEL, strES

End Sub

Private Sub txtText_LostFocus()

8100  On Error GoTo txtText_LostFocus_Error



8110  txtText.Visible = False

8120  Exit Sub

txtText_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

8130  intEL = Erl
8140  strES = Err.Description
8150  LogError "frmTestCodeMapping", "txtText_LostFocus", intEL, strES

End Sub

Private Sub LoadControls()
8160  On Error GoTo LoadControls_Error

8170  txtText.Visible = False
8180  txtText = ""
      'gRD.SetFocus

8190  Select Case grd.Col
          Case 3:
8200          txtText.Move grd.Left + grd.CellLeft + 5, _
                           grd.Top + grd.CellTop + 5, _
                           grd.CellWidth - 20, grd.CellHeight - 20
8210          txtText.Text = grd.TextMatrix(grd.row, grd.Col)
8220          txtText.Visible = True
8230          txtText.SelStart = 0
8240          txtText.SelLength = Len(txtText)
8250          txtText.SetFocus

8260  End Select

8270  Exit Sub

LoadControls_Error:

      Dim strES As String
      Dim intEL As Integer

8280  intEL = Erl
8290  strES = Err.Description
8300  LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
8310  If KeyCode = vbKeyUp Then
          'GoOneRowUp
8320  ElseIf KeyCode = vbKeyDown Then
          'GoOneRowDown
8330  ElseIf KeyCode = 13 Then
8340      txtText.Visible = False
8350  Else
8360      grd.TextMatrix(grd.row, grd.Col) = txtText
8370      cmdSave.Visible = True
8380  End If
End Sub

Private Sub GoOneRowUp()
8390  If grd.row > 1 Then
8400      grd.row = grd.row - 1
8410      LoadControls
8420  End If
End Sub
Private Sub GoOneRowDown()
8430  If grd.row < grd.Rows - 1 Then
8440      grd.row = grd.row + 1
8450      LoadControls
8460  End If
End Sub



