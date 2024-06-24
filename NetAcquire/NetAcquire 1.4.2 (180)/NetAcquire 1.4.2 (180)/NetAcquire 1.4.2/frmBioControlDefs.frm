VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBioControlDefs 
   Caption         =   "NetAcquire - Biochemistry Controls"
   ClientHeight    =   6420
   ClientLeft      =   555
   ClientTop       =   840
   ClientWidth     =   8415
   LinkTopic       =   "Form2"
   ScaleHeight     =   6420
   ScaleWidth      =   8415
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   7230
      Picture         =   "frmBioControlDefs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1035
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   7230
      Picture         =   "frmBioControlDefs.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3630
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5805
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10239
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "Test Name                          |<Low    |<High   |<Low    |<High   |<Low    |<High    "
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5430
      TabIndex        =   5
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4050
      TabIndex        =   4
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2700
      TabIndex        =   3
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmBioControlDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

6510      On Error GoTo FillG_Error

6520      g.Rows = 2
6530      g.AddItem ""
6540      g.RemoveItem 1

6550      sql = "Select distinct Shortname, PrintPriority, " & _
              "LControlLow, LControlHigh, " & _
              "NControlLow, NControlHigh, " & _
              "HControlHigh, HControlLow " & _
              "from BioTestDefinitions " & _
              "Order by PrintPriority"
6560      Set tb = New Recordset
6570      RecOpenServer 0, tb, sql
6580      Do While Not tb.EOF
6590          g.AddItem tb!ShortName & vbTab & _
                  tb!LControlLow & vbTab & _
                  tb!LControlHigh & vbTab & _
                  tb!NControlLow & vbTab & _
                  tb!NControlHigh & vbTab & _
                  tb!HControlHigh & vbTab & _
                  tb!HControlLow & ""
6600          tb.MoveNext
6610      Loop

6620      If g.Rows > 2 Then g.RemoveItem 1

6630      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

6640      intEL = Erl
6650      strES = Err.Description
6660      LogError "fBioControlDefs", "FillG", intEL, strES, sql


End Sub
Private Sub SaveG()

          Dim sql As String

6670      On Error GoTo SaveG_Error

6680      sql = "Update BioTestDefinitions " & _
              "Set LControlLow = '" & Val(g.TextMatrix(g.row, 1)) & "', " & _
              "LControlHigh = '" & Val(g.TextMatrix(g.row, 2)) & "', " & _
              "NControlLow = '" & Val(g.TextMatrix(g.row, 3)) & "', " & _
              "NControlHigh = '" & Val(g.TextMatrix(g.row, 4)) & "', " & _
              "HControlLow = '" & Val(g.TextMatrix(g.row, 5)) & "', " & _
              "HControlHigh = '" & Val(g.TextMatrix(g.row, 6)) & "' " & _
              "where ShortName = '" & g.TextMatrix(g.row, 0) & "'"
6690      Cnxn(0).Execute sql

6700      Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer

6710      intEL = Erl
6720      strES = Err.Description
6730      LogError "fBioControlDefs", "SaveG", intEL, strES, sql


End Sub




Private Sub bcancel_Click()

6740      Unload Me

End Sub


Private Sub bPrint_Click()

          Dim Y As Integer
          Dim X As Integer

6750      Screen.MousePointer = 11

6760      Printer.Print
6770      Printer.Font.Name = "Courier New"
6780      Printer.Font.size = 12

6790      Printer.Print "List of Control Limits."
6800      Printer.Print

6810      For Y = 0 To g.Rows - 1
6820          g.row = Y
6830          g.Col = 0
6840          Printer.Print g; Tab(20);
6850          For X = 1 To 6
6860              g.Col = X
6870              Printer.Print Left$(g & "       ", 7); " ";
6880          Next
6890          Printer.Print
6900      Next

6910      Printer.EndDoc

6920      Screen.MousePointer = 0

End Sub


Private Sub Form_Load()

6930      g.Font.Bold = True

6940      FillG

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean

6950      If g.MouseRow = 0 Then
6960          If SortOrder Then
6970              g.Sort = flexSortGenericAscending
6980          Else
6990              g.Sort = flexSortGenericDescending
7000          End If
7010          SortOrder = Not SortOrder
7020          Exit Sub
7030      End If

7040      If g.Col = 0 Then
7050          Exit Sub
7060      ElseIf g.Col = 1 Then
7070          g.Enabled = False
7080          g = iBOX("Low Control Low Limit?", , g)
7090          SaveG
7100          g.Enabled = True
7110      ElseIf g.Col = 2 Then
7120          g.Enabled = False
7130          g = iBOX("Low Control High Limit?", , g)
7140          SaveG
7150          g.Enabled = True
7160      ElseIf g.Col = 3 Then
7170          g.Enabled = False
7180          g = iBOX("Normal Control Low Limit?", , g)
7190          SaveG
7200          g.Enabled = True
7210      ElseIf g.Col = 4 Then
7220          g.Enabled = False
7230          g = iBOX("Normal Control High Limit?", , g)
7240          SaveG
7250          g.Enabled = True
7260      ElseIf g.Col = 5 Then
7270          g.Enabled = False
7280          g = iBOX("High Control Low Limit?", , g)
7290          SaveG
7300          g.Enabled = True
7310      ElseIf g.Col = 6 Then
7320          g.Enabled = False
7330          g = iBOX("High Control High Limit?", , g)
7340          SaveG
7350          g.Enabled = True
7360      End If

End Sub


