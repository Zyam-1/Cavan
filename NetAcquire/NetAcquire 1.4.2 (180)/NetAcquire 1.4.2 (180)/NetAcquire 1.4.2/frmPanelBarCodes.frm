VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPanelBarCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Panel Bar Codes"
   ClientHeight    =   5910
   ClientLeft      =   1650
   ClientTop       =   780
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cSampleType 
      Height          =   315
      Left            =   4650
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   1845
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   4950
      TabIndex        =   0
      Top             =   1770
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   9604
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Panel Name                |<Bar Code                "
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
End
Attribute VB_Name = "frmPanelBarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

      Dim strPanelType As String
      Dim tb As Recordset
      Dim tx As Recordset
      Dim sql As String
      Dim BarCode As String

20440 On Error GoTo FillG_Error

20450 strPanelType = ListCodeFor("ST", cSampleType)

20460 g.Visible = False
20470 g.Rows = 2
20480 g.AddItem ""
20490 g.RemoveItem 1

20500 sql = "Select distinct PanelName, ListOrder from Panels " & _
            "Where PanelType = '" & strPanelType & "' " & _
            "Order by ListOrder"
20510 Set tb = New Recordset
20520 RecOpenServer 0, tb, sql

20530 Do While Not tb.EOF
20540   sql = "Select BarCode from Panels where PanelName = '" & tb!PanelName & "'"
20550   Set tx = New Recordset
20560   RecOpenServer 0, tx, sql
20570   If Not tx.EOF Then
20580     BarCode = tx!BarCode & ""
20590   Else
20600     BarCode = ""
20610   End If
20620   g.AddItem tb!PanelName & vbTab & BarCode
20630   tb.MoveNext
20640 Loop

20650 If g.Rows > 2 Then
20660   g.RemoveItem 1
20670 End If

20680 g.Visible = True

20690 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

20700 intEL = Erl
20710 strES = Err.Description
20720 LogError "fPanelBarCodes", "FillG", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

20730 Unload Me

End Sub


Private Sub cSampleType_Click()

20740 FillG

End Sub


Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

20750 On Error GoTo Form_Load_Error

20760 cSampleType.Clear

20770 sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "order by ListOrder"
20780 Set tb = New Recordset
20790 RecOpenServer 0, tb, sql
20800 Do While Not tb.EOF
20810   cSampleType.AddItem tb!Text & ""
20820   tb.MoveNext
20830 Loop
20840 If cSampleType.ListCount > 0 Then
20850   cSampleType.ListIndex = 0
20860 End If

20870 FillG

20880 Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

20890 intEL = Erl
20900 strES = Err.Description
20910 LogError "fPanelBarCodes", "Form_Load", intEL, strES, sql


End Sub


Private Sub g_Click()

      Static SortOrder As Boolean
      Dim PanelName As String
      Dim BarCode As String
      Dim PanelType As String
      Dim sql As String

20920 On Error GoTo g_Click_Error

20930 If g.MouseRow = 0 Then
20940   If SortOrder Then
20950     g.Sort = flexSortGenericAscending
20960   Else
20970     g.Sort = flexSortGenericDescending
20980   End If
20990   SortOrder = Not SortOrder
21000   Exit Sub
21010 End If

21020 PanelType = ListCodeFor("ST", cSampleType)

21030 If g.Col = 1 Then
21040   g.Enabled = False
21050   g = iBOX("Enter Bar Code", , g)
21060   g.Enabled = True
21070   PanelName = g.TextMatrix(g.row, 0)
21080   BarCode = g.TextMatrix(g.row, 1)
21090   sql = "Update Panels set BarCode = '" & BarCode & "' " & _
              "where PanelName = '" & PanelName & "' " & _
              "and PanelType = '" & PanelType & "' " & _
              "and Hospital = '" & HospName(0) & "'"
21100   Cnxn(0).Execute sql
21110 End If

21120 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

21130 intEL = Erl
21140 strES = Err.Description
21150 LogError "fPanelBarCodes", "g_Click", intEL, strES, sql


End Sub


