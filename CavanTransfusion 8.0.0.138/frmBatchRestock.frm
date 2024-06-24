VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchRestock 
   Caption         =   "NetAcquire - Batch Restock (Single Product)"
   ClientHeight    =   5460
   ClientLeft      =   1395
   ClientTop       =   1215
   ClientWidth     =   10800
   Icon            =   "frmBatchRestock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10800
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3675
      Left            =   210
      TabIndex        =   8
      Top             =   1140
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6482
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmBatchRestock.frx":08CA
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   9480
      Picture         =   "frmBatchRestock.frx":0983
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3750
      Width           =   1005
   End
   Begin VB.CommandButton btnRestock 
      Caption         =   "Re&Stock"
      Height          =   705
      Left            =   9450
      Picture         =   "frmBatchRestock.frx":184D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1005
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "&Remove"
      Height          =   705
      Left            =   8160
      Picture         =   "frmBatchRestock.frx":1C8F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   330
      Width           =   1005
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   420
      Width           =   2325
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   7
      Top             =   4980
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2040
      Picture         =   "frmBatchRestock.frx":20D1
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "To log all entries back into stock, click 'Restock'"
      Height          =   645
      Left            =   9360
      TabIndex        =   5
      Top             =   1590
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "To remove a unit from the list, highlight the unit then click 'Remove'"
      Height          =   465
      Left            =   5490
      TabIndex        =   4
      Top             =   450
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan or enter each Unit to be Restocked"
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1845
   End
End
Attribute VB_Name = "frmBatchRestock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()

10  Unload Me

End Sub


Private Sub btnRemove_Click()

    Dim Y As Integer

10  g.Col = 0
20  For Y = 0 To g.Rows - 1
30      g.Row = Y
40      If g.CellBackColor = vbYellow Then
50          If g.Rows = 2 Then
60              g.AddItem ""
70          End If
80          g.RemoveItem Y
90          Exit For
100     End If
110 Next

End Sub

Private Sub btnRestock_Click()

    Dim n As Integer
    Dim BarCode As String
    Dim sql As String
    Dim tD As Recordset
    Dim Ward As String
    Dim DoB As String
    Dim Typenex As String
    Dim DateTimeNow As String
    Dim Ps As Products
    Dim p As Product

10  On Error GoTo btnRestock_Click_Error

20  DateTimeNow = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

30  For n = g.Rows - 1 To 1 Step -1

40      Set Ps = New Products
50      Ps.LoadLatestISBT128 g.TextMatrix(n, 0), ProductBarCodeFor(g.TextMatrix(n, 1))
60      If Ps.Count = 1 Then
70          Set p = Ps(1)
80          BarCode = p.BarCode

90          sql = "Select Ward, Dob, Typenex from PatientDetails where " & _
                  "Labnumber = '" & p.SampleID & "'"
100         Set tD = New Recordset
110         RecOpenClientBB 0, tD, sql
120         If Not tD.EOF Then
130             Ward = tD!Ward & ""
140             DoB = tD!DoB & ""
150             Typenex = tD!Typenex & ""
160         End If

170         sql = "Insert into Reclaimed " & _
                  "( Name, Chart, Unit, [Group], Product, xmDate, " & _
                "  DateTime, Operator, Ward, DoB, Typenex ) VALUES " & _
                  "( '" & AddTicks(p.PatName) & "', " & _
                "  '" & p.Chart & "', " & _
                "  '" & p.ISBT128 & "', " & _
                "  '" & Bar2Group(p.GroupRh) & "', " & _
                "  '" & ProductWordingFor(p.BarCode) & "', " & _
                "  '" & Format(p.RecordDateTime, "dd/MMM/yyyy HH:nn:ss") & "', " & _
                "  '" & DateTimeNow & "', " & _
                "  '" & UserCode & "', " & _
                "  '" & Ward & "', " & _
                "  '" & Format(DoB, "dd/MMM/yyyy") & "', " & _
                "  '" & Typenex & "' )"
180         CnxnBB(0).Execute sql

190         p.PackEvent = "R"
200         p.Chart = ""
210         p.PatName = ""
220         p.UserName = UserCode
230         p.RecordDateTime = DateTimeNow
240         p.Save

250     End If
260     If g.Rows = 2 Then
270         g.AddItem ""
280     End If
290     g.RemoveItem n
300 Next

310 Exit Sub

btnRestock_Click_Error:

    Dim strES As String
    Dim intEL As Integer

320 intEL = Erl
330 strES = Err.Description
340 LogError "frmBatchRestock", "btnRestock_Click", intEL, strES, sql

End Sub


Private Sub Form_Load()

10  g.Rows = 2
20  g.AddItem ""
30  g.RemoveItem 1

End Sub


Private Sub g_Click()

    Dim ySave As Integer
    Dim Y As Integer
    Dim X As Integer

10  If g.MouseRow = 0 Then Exit Sub

20  ySave = g.Row
30  g.Col = 0
40  For Y = 1 To g.Rows - 1
50      g.Row = Y
60      If g.CellBackColor = vbYellow Then
70          For X = 0 To g.Cols - 1
80              g.Col = X
90              g.CellBackColor = 0
100         Next
110     End If
120 Next
130 g.Row = ySave
140 For X = 0 To g.Cols - 1
150     g.Col = X
160     g.CellBackColor = vbYellow
170 Next

End Sub

Private Sub txtInput_LostFocus()

    Dim Ps As New Products
    Dim p As Product
    Dim s As String
    Dim f As Form

10  On Error GoTo txtInput_LostFocus_Error

20  If Len(Trim$(txtInput)) > 0 Then
30      txtInput = UCase$(txtInput)

40      If Left$(txtInput, 1) = "=" Then
50          s = ISOmod37_2(Mid$(txtInput, 2, 13))
60          txtInput = Mid$(txtInput, 2, 13) & " " & s
70      End If


80      Ps.LoadLatestByUnitNumberISBT128 txtInput
90      If Ps.Count = 0 Then
100         iMsg "Unit not found", vbInformation + vbOKOnly
110         If TimedOut Then Unload Me: Exit Sub
120     ElseIf Ps.Count > 1 Then    'multiple products found
130         Set f = New frmSelectFromMultiple
140         f.ProductList = Ps
150         f.Show 1
160         Set p = f.SelectedProduct
170         Unload f
180         Set f = Nothing
190     Else
200         Set p = Ps.Item(1)
210     End If

220     If Not p Is Nothing Then
230         s = p.ISBT128 & vbTab & _
                ProductWordingFor(p.BarCode) & vbTab & _
                p.DateExpiry & vbTab & _
                SupplierNameFor(p.Supplier)
240         g.AddItem s

250         If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
260             g.RemoveItem 1
270         End If

280         txtInput = ""
290         txtInput.SetFocus
300     End If
310 End If

320 Exit Sub

txtInput_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

330 intEL = Erl
340 strES = Err.Description
350 LogError "frmBatchRestock", "txtInput_LostFocus", intEL, strES

End Sub


