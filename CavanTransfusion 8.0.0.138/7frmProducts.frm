VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Products"
   ClientHeight    =   8268
   ClientLeft      =   396
   ClientTop       =   660
   ClientWidth     =   9840
   ForeColor       =   &H00C0C0C0&
   Icon            =   "7frmProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8268
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   800
      Left            =   7500
      Picture         =   "7frmProducts.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   800
      Left            =   6420
      Picture         =   "7frmProducts.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   6420
      Picture         =   "7frmProducts.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   800
      Left            =   7500
      Picture         =   "7frmProducts.frx":18A8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Product"
      Height          =   1725
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   5985
      Begin VB.OptionButton o 
         Caption         =   "LG Octaplas"
         Height          =   195
         Index           =   5
         Left            =   3180
         TabIndex        =   18
         Top             =   1410
         Width           =   1395
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   4685
         TabIndex        =   11
         Top             =   1200
         Width           =   1000
      End
      Begin VB.OptionButton o 
         Caption         =   "Cryoprecipitate"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   10
         Top             =   1410
         Width           =   1395
      End
      Begin VB.OptionButton o 
         Caption         =   "Platelets"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   9
         Top             =   1410
         Width           =   945
      End
      Begin VB.OptionButton o 
         Caption         =   "Plasma"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   8
         Top             =   1170
         Width           =   885
      End
      Begin VB.OptionButton o 
         Caption         =   "Red Cells"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   1170
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton o 
         Caption         =   "Whole Blood"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox tBarCode 
         Height          =   285
         Left            =   870
         MaxLength       =   10
         TabIndex        =   5
         Top             =   660
         Width           =   1365
      End
      Begin VB.TextBox tWording 
         Height          =   285
         Left            =   870
         MaxLength       =   50
         TabIndex        =   4
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BarCode"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Wording"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   330
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5895
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Drag to set new List Order"
      Top             =   1950
      Width           =   9465
      _ExtentX        =   16701
      _ExtentY        =   10393
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   $"7frmProducts.frx":1F12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   15
      Top             =   7980
      Width           =   9435
      _ExtentX        =   16637
      _ExtentY        =   296
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
      Left            =   8520
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Temp As String

Private Sub SaveG()

      Dim Y As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveG_Error

20    For Y = 1 To g.Rows - 1
30        g.row = Y
40        If g.CellBackColor = vbRed Then
50            CnxnBB(0).Execute "Delete From ProductList Where BarCode='" & tBarCode & "'"
60        Else
70            sql = "Select * from ProductList where " & _
                    "BarCode = '" & g.TextMatrix(Y, 1) & "'"
80            Set tb = New Recordset
90            RecOpenClientBB 0, tb, sql
100           If tb.EOF Then tb.AddNew
110           With tb
120             !Wording = g.TextMatrix(Y, 0)
130             !BarCode = g.TextMatrix(Y, 1)
140             !ListOrder = Y
150             !Batch = False
160             !Generic = g.TextMatrix(Y, 2)
170             !InUse = 1
180             .Update
190           End With
200       End If
210   Next

220   FillG

230   tWording = ""
240   tBarCode = ""
250   tBarCode.SetFocus
260   cmdSave.Visible = False

270   Exit Sub

SaveG_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmProducts", "SaveG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

      Dim Generic As String
      Dim n As Integer
  
10    If Trim$(tWording) = "" Or Trim$(tBarCode) = "" Then
20      Exit Sub
30    End If

40    For n = 0 To 5
50      If o(n) Then
60        Generic = o(n).Caption
70        Exit For
80      End If
90    Next

      Dim boolItemFound As Boolean
100   boolItemFound = False
      Dim X As Integer
110   For X = 1 To g.Rows - 1
120       If tBarCode = g.TextMatrix(X, 1) Then
130           boolItemFound = True 'item found
140           Exit For
    
150       End If
160   Next X
170   If boolItemFound Then
180       Call MarkGridRow(g, g.row, 0, vbBlack, False, True, False)
190       g.TextMatrix(g.row, 0) = tWording
200       g.TextMatrix(g.row, 1) = tBarCode
210       g.TextMatrix(g.row, 2) = Generic
220   Else
230       g.AddItem tWording & vbTab & tBarCode & vbTab & Generic
240   End If

250   tWording = ""
260   tBarCode = ""

      'SaveG
      'FillG
270   cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 9
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "Product List"

      '****Report body

90    For i = 1 To 108
100       Printer.Print "_";
110   Next i
120   Printer.Print
130   Printer.Print FormatString("Wording", 70, "|");
140   Printer.Print FormatString("Bar Code", 15, "|");
150   Printer.Print FormatString("Generic", 20, "|")
160   Printer.Font.Bold = False
170   For i = 1 To 108
180       Printer.Print "-";
190   Next i
200   Printer.Print
210   For Y = 1 To g.Rows - 1
220       Printer.Print FormatString(g.TextMatrix(Y, 0), 70, "|");
230       Printer.Print FormatString(g.TextMatrix(Y, 1), 15, "|");
240       Printer.Print FormatString(g.TextMatrix(Y, 2), 20, "|")
250   Next


260   Printer.EndDoc

270   For Each Px In Printers
280     If Px.DeviceName = OriginalPrinter Then
290       Set Printer = Px
300       Exit For
310     End If
320   Next

End Sub

Private Sub cmdSave_Click()

10    SaveG
20    FillG
End Sub



Sub FillG()
  
      Dim s As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from ProductList order by ListOrder"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!Wording & vbTab & _
            tb!BarCode & vbTab & _
            tb!Generic
100     g.AddItem s
110     tb.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmProducts", "FillG", intEL, strES, sql


End Sub



Private Sub cmdXL_Click()
       Dim strHeading As String

10    strHeading = "Product List" & vbCr
20    strHeading = strHeading & " " & vbCr
30    ExportFlexGrid g, Me, strHeading
End Sub

Private Sub Form_Load()

10    g.Font.Bold = True

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
20        FillG
      '**************************************

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Visible Then
30      Answer = iMsg("Cancel without Saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70        Exit Sub
80      End If
90    End If

End Sub



Private Sub g_Click()

      Static SortOrder As Boolean

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If SortOrder Then
40        g.Sort = flexSortGenericAscending
50      Else
60        g.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90      Exit Sub
100   End If
110   If g.col = 0 Then
120       g.Enabled = False
130       If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then
140           tWording = g.TextMatrix(g.row, 0)
150           tBarCode = g.TextMatrix(g.row, 1)
160           Select Case g.TextMatrix(g.row, 2)
                  Case "Whole Blood"
170                   o(0).Value = True
180               Case "Red Cells"
190                   o(1).Value = True
200               Case "Plasma"
210                   o(2).Value = True
220               Case "Platelets"
230                   o(3).Value = True
240               Case "Cryoprecipitate"
250                   o(4).Value = True

260           End Select
270           Call MarkGridRow(g, g.row, vbRed, vbYellow, True, True, True)
    
280       End If
290       g.Enabled = True
300   End If

310   Exit Sub

g_Click_Error:

Dim strES As String
Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmProducts", "g_Click", intEL, strES

End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim n As Integer
      Static PrevY As Integer

10    If Button = vbLeftButton And g.MouseRow > 0 Then
20      If Temp = "" Then
30        PrevY = g.MouseRow
40        For n = 0 To g.Cols - 1
50          Temp = Temp & g.TextMatrix(g.row, n) & vbTab
60        Next
70        Temp = Left$(Temp, Len(Temp) - 1)
80        Exit Sub
90      Else
100       If g.MouseRow <> PrevY Then
110         g.RemoveItem PrevY
120         If g.MouseRow <> PrevY Then
130           g.AddItem Temp, g.MouseRow
140           PrevY = g.MouseRow
150         Else
160           g.AddItem Temp
170           PrevY = g.Rows - 1
180         End If
190       End If
200     End If
210   End If

End Sub


Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    cmdSave.Visible = True
20    Temp = ""

End Sub




Private Sub tBarCode_Change()

10    cmdSave.Visible = Trim$(tWording) <> "" And Trim$(tBarCode) <> ""

End Sub

Private Sub tWording_Change()

10    cmdSave.Visible = Trim$(tWording) <> "" And Trim$(tBarCode) <> ""

End Sub


