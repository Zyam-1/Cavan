VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocklist"
   ClientHeight    =   8850
   ClientLeft      =   285
   ClientTop       =   405
   ClientWidth     =   14685
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
   HelpContextID   =   10000
   Icon            =   "7frmStockList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8850
   ScaleWidth      =   14685
   Begin VB.Frame Frame1 
      Caption         =   "Expiry within"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   9840
      TabIndex        =   10
      Top             =   150
      Width           =   3165
      Begin VB.OptionButton oTimeLimit 
         Alignment       =   1  'Right Justify
         Caption         =   "One Week"
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
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton oTimeLimit 
         Alignment       =   1  'Right Justify
         Caption         =   "One Month"
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
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   555
         Width           =   1095
      End
      Begin VB.OptionButton oTimeLimit 
         Caption         =   "No Limit"
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
         Index           =   2
         Left            =   1500
         TabIndex        =   12
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox cinclude 
         Caption         =   "Include Expired"
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
         HelpContextID   =   10020
         Left            =   1500
         TabIndex        =   11
         Top             =   540
         Width           =   1455
      End
   End
   Begin VB.CommandButton bLookUp 
      Caption         =   "&Look Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   13545
      Picture         =   "7frmStockList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   1035
   End
   Begin VB.CommandButton bModify 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   13545
      Picture         =   "7frmStockList.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1710
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5325
      Left            =   -15
      TabIndex        =   4
      Top             =   1080
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   9393
      _Version        =   393216
      Cols            =   7
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
      MergeCells      =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmStockList.frx":1BD6
   End
   Begin VB.ComboBox lstproduct 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   5115
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
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
      Height          =   1065
      HelpContextID   =   10080
      Left            =   13545
      Picture         =   "7frmStockList.frx":1CA0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3990
      Width           =   1035
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
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
      Height          =   735
      HelpContextID   =   10070
      Left            =   7170
      Picture         =   "7frmStockList.frx":2B6A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
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
      Height          =   1065
      HelpContextID   =   10090
      Left            =   13545
      Picture         =   "7frmStockList.frx":2FF5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7410
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grdBat 
      Height          =   1575
      Left            =   -15
      TabIndex        =   7
      Top             =   6945
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   6
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
      AllowUserResizing=   1
      FormatString    =   $"7frmStockList.frx":3EBF
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   -15
      TabIndex        =   9
      Top             =   6420
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Batches"
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
      Left            =   -15
      TabIndex        =   8
      Top             =   6690
      Width           =   585
   End
End
Attribute VB_Name = "frmStockList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillG()

      Dim sn As Recordset
      Dim Sql As String
      Dim s As String
      Dim limit As String
      Dim Prod As String
      Dim n As Integer
      Dim Found As Integer
      Dim Org As String
Dim r As Integer

10    On Error GoTo FillG_Error

20    If Trim$(lstproduct) = "" Then
30      iMsg "Specify Product", vbCritical
40      If TimedOut Then Unload Me: Exit Sub
50      Exit Sub
60    End If

70    g.Rows = 2
80    g.AddItem ""
90    g.RemoveItem 1

100   Found = False
110   Org = lstproduct
120   If lstproduct = "All" Then
130     For n = 1 To lstproduct.ListCount - 1
140       lstproduct.ListIndex = n
150       GoSub searchforit
160     Next
170   Else
180     GoSub searchforit
190   End If

200   If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
210     g.RemoveItem 1
220   End If

230   If Not Found Then
240     iMsg "None found.", vbInformation
250     If TimedOut Then Unload Me: Exit Sub
260   End If

270   lstproduct = Org

280   Load_Batch

290   Exit Sub

searchforit:
300   Prod = " and (barcode = '" & ProductBarCodeFor(lstproduct) & "' )"

310   limit = " AND dateexpiry <= '"
320   If oTimeLimit(0) Then 'one week
330     limit = limit & Format(Date + 7, "dd/mmm/yyyy") & "'"
340   ElseIf oTimeLimit(1) Then 'one month
350     limit = limit & Format(Date + 31, "dd/mmm/yyyy") & "'"
360   Else    'no limit
370     limit = ""
380   End If

390   Sql = "SELECT * FROM latest WHERE " & _
            "Event IN ( 'C', 'R', " & _
            " 'X', 'P', 'K')"
400   If cinclude = 0 Then
410     Sql = Sql & " and dateexpiry >= '" & Format(Now, "dd/mmm/yyyy HH:mm") & "'"
420   End If
430   Sql = Sql & limit & Prod & _
            " order by dateexpiry"

440   Set sn = New Recordset
450   RecOpenClientBB 0, sn, Sql
460   If Not sn.EOF Then
470     Found = True
480     g.AddItem lstproduct & vbTab & lstproduct & vbTab & lstproduct & vbTab & lstproduct & vbTab & lstproduct & vbTab & lstproduct & vbTab & lstproduct
490   r = g.Rows - 1
500   g.MergeRow(r) = True

510     g.Col = 0
520     g.Row = r
530     g.CellBackColor = vbCyan
540   End If
550   Do While Not sn.EOF
560     s = sn!ISBT128 & "" & vbTab & vbTab & _
            Bar2Group(sn!GroupRh & "") & vbTab
570     If QueryChecked(sn!BarCode, sn!ISBT128 & "") Then
580       s = s & "Yes"
590     End If
600     s = s & vbTab & Format(sn!DateExpiry, "dd/mm/yyyy HH:mm") & vbTab & _
            sn!Screen & vbTab & sn!Notes & ""
610     g.AddItem s
  
620     If TagIsPresent(sn!ISBT128 & "", CDate(sn!DateExpiry)) Then
630       g.Row = g.Rows - 1
640       g.Col = 1
650       g = "!"
660       g.CellBackColor = vbRed
670     End If
  
680     sn.MoveNext
690   Loop
700   Return

710   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

720   intEL = Erl
730   strES = Err.Description
740   LogError "frmStockList", "FillG", intEL, strES, Sql

End Sub

Private Sub Load_Batch()

      Dim sn As Recordset
      Dim Sql As String
      Dim s As String

10    On Error GoTo Load_Batch_Error

20    With grdBat
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60    End With

70    Sql = "Select * from batchproductlist where " & _
            "currentstock > 0 " & _
            "and dateexpiry >= '" & Format(Now, "dd/mmm/yyyy") & "'"
80    Set sn = New Recordset
90    RecOpenServerBB 0, sn, Sql

100   Do While Not sn.EOF
110     s = sn!BatchNumber & vbTab & sn!Group & vbTab & _
            sn!Product & vbTab & Format(sn!DateExpiry, "dd/mm/yyyy HH:mm") & vbTab & sn!CurrentStock
120     grdBat.AddItem s
130     s = ""
140     sn.MoveNext
150   Loop

160   If grdBat.Rows > 2 And grdBat.TextMatrix(1, 0) = "" Then grdBat.RemoveItem 1

170   Exit Sub

Load_Batch_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmStockList", "Load_Batch", intEL, strES, Sql


End Sub
Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bLookUp_Click()

      Dim RowSave As Integer
      Dim n As Integer

10    g.Col = 0
20    If Trim$(g) = "" Then Exit Sub
30    If g = "Unit" Then Exit Sub
40    If g.CellBackColor = vbCyan Then Exit Sub

50    RowSave = g.Row

60    With frmUnitHistory
70      .UnitNumber = g.TextMatrix(g.Row, 0)
80      .Expiry = g.TextMatrix(g.Row, 4)
90      For n = g.Row To 1 Step -1
100       g.Row = n
110       If g.CellBackColor = vbCyan Then
120         .ProductName = g
130         Exit For
140       End If
150     Next
160     .cmdSearch.Value = True
170     .Show 1
180   End With

190   g.Row = RowSave

End Sub


Private Sub bModify_Click()

      Dim s As String
      Dim Result As String
      Dim Sql As String
      Dim Unit As String
      Dim pc As String
      Dim Temp As String

10    On Error GoTo bModify_Click_Error

20    g.Col = 0
30    If Trim$(g) = "" Then Exit Sub
40    If g = "Unit" Then Exit Sub

50    g.Col = 0
60    Unit = g

70    If TechnicianMemberOf(UserCode) = "Managers" Then
80      frmAmend.txtUnitNumber = Unit
90      frmAmend.Show 1
100   Else
110     s = "Enter amendments to Typed product." & vbCrLf & _
            "Unit No." & Unit & " "
120     g.Col = 2  'group
130     s = s & " - " & g & vbCrLf
140     g.Col = 4 'expiry
150     s = s & "Expiry:" & g & vbCrLf & _
                "Note:- Only additions to existing" & vbCrLf & _
                "Antigens are allowed." & vbCrLf & _
                "Changes are irreversible!" & vbCrLf & _
                "Enter extra Type:"
  
160     g.Col = 0
170     Do
180       pc = ProductBarCodeFor(g)
190       g.Row = g.Row - 1
200     Loop Until pc <> "???"
  
210     Result = iBOX(s, "Modify Type")
220     If TimedOut Then Unload Me: Exit Sub
230     If Trim$(Result) = "" Then Exit Sub
240     Temp = AntigenDescription(Result)
250     If Temp <> "" Then Result = Temp

260     Sql = "UPDATE Product " & _
              "SET Screen = CONVERT(nvarchar(100), Screen) + ' ' + '" & Result & "' " & _
              "WHERE isbt128 = '" & Unit & "' " & _
              "AND BarCode = '" & pc & "'"
270     CnxnBB(0).Execute Sql

280     Sql = "UPDATE Latest " & _
              "SET Screen = CONVERT(nvarchar(100), Screen) + ' ' + '" & Result & "' " & _
              "WHERE isbt128 = '" & Unit & "' " & _
              "AND BarCode = '" & pc & "'"
290     CnxnBB(0).Execute Sql

300   End If

310   cmdSearch.Value = True

320   Exit Sub

bModify_Click_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmStockList", "bModify_Click", intEL, strES, Sql

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
80    Printer.Print "Stock List. "
90    Printer.Print lstproduct

      '****Report body


100   For i = 1 To 108
110       Printer.Print "_";
120   Next i
130   Printer.Print
140   Printer.Print FormatString("Unit No.", 20, "|");
150   Printer.Print FormatString("Group", 10, "|");
160   Printer.Print FormatString("Expiry", 20, "|");
170   Printer.Print FormatString("Screen", 64, "|")

180   Printer.Font.Bold = False
190   For i = 1 To 108
200       Printer.Print "-";
210   Next i
220   Printer.Print
230   For Y = 1 To g.Rows - 1
240       Printer.Print FormatString(g.TextMatrix(Y, 0), 20, "|");
250       Printer.Print FormatString(g.TextMatrix(Y, 2), 10, "|");
260       Printer.Print FormatString(g.TextMatrix(Y, 4), 20, "|");
270       Printer.Print FormatString(g.TextMatrix(Y, 5), 64, "|")

 
280   Next


290   Printer.EndDoc



300   For Each Px In Printers
310     If Px.DeviceName = OriginalPrinter Then
320       Set Printer = Px
330       Exit For
340     End If
350   Next

End Sub

Private Sub cmdSearch_Click()

10    FillG

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim Sql As String

10    On Error GoTo Form_Load_Error

20    lstproduct.Clear
30    lstproduct.AddItem "All"

40    Sql = "Select * from ProductList order by ListOrder"
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, Sql

70    Do While Not tb.EOF
80      lstproduct.AddItem tb!Wording
90      tb.MoveNext
100   Loop
110   lstproduct.ListIndex = 1

120   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmStockList", "Form_Load", intEL, strES, Sql

End Sub


Private Sub g_DblClick()
'      Dim Sql As String
'
'10    On Error GoTo g_DblClick_Error
'
'20        If g.TextMatrix(g.Row, 0) <> "" And g.TextMatrix(g.Row, 4) <> "" Then
'
'30            Sql = "Update Latest Set DateExpiry = DateAdd(Hour, 23, DateExpiry) WHERE (ISBT128 = '" & g.TextMatrix(g.Row, 0) & "') and DateExpiry = '" & Format(g.TextMatrix(g.Row, 4), "dd/mmm/yyyy") & "'"
'40            CnxnBB(0).Execute Sql
'
'              iMsg Sql
'50            Sql = "U  pdate Latest Set DateExpiry = DateAdd(Minute, 59, DateExpiry) WHERE (ISBT128 = '" & g.TextMatrix(g.Row, 0) & "') and DateExpiry = '" & Format(g.TextMatrix(g.Row, 4), "dd/mmm/yyyy") & " 23:00" & "'"
'60            CnxnBB(0).Execute Sql
'
'70            Sql = "Update Product Set DateExpiry = DateAdd(Hour, 23, DateExpiry) WHERE (ISBT128 = '" & g.TextMatrix(g.Row, 0) & "') and DateExpiry = '" & Format(g.TextMatrix(g.Row, 4), "dd/mmm/yyyy") & "'"
'80            CnxnBB(0).Execute Sql
'90            Sql = "Update Product Set DateExpiry = DateAdd(Minute, 59, DateExpiry) WHERE (ISBT128 = '" & g.TextMatrix(g.Row, 0) & "') and DateExpiry = '" & Format(g.TextMatrix(g.Row, 4), "dd/mmm/yyyy") & " 23:00" & "'"
'100           CnxnBB(0).Execute Sql
'110       End If
'
'120   Exit Sub
'
'g_DblClick_Error:
'
'       Dim strES As String
'       Dim intEL As Integer
'
'130    intEL = Erl
'140    strES = Err.Description
'150    LogError "fstocklist", "g_DblClick", intEL, strES, Sql
          
End Sub

Private Sub g_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim BarCode As String
      Dim PackNumber As String
      Dim intGY As Integer
      Dim intGX As Integer
      Dim Expiry As String
      Dim xSave As Integer
      Dim ySave As Integer
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo g_MouseDown_Error

20    If g.MouseRow = 0 Then Exit Sub

30    xSave = g.Col
40    ySave = g.Row

50    g.Col = 0
60    For intGY = 1 To g.Rows - 1
70      g.Row = intGY
80      If g.CellBackColor = vbYellow Then
90        For intGX = 0 To g.Cols - 1
100         g.Col = intGX
110         g.CellBackColor = 0
120       Next
130       Exit For
140     End If
150   Next
160   g.Row = ySave
170   For intGX = 0 To g.Cols - 1
180     g.Col = intGX
190     g.CellBackColor = vbYellow
200   Next
210   g.Col = xSave
220   g.Row = ySave

230   If g.Col = 1 Then
240     frmUnitNotes.txtUnitNumber = g.TextMatrix(g.Row, 0)
250     frmUnitNotes.txtExpiry = g.TextMatrix(g.Row, 4)
260     frmUnitNotes.Show 1
270     FillG
280   ElseIf g.Col = 3 Then
290     If Trim$(g.TextMatrix(g.Row, 2)) = "" Then
300       Exit Sub
310     End If
320     g.TextMatrix(g.Row, 3) = IIf(g.TextMatrix(g.Row, 3) = "", "Yes", "")
330     PackNumber = g.TextMatrix(g.Row, 0)
340     For intGY = g.Row To 1 Step -1
350       BarCode = ProductBarCodeFor(g.TextMatrix(intGY, 0))
360       If BarCode <> "???" Then
370         Exit For
380       End If
390     Next
400     Expiry = Format(g.TextMatrix(g.Row, 4), "dd/mmm/yyyy")

410     Ps.LoadLatestISBT128 PackNumber, BarCode
420     If Ps.Count = 1 Then
430       Set p = Ps(1)
440       p.Checked = g.TextMatrix(g.Row, 3) = "Yes"
450       p.UserName = UserCode
460       p.RecordDateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
470       p.Save
480     End If
490   End If

500   Exit Sub

g_MouseDown_Error:

      Dim strES As String
      Dim intEL As Integer

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmStockList", "g_MouseDown", intEL, strES

End Sub




