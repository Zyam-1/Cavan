VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUnitFating 
   Caption         =   "NetAcquire - Unit Fating"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15345
   Icon            =   "frmUnitFating.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   15345
   Begin VB.CommandButton cmdUnitHistory 
      Caption         =   "Unit History"
      Enabled         =   0   'False
      Height          =   765
      Left            =   9660
      Picture         =   "frmUnitFating.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   270
      Width           =   975
   End
   Begin VB.Frame fraFate 
      Caption         =   "Fate of Unit"
      Enabled         =   0   'False
      Height          =   2475
      Left            =   13020
      TabIndex        =   14
      Top             =   1230
      Width           =   2085
      Begin VB.CommandButton cmdFate 
         Caption         =   "Fate Unit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1710
         Width           =   1215
      End
      Begin VB.OptionButton optFate 
         Caption         =   "Returned to Supplier"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   960
         Width           =   1785
      End
      Begin VB.OptionButton optFate 
         Caption         =   "Transfused"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   18
         Top             =   1230
         Width           =   1185
      End
      Begin VB.OptionButton optFate 
         Caption         =   "Expired"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   690
         Width           =   855
      End
      Begin VB.OptionButton optFate 
         Caption         =   "Pack Dispatched"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   16
         Top             =   2190
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optFate 
         Caption         =   "Destroyed"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show"
      Height          =   1065
      Left            =   3300
      TabIndex        =   9
      Top             =   120
      Width           =   2115
      Begin VB.CheckBox chkShow 
         Caption         =   "Pending Transfusion"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   810
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Cross-matched"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Pending X-Match"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   11
         Top             =   390
         Width           =   1635
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Issued"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   210
         Width           =   825
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdFate 
      Height          =   6465
      Left            =   150
      TabIndex        =   8
      Top             =   1320
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   11404
      _Version        =   393216
      Cols            =   10
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmUnitFating.frx":0D0C
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1065
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   3045
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   675
         Left            =   1830
         Picture         =   "frmUnitFating.frx":0E3E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   147324929
         CurrentDate     =   38373
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   147324929
         CurrentDate     =   38373
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   13560
      Picture         =   "frmUnitFating.frx":1148
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "bCancel"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   5610
      Picture         =   "frmUnitFating.frx":17B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   975
   End
   Begin VB.CommandButton bprint 
      Caption         =   "Print"
      Height          =   765
      Left            =   7740
      Picture         =   "frmUnitFating.frx":1ABC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   22
      Top             =   7890
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
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
      Left            =   6600
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmUnitFating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private CurrentFateCode As String
Private CurrentPrompt As String
Private DatePickerCaption As String
Private CurrentOptions As New Collection
Private Sub DisableFraFate()

      Dim n As Integer

10    For n = 0 To 4
20      optFate(n).Value = 0
30    Next
40    fraFate.Enabled = False
50    cmdFate.Enabled = False
60    cmdUnitHistory.Enabled = False

End Sub

Private Sub EnableFraFate()

10    fraFate.Enabled = True

20    cmdUnitHistory.Enabled = True

End Sub


Private Sub FillGrdFate()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim Criteria As String

10    On Error GoTo FillGrdFate_Error

20    DisableFraFate

30    With grdFate
40      .Visible = False
50      .Rows = 2
60      .AddItem ""
70      .RemoveItem 1
80    End With

90    Criteria = WhereClause()
100   If Criteria = "" Then
110     grdFate.Visible = True
120     Exit Sub
130   End If

140   sql = "Select * from Latest where " & _
            Criteria & _
            "and DateTime between '" & Format$(dtFrom, "dd/MMM/yyyy") & _
            "' and '" & Format$(dtTo, "dd/MMM/yyyy") & " 23:59' " & _
            "order by DateTime desc"
150   Set tb = New Recordset
160   RecOpenServerBB 0, tb, sql

170   Do While Not tb.EOF
180     s = tb!ISBT128 & "" & vbTab & _
            ProductWordingFor(tb!BarCode & "") & vbTab & _
            gEVENTCODES(tb!Event).Text & vbTab & _
            Format$(tb!DateTime, "dd/MM/yy HH:nn:ss") & vbTab & _
            Format$(tb!DateExpiry, "dd/MM/yy HH:nn") & vbTab & _
            Bar2Group(tb!GroupRh & "") & vbTab & _
            tb!LabNumber & vbTab & _
            tb!PatName & vbTab & _
            tb!Patid & vbTab & _
            tb!Notes & ""
190     grdFate.AddItem s
200     tb.MoveNext
210   Loop

220   With grdFate
230     If .Rows > 2 Then
240       .RemoveItem 1
250     End If
260     .Visible = True
270   End With

280   Exit Sub

FillGrdFate_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmUnitFating", "FillGrdFate", intEL, strES, sql


End Sub

Private Function WhereClause() As String

      Dim s As String

10    s = ""
20    If chkShow(0).Value = 1 Then
30      s = "Event = 'I'"
40    End If
50    If chkShow(1).Value = 1 Then
60      If s <> "" Then s = s & " or Event = 'P'"
70    End If
80    If chkShow(2).Value = 1 Then
90      If s <> "" Then s = s & " or Event = 'X'"
100   End If
110   If chkShow(3).Value = 1 Then
120     If s <> "" Then s = s & " or Event = 'Y'"
130   End If

140   If s <> "" Then
150     s = " (" & s & ") "
160   End If

170   WhereClause = s

End Function

Private Sub bprint_Click()
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub
30    Printer.FontName = "Courier New"
40    Printer.Font.Size = 9
50    Printer.Orientation = vbPRORPortrait
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                         Search Results For Unit Fating"
90    Printer.Print "                                         From "; dtFrom.Value; " To "; dtTo.Value

      'Heading Section

100   Printer.Font.Bold = True
110   For i = 1 To 108
120       Printer.Print "_";
130   Next i
140   Printer.Print
150   Printer.Print FormatString("Unit #", 15, "|"); 'Unit Number
160   Printer.Print FormatString("Product", 19, "|"); 'Product
170   Printer.Print FormatString("Latest.", 8, "|"); 'Latest
180   Printer.Print FormatString("Date/Time", 14, "|"); 'Date/Time
190   Printer.Print FormatString("Expiry", 8, "|"); 'Expiry Date
200   Printer.Print FormatString("Group", 5, "|"); 'Group
210   Printer.Print FormatString("SampID", 8, "|"); 'SampleID
220   Printer.Print FormatString("Name", 15, "|"); 'Name
230   Printer.Print FormatString("Chart #", 8, "|") 'Chart Number

240   Printer.Font.Bold = False
250   For i = 1 To 108
260       Printer.Print "-";
270   Next i
  
280   Printer.Print
290   For Y = 1 To grdFate.Rows - 1
    
              'Detail section

300           Printer.Print FormatString(grdFate.TextMatrix(Y, 0), 15, "|"); 'Unit Number
310           Printer.Print FormatString(grdFate.TextMatrix(Y, 1), 19, "|"); 'Product
320           Printer.Print FormatString(Left$(grdFate.TextMatrix(Y, 2), 8), 8, "|"); 'Latest
330           Printer.Print FormatString(grdFate.TextMatrix(Y, 3), 14, "|"); 'Date/Time
340           Printer.Print FormatString(grdFate.TextMatrix(Y, 4), 8, "|"); 'Expiry Date
350           Printer.Print FormatString(grdFate.TextMatrix(Y, 5), 5, "|"); 'Group
360           Printer.Print FormatString(grdFate.TextMatrix(Y, 6), 8, "|"); 'SampleID
370           Printer.Print FormatString(grdFate.TextMatrix(Y, 7), 15, "|"); 'Name
380           Printer.Print FormatString(grdFate.TextMatrix(Y, 8), 8, "|") 'Chart Number
  
390   Next

400   Printer.EndDoc

410   For Each Px In Printers
420     If Px.DeviceName = OriginalPrinter Then
430       Set Printer = Px
440       Exit For
450     End If
460   Next
End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdFate_Click()

      Dim BarCode As String
      Dim Expiry As String
      Dim UnitNumber As String
      Dim f As Form
      Dim Accepted As Boolean
      Dim Reason As String
      Dim Comment As String
      Dim EventDateTime As String
      Dim s As Variant
      Dim Ps As New Products
      Dim p As Product

10    On Error GoTo cmdFate_Click_Error

20    BarCode = ProductBarCodeFor(grdFate.TextMatrix(grdFate.row, 1))
30    Expiry = Format$(grdFate.TextMatrix(grdFate.row, 4), "dd/MMM/yyyy")
40    UnitNumber = grdFate.TextMatrix(grdFate.row, 0)

50    Ps.LoadLatestISBT128 UnitNumber, BarCode
60    If Ps.Count > 0 Then
70      Set p = Ps(1)

80      Set f = New frmQueryValidate
90      With f
100       .ShowDatePicker = True
110       .DatePickerCaption = DatePickerCaption
120       .Prompt = CurrentPrompt
130       .Options = New Collection
140       For Each s In CurrentOptions
150         .Options.Add s
160       Next
170       .CurrentStatus = gEVENTCODES(p.PackEvent).Text
180       If p.PackEvent = "Y" Then
190         .DateTimeRemovedFromLab = p.RecordDateTime
200       Else
210         .DateTimeRemovedFromLab = ""
220       End If
230       .NewStatus = gEVENTCODES(CurrentFateCode).Text
240       .UnitNumber = UnitNumber
250       .Chart = grdFate.TextMatrix(grdFate.row, 8)
260       .PatientName = grdFate.TextMatrix(grdFate.row, 7)
270       .Show 1
280       Accepted = .retval
290       Reason = .Reason
300       Comment = .Comment
310       EventDateTime = .DateTimeReturn
320     End With
330     Set f = Nothing

340     If Accepted Then
  
350       p.PackEvent = CurrentFateCode
360       p.UserName = UserCode
370       p.RecordDateTime = Format$(EventDateTime, "dd/MMM/yyyy HH:nn:ss")
380       p.Notes = Comment
390       p.Save
    
400     End If
  
410   Else
420     iMsg "Unit Number not found", vbCritical
430     If TimedOut Then Unload Me: Exit Sub
440     Exit Sub
450   End If

460   FillGrdFate

470   Exit Sub

cmdFate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmUnitFating", "cmdFate_Click", intEL, strES

End Sub

Private Sub cmdSearch_Click()

10    FillGrdFate

End Sub

Private Sub cmdUnitHistory_Click()

10    With frmUnitHistory
20      .UnitNumber = grdFate.TextMatrix(grdFate.row, 0)
30      .lblExpiry = grdFate.TextMatrix(grdFate.row, 4)
40      .ProductName = grdFate.TextMatrix(grdFate.row, 1)
50      .cmdSearch = True
60      .Show 1
70    End With

End Sub

Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Unit Fating" & vbCr
20    strHeading = strHeading & "From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid grdFate, Me, strHeading

End Sub

Private Sub Form_Load()

10    dtFrom.Value = Format$(Now - 7, "dd/MM/yyyy")
20    dtTo.Value = Format$(Now, "dd/MM/yyyy")

End Sub


Private Sub grdFate_Click()

      Dim X As Integer
      Dim Y As Integer
      Dim ySave As Integer

10    DisableFraFate

20    If grdFate.MouseRow = 0 Then
30      If InStr(grdFate.TextMatrix(0, grdFate.col), "Date") <> 0 Then
40        grdFate.Sort = 9
50      Else
60        If SortOrder Then
70          grdFate.Sort = flexSortGenericAscending
80        Else
90          grdFate.Sort = flexSortGenericDescending
100       End If
110     End If
120     SortOrder = Not SortOrder
130   ElseIf grdFate.TextMatrix(1, 0) <> "" Then
  
140     EnableFraFate
150     ySave = grdFate.row
160     grdFate.col = 0
170     For Y = 1 To grdFate.Rows - 1
180       grdFate.row = Y
190       If grdFate.CellBackColor = vbRed Then
200         For X = 0 To grdFate.Cols - 1
210           grdFate.col = X
220           grdFate.CellBackColor = 0
230         Next
240         Exit For
250       End If
260     Next
270     grdFate.row = ySave
280     For X = 0 To grdFate.Cols - 1
290       grdFate.col = X
300       grdFate.CellBackColor = vbRed
310     Next
'320     txtComment = grdFate.TextMatrix(grdFate.Row, 9)
  
320   End If

End Sub

Private Sub grdFate_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(grdFate.TextMatrix(Row1, grdFate.col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(grdFate.TextMatrix(Row2, grdFate.col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(grdFate.TextMatrix(Row1, grdFate.col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(grdFate.TextMatrix(Row2, grdFate.col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("s", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("s", d2, d1))
150   End If

End Sub


Private Sub optFate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    Set CurrentOptions = New Collection

20    DatePickerCaption = "Event Date/Time"

30    Select Case Index
  
  Case 0
40        CurrentFateCode = "D"
50        CurrentPrompt = "Enter Reason for Destroying."
60        CurrentOptions.Add "Product Expired"
70        CurrentOptions.Add "Out of Fridge > 30 min"

80      Case 1
90        CurrentFateCode = "F"
100       CurrentPrompt = "Enter Details of Despatch."
110       CurrentOptions.Add "Unit Transferred to Drogheda"
120       CurrentOptions.Add "Inter-Hospital Transfer"
130       CurrentOptions.Add "Laboratory Use"
  
140     Case 2
150       CurrentFateCode = "J"
160       CurrentPrompt = "Enter Details of Expiry."
170       CurrentOptions.Add "Product Expired"
180       CurrentOptions.Add "Out of Fridge > 30 min"
  
190     Case 3
200       CurrentFateCode = "T"
210       CurrentPrompt = "Enter reason for return."
220       CurrentOptions.Add "Returned for credit"
230       CurrentOptions.Add "Returned not for credit"
  
240     Case 4
250       CurrentFateCode = "S"
260       CurrentPrompt = ""
270       DatePickerCaption = "Transfusion END Date/Time"

280   End Select

290   cmdFate.Enabled = True

End Sub


