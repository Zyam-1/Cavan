VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUsage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C/T Ratio"
   ClientHeight    =   8835
   ClientLeft      =   180
   ClientTop       =   360
   ClientWidth     =   9465
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
   Icon            =   "7frmUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8835
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   8220
      Picture         =   "7frmUsage.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7350
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   180
      TabIndex        =   19
      Top             =   180
      Width           =   5445
      Begin VB.OptionButton obetween 
         Caption         =   "&Year to Date "
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
         Index           =   7
         Left            =   3600
         TabIndex        =   26
         Top             =   1005
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last Full Y&ear"
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
         Index           =   6
         Left            =   3600
         TabIndex        =   33
         Top             =   750
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last F&ull Quarter"
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
         Index           =   5
         Left            =   3600
         TabIndex        =   32
         Top             =   495
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last &Quarter"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last &Full Month"
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
         Index           =   3
         Left            =   1740
         TabIndex        =   30
         Top             =   1005
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last &Month"
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
         Left            =   1740
         TabIndex        =   29
         Top             =   750
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "Last &Week"
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
         Left            =   1740
         TabIndex        =   28
         Top             =   495
         Width           =   1815
      End
      Begin VB.OptionButton obetween 
         Caption         =   "&Today"
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
         Left            =   1740
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   36963
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   840
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   92536833
         CurrentDate     =   36963
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2235
      Left            =   5730
      TabIndex        =   5
      Top             =   2160
      Width           =   3555
      Begin VB.OptionButton oSpecific 
         Caption         =   "Specific"
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   1590
         Width           =   1125
      End
      Begin VB.OptionButton oGeneric 
         Alignment       =   1  'Right Justify
         Caption         =   "Generic"
         Height          =   315
         Left            =   690
         TabIndex        =   8
         Top             =   660
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.ComboBox cproduct 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Frame FrameProduct 
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   1770
         TabIndex        =   6
         Top             =   0
         Width           =   1785
         Begin VB.OptionButton o 
            Caption         =   "Whole Blood"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   510
            Width           =   1215
         End
         Begin VB.OptionButton o 
            Caption         =   "Platelets"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   990
            Width           =   1305
         End
         Begin VB.OptionButton o 
            Caption         =   "Plasma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   750
            Width           =   945
         End
         Begin VB.OptionButton o 
            Caption         =   "Cryoprecipitate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   1230
            Width           =   1425
         End
         Begin VB.OptionButton o 
            Caption         =   "Red Cells"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   6780
      TabIndex        =   4
      Top             =   360
      Width           =   1395
      Begin VB.OptionButton oSearchBy 
         Caption         =   "Ward"
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
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1050
         Width           =   795
      End
      Begin VB.OptionButton oSearchBy 
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
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton oSearchBy 
         Caption         =   "Conditions"
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
         TabIndex        =   16
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton oSearchBy 
         Caption         =   "Procedure"
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
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6915
      Left            =   150
      TabIndex        =   3
      Top             =   1590
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   12197
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      ForeColorFixed  =   16711680
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   2
      FormatString    =   "<Source                         |^XMatched|^Transfused|^Ratio      "
   End
   Begin VB.CommandButton bprint 
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
      Height          =   915
      Left            =   7020
      Picture         =   "7frmUsage.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton bstart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5820
      Picture         =   "7frmUsage.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   915
      Left            =   8220
      Picture         =   "7frmUsage.frx":1548
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   23
      Top             =   8580
      Width           =   5625
      _ExtentX        =   9922
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
      Height          =   345
      Left            =   8040
      TabIndex        =   24
      Top             =   8130
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblWait 
      AutoSize        =   -1  'True
      Caption         =   "Calculating - Please wait...."
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
      Left            =   5850
      TabIndex        =   25
      Top             =   7140
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function ProductList() As String

      Dim s As String
      Dim tb As Recordset
      Dim sql As String
      Dim GenericName As String
      Dim n As Integer

10    On Error GoTo ProductList_Error

20    s = "BarCode = '"

30    If oSpecific Then
40      s = s & ProductBarCodeFor(cproduct) & "'"
50    Else
60      For n = 0 To 4
70        If o(n) Then
80          GenericName = o(n).Caption
90          Exit For
100       End If
110     Next
120     sql = "Select BarCode from ProductList where " & _
              "Generic = '" & GenericName & "'"
130     Set tb = New Recordset
140     RecOpenServerBB 0, tb, sql
150     Do While Not tb.EOF
160       s = s & tb!BarCode & "' or BarCode = '"
170       tb.MoveNext
180     Loop
190     s = Left$(s, Len(s) - 15)
200   End If

210   ProductList = s

220   Exit Function

ProductList_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmUsage", "ProductList", intEL, strES, sql

  
End Function

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

      Dim X As Integer
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "Crossmatched / Transfused Ratio"
40    Printer.Print cproduct
50    Printer.Print "Between " & dtFrom & " and " & dtTo
60    Printer.Print

70    For Y = 0 To g.Rows - 1
80      g.Row = Y
90      For X = 0 To g.Cols - 1
100       g.Col = X
110       Printer.Print Tab(Choose(X + 1, 1, 25, 35, 45));
120       Printer.Print g;
130     Next
140     Printer.Print
150   Next
160   Printer.EndDoc

170   For Each Px In Printers
180     If Px.DeviceName = OriginalPrinter Then
190       Set Printer = Px
200       Exit For
210     End If
220   Next

End Sub

Private Sub bstart_Click()

      Dim n As Integer
      Dim sql As String
      Dim tc As Recordset
      Dim Ratio As Single
      Dim ProdList As String
      Dim strFromTime As String
      Dim strToTime As String

10    On Error GoTo bstart_Click_Error

20    lblWait.Visible = True
30    lblWait.Refresh

40    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
50    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

60    For n = 1 To g.Rows - 1
70      g.TextMatrix(n, 1) = ""
80      g.TextMatrix(n, 2) = ""
90      g.TextMatrix(n, 3) = ""
100   Next

110   ProdList = ProductList()

120   For n = 1 To g.Rows - 1
130     sql = "Select count (LabNumber) as tot from Product where " & _
        "Event = 'X' " & _
        "AND (" & ProdList & ") " & _
        "AND LabNumber IN " & _
        "( Select LabNumber from PatientDetails where " & _
           Source() & " = '" & g.TextMatrix(n, 0) & "'" & _
        "  and DateTime between '" & strFromTime & "' " & _
        "  and '" & strToTime & "' )"
140     Set tc = New Recordset
150     RecOpenServerBB 0, tc, sql
160     If tc!Tot <> 0 Then
170       g.TextMatrix(n, 1) = Val(g.TextMatrix(n, 1)) + tc!Tot
180       g.Refresh

190       sql = "Select count (LabNumber) as tot from Product where " & _
          "Event = 'S' " & _
          "AND (" & ProdList & ") " & _
          "AND LabNumber IN " & _
          "( Select LabNumber from PatientDetails where " & _
             Source() & " = '" & g.TextMatrix(n, 0) & "'" & _
          "  and DateTime between '" & strFromTime & "' " & _
          "  and '" & strToTime & "' )"
200       Set tc = New Recordset
210       RecOpenClientBB 0, tc, sql
220       If tc!Tot <> 0 Then
230   g.TextMatrix(n, 2) = Val(g.TextMatrix(n, 2)) + tc!Tot
240   g.Refresh
250       End If
260     End If
270   Next
  
280   For n = 1 To g.Rows - 1
290     If g.TextMatrix(n, 1) <> "" Then
300       Ratio = Val(g.TextMatrix(n, 2)) / Val(g.TextMatrix(n, 1))
310       g.TextMatrix(n, 3) = Format$(Ratio * 100, "0.0") & "%"
320       g.Refresh
330     End If
340   Next

350   lblWait.Visible = False
360   lblWait.Refresh

370   Exit Sub

bstart_Click_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmUsage", "bstart_Click", intEL, strES, sql

End Sub

Private Sub FillcProduct()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillcProduct_Error

20    sql = "Select * from ProductList order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServerBB 0, tb, sql
50    Do While Not tb.EOF
60      cproduct.AddItem tb!Wording
70      tb.MoveNext
80    Loop

90    cproduct.ListIndex = 0

100   Exit Sub

FillcProduct_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmUsage", "FillcProduct", intEL, strES, sql


End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub dtFrom_CloseUp()

      Dim strFromTime As String
      Dim strToTime As String

10    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
20    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

30    loadlist strFromTime, strToTime

End Sub


Private Sub dtTo_CloseUp()

      Dim strFromTime As String
      Dim strToTime As String

10    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
20    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

30    loadlist strFromTime, strToTime

End Sub


Private Sub Form_Activate()

      Dim strFromTime As String
      Dim strToTime As String

10    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
20    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

30    loadlist strFromTime, strToTime

End Sub

Private Sub Form_Load()

10    dtTo = Format(Now, "dd/mmm/yyyy")
20    dtFrom = Format(Now - 7, "dd/mmm/yyyy")
30    FillcProduct

End Sub

Private Sub loadlist(ByVal strFromTime As String, ByVal strToTime As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
      Dim Found As Boolean
      Dim BlankFound As Boolean

10    On Error GoTo loadlist_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    BlankFound = False

60    sql = "select distinct " & Source() & " as Tp " & _
            "from patientdetails where " & _
            "datetime between '" & strFromTime & "' " & _
            " and '" & strToTime & "'"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql

90    Do While Not tb.EOF
100     If Trim$(tb!tP & "") = "" Then
110       BlankFound = True
120     End If
130     Found = False
140     For n = 1 To g.Rows - 1
150       If g.TextMatrix(n, 0) = Trim$(tb!tP & "") Then
160         Found = True
170         Exit For
180       End If
190     Next
200     If Not Found Then
210       g.AddItem UCase$(Trim$(tb!tP & ""))
220     End If
230     tb.MoveNext
240   Loop

250   If Not BlankFound Then
260     If g.Rows > 2 Then
270       g.RemoveItem 1
280     End If
290   End If

300   Exit Sub

loadlist_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmUsage", "loadlist", intEL, strES, sql


End Sub



Private Sub obetween_Click(Index As Integer)
           Dim upto As String
      Dim FromTime As String
      Dim ToTime As String

10    dtFrom = Format(BetweenDates(Index, upto), "dd/mmm/yyyy")
20    dtTo = Format(upto, "dd/mmm/yyyy")

30    FromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
40    ToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

50    loadlist FromTime, ToTime
End Sub

Private Sub oGeneric_Click()

      Dim n As Integer

10    FrameProduct.Enabled = True
20    cproduct.Enabled = False
30    cproduct.ListIndex = -1

40    For n = 0 To 4
50      o(n).Enabled = True
60    Next
70    o(1) = True

End Sub

Private Sub osearchby_Click(Index As Integer)

      Dim strFromTime As String
      Dim strToTime As String

10    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
20    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

30    loadlist strFromTime, strToTime

40    g.TextMatrix(0, 0) = oSearchBy(Index).Caption

End Sub

Private Function Source() As String

      Dim n As Integer

10    For n = 0 To 3
20      If oSearchBy(n).Value = True Then
30        Source = "[" & oSearchBy(n).Caption & "]"
40        Exit For
50      End If
60    Next

End Function

Private Sub oSpecific_Click()

      Dim n As Integer

10    FrameProduct.Enabled = False
20    cproduct.Enabled = True
30    cproduct.ListIndex = 0

40    For n = 0 To 4
50      o(n).Enabled = False
60      o(n) = False
70    Next

End Sub


