VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fprodmove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Transfusion"
   ClientHeight    =   8910
   ClientLeft      =   465
   ClientTop       =   495
   ClientWidth     =   12030
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
   Icon            =   "fprodmove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8910
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
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
      Height          =   900
      Left            =   7890
      Picture         =   "fprodmove.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
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
      Height          =   900
      Left            =   4890
      Picture         =   "fprodmove.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   900
   End
   Begin VB.Frame Frame2 
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
      Height          =   945
      Left            =   3270
      TabIndex        =   11
      Top             =   60
      Width           =   1515
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147128321
         CurrentDate     =   38738
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147128321
         CurrentDate     =   38738
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show products capable of being"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   2805
      Begin VB.OptionButton o 
         Caption         =   "ReStocked"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1320
         TabIndex        =   10
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton o 
         Caption         =   "Destroyed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Returned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   8
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton o 
         Alignment       =   1  'Right Justify
         Caption         =   "Transfused"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7125
      Left            =   120
      TabIndex        =   5
      Top             =   1470
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12568
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      ScrollTrack     =   -1  'True
      HighLight       =   2
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"fprodmove.frx":105F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   900
      Left            =   6390
      Picture         =   "fprodmove.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton bconfirm 
      Appearance      =   0  'Flat
      Caption         =   "Transfuse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   9390
      Picture         =   "fprodmove.frx":17BC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton bviewall 
      Appearance      =   0  'Flat
      Caption         =   "View &All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4890
      TabIndex        =   2
      Top             =   1080
      Width           =   900
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
      Height          =   900
      Left            =   10890
      Picture         =   "fprodmove.frx":1AC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   60
      TabIndex        =   17
      Top             =   8670
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   397
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
      Left            =   7733
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   570
      Picture         =   "fprodmove.frx":2130
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label lhistory 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Unit Number to show full history."
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   1110
      Visible         =   0   'False
      Width           =   3795
   End
End
Attribute VB_Name = "fprodmove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim detailsselected As Integer
Dim showinghistory As Integer

Private SortOrder As Boolean

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bConfirm_Click()

      Dim NewEventCode As String
      Dim NumberToFind As String
      Dim pid As String
      Dim pname As String
      Dim timecode As String
      Dim Product As String
      Dim GridUser As String
      Dim DateTime As String
      Dim GroupRh As String
      Dim DateExpiry As String
      Dim LabNumber As String
Dim Ps As New Products
Dim p As Product

10    On Error GoTo bConfirm_Click_Error

20    timecode = Now

30    Select Case Left$(bconfirm.Caption, 3)
        Case "Ret": NewEventCode = "T"
40      Case "Res": NewEventCode = "R"
50      Case "Des": NewEventCode = "D"
60      Case "Tra": NewEventCode = "S"
70      Case "Iss": NewEventCode = "I"
80    End Select

90    If Not detailsselected And NewEventCode = "S" Then
100     Beep
110     iMsg "Select Patient."
120     If TimedOut Then Unload Me: Exit Sub
130     Exit Sub
140   End If

150   g.col = 0
160   NumberToFind = g
170   g.col = 1
180   Product = ProductBarCodeFor(g)
190   g.col = 2
200   GroupRh = Group2Bar(g)
210   g.col = 3
220   DateExpiry = Format(g, "dd/mmm/yyyy hh:mm")
230   g.col = 9
240   LabNumber = g

250   If NewEventCode = "S" Then
260     g.col = 5
270     pid = g
280     g.col = 6
290     pname = g
300   End If

310   g.col = 7
320   GridUser = g
330   DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")

340   Ps.LoadLatestISBT128 NumberToFind, Product
350   If Ps.Count > 0 Then
360     Set p = Ps.Item(1)
370     p.ISBT128 = NumberToFind
380     p.RecordDateTime = DateTime
390     p.PackEvent = NewEventCode

391     If UCase(p.PackEvent) = "R" Or UCase(p.PackEvent) = "D" Or UCase(p.PackEvent) = "T" Then 'R - Restocked, D -Destroyed, T - Return to Supplier
392         p.cco = False
393         p.ccor = False
394         p.cen = False
395         p.cenr = False
396         p.crt = False
397         p.crtr = False
398     End If

400     p.Chart = pid
410     p.PatName = pname
420     p.UserName = UserCode
430     p.SampleID = LabNumber
440     p.Save
450   End If

460   bviewall.Value = True

470   Exit Sub

bConfirm_Click_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "fprodmove", "bConfirm_Click", intEL, strES

End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim Title As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

      '****Report heading
30    Printer.Orientation = vbPRORPortrait
40    Printer.Font.Name = "Courier New"
50    Printer.Font.Size = 9
60    Printer.Font.Bold = True
70    If o(0).Value = True Then
80        Title = o(0).Caption
90    ElseIf o(1).Value = True Then
100       Title = o(1).Caption
110   ElseIf o(2).Value = True Then
120       Title = o(2).Caption
130   ElseIf o(3).Value = True Then
140       Title = o(3).Caption
150   End If
160   Printer.Print "                                        "; Title; " Units"
170   Printer.Print
180   Printer.Print "Search Results for "; Title; " units From "; dtFrom.Value; " To "; dtTo.Value

      '****Report body

190   Printer.Print
200   Printer.Print FormatString("Unit", 16, "|");
210   Printer.Print FormatString("Product", 64, "|");
220   Printer.Print FormatString("Group", 10, "|");
230   Printer.Print FormatString("Exp.Date", 14, "|")

240   Printer.Print "------------------------------------------------------------------------------------------------------------"
250   Printer.Font.Bold = False
260   For Y = 1 To g.Rows - 1
270       Printer.Print FormatString(g.TextMatrix(Y, 0), 16, "|");
280       Printer.Print FormatString(g.TextMatrix(Y, 1), 64, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 2), 10, "|");
300       Printer.Print FormatString(g.TextMatrix(Y, 3), 14, "|")
    
    
  
310   Next

320   Printer.EndDoc

330   For Each Px In Printers
340     If Px.DeviceName = OriginalPrinter Then
350       Set Printer = Px
360       Exit For
370     End If
380   Next

End Sub

Private Sub bviewall_Click()

10    FillG
20    bviewall.Enabled = False
30    lhistory.Visible = True
40    bconfirm.Enabled = False
50    detailsselected = False
60    showinghistory = False

End Sub

Private Sub cmdSearch_Click()

10    FillG

20    If o(0) Then
30      bconfirm.Caption = "Transfuse"
40    ElseIf o(1) Then
50      bconfirm.Caption = "Return"
60    ElseIf o(2) Then
70      bconfirm.Caption = "Destroy"
80    Else
90      bconfirm.Caption = "Restock"
100   End If

End Sub

Private Sub cmdXL_Click()

      Dim strHeading As String

10    If o(0).Value = True Then
20        strHeading = o(0).Caption
30    ElseIf o(1).Value = True Then
40        strHeading = o(1).Caption
50    ElseIf o(2).Value = True Then
60        strHeading = o(2).Caption
70    ElseIf o(3).Value = True Then
80        strHeading = o(3).Caption
90    End If

100   strHeading = strHeading & " Units" & vbCr
110   strHeading = strHeading & "Search Results From " & dtFrom.Value & " To " & dtTo.Value & vbCr
120   strHeading = strHeading & " " & vbCr

130   ExportFlexGrid g, Me, strHeading

End Sub



Private Sub Form_Load()

10    dtFrom = Format$(Now - 7, "dd/MM/yyyy")
20    dtTo = Format$(Now, "dd/MM/yyyy")

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
30        FillG
      '**************************************
End Sub


Private Sub g_Click()

      Dim s As String
      Dim lot As String
      Dim Product As String
      Dim ws As Recordset
      Dim sql As String

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If InStr(g.TextMatrix(0, g.col), "Date") <> 0 Then
40        g.Sort = 9
50      Else
60        If SortOrder Then
70          g.Sort = flexSortGenericAscending
80        Else
90          g.Sort = flexSortGenericDescending
100       End If
110     End If
120     SortOrder = Not SortOrder
130     Exit Sub
140   End If

150   g.col = 0

160   If Trim$(g) = "" Then
170     Beep
180     Exit Sub
190   End If

200   If showinghistory Then
210     detailsselected = True
220     Exit Sub
230   End If

240   lot = g
250   g.col = 1
260   Product = g
270   Product = ProductBarCodeFor(Product)

280   g.Rows = 2
290   g.AddItem ""
300   g.RemoveItem 1

310   sql = "select * from product where " & _
            "ISBT128 = '" & lot & "' " & _
            "and barcode = '" & Product & "' " & _
            "order by Counter"
320   Set ws = New Recordset
330   RecOpenServerBB 0, ws, sql

340   ws.MoveLast
350   Do
360     s = ws!ISBT128 & vbTab
370     s = s & ProductWordingFor(ws!BarCode & "") & vbTab
380     s = s & Bar2Group(ws!GroupRh & "") & vbTab
390     s = s & Format$(ws!DateExpiry, "dd/MMM/yyyy HH:mm") & vbTab
400     s = s & gEVENTCODES(ws!Event & "").Text & vbTab
410     s = s & ws!Patid & vbTab
420     s = s & ws!PatName & vbTab
430     s = s & ws!Operator & vbTab
440     s = s & Format(ws!DateTime, "dd/mm/yy hh:mm:ss") & vbTab
450     s = s & ws!LabNumber & ""
460     g.AddItem s, 1
470     ws.MovePrevious
480   Loop While Not ws.BOF

490   bviewall.Enabled = True
500   lhistory.Visible = False
510   bconfirm.Enabled = True
520   showinghistory = True

530   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "fprodmove", "g_Click", intEL, strES, sql

End Sub

Private Sub FillG()


      Dim s As String
      Dim FromDate As String
      Dim ToDate As String
Dim Ps As New Products
Dim p As Product
Dim Events As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    FromDate = Format$(dtFrom, "dd/MMM/yyyy")
60    ToDate = Format$(dtTo, "dd/MMM/yyyy")

70    Ps.LoadLatestBetweenDates FromDate, ToDate

80    If o(0) Then  'capable of being transfused
90      Events = "XI"
100   ElseIf o(1) Then 'capable of being returned to supplier
110     Events = "CRW"
120   ElseIf o(2) Then 'capable of being destroyed
130     Events = "CRXPIW"
140   ElseIf o(3) Then 'capable of being restocked
150     Events = "XIW"
160   End If

170   For Each p In Ps
180     If InStr(Events, p.PackEvent) > 0 Then
190       s = p.ISBT128 & vbTab & _
              ProductWordingFor(p.BarCode & "") & vbTab & _
              Bar2Group(p.GroupRh & "") & vbTab & _
              Format$(p.DateExpiry, "dd/MMM/yyyy HH:mm") & vbTab & _
              gEVENTCODES(p.PackEvent).Text & vbTab & _
              p.Chart & vbTab & _
              p.PatName & vbTab & _
              p.UserName & vbTab & _
              p.RecordDateTime & vbTab & _
              p.SampleID
200       g.AddItem s
210     End If
220   Next

230   lhistory.Visible = True
240   showinghistory = False

250   If g.Rows > 2 Then
260     g.RemoveItem 1
270   End If

280   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fprodmove", "FillG", intEL, strES

End Sub

Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

      Dim d1 As String
      Dim d2 As String

10    If Not IsDate(g.TextMatrix(Row1, g.col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(g.TextMatrix(Row2, g.col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(g.TextMatrix(Row1, g.col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(g.TextMatrix(Row2, g.col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("D", d2, d1))
150   End If

End Sub

