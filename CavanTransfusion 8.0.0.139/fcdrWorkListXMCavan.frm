VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fWorkListXM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cross Match Worklist"
   ClientHeight    =   8085
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   12180
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
   Icon            =   "fcdrWorkListXMCavan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8085
   ScaleWidth      =   12180
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
      Height          =   825
      Left            =   4500
      Picture         =   "fcdrWorkListXMCavan.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   150
      Width           =   1215
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
      Height          =   1185
      Left            =   150
      TabIndex        =   6
      Top             =   60
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
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
         Left            =   1860
         Picture         =   "fcdrWorkListXMCavan.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   270
         TabIndex        =   8
         Top             =   750
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   147324929
         CurrentDate     =   38525
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   147324929
         CurrentDate     =   38525
      End
   End
   Begin VB.CommandButton bTransfer 
      Caption         =   "Transfer to Main Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6150
      Picture         =   "fcdrWorkListXMCavan.frx":80D6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1245
   End
   Begin VB.CommandButton bConvertGH 
      Caption         =   "Convert to Group/Hold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7410
      Picture         =   "fcdrWorkListXMCavan.frx":8740
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   150
      Width           =   1245
   End
   Begin VB.CommandButton bdelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9240
      Picture         =   "fcdrWorkListXMCavan.frx":8DAA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   855
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
      Height          =   825
      Left            =   3240
      Picture         =   "fcdrWorkListXMCavan.frx":9414
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   1215
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
      Height          =   825
      Left            =   11040
      Picture         =   "fcdrWorkListXMCavan.frx":9A7E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      DragIcon        =   "fcdrWorkListXMCavan.frx":A0E8
      Height          =   6315
      Left            =   105
      TabIndex        =   4
      Top             =   1320
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   11139
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   $"fcdrWorkListXMCavan.frx":A52A
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   150
      TabIndex        =   12
      Top             =   7680
      Width           =   11895
      _ExtentX        =   20981
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
      Left            =   4470
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "fWorkListXM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    frmxmatch.cmdSave.Enabled = False: frmxmatch.bHold.Enabled = False

20    Unload Me

End Sub

Private Sub bConvertGH_Click()

      Dim ds As Recordset
      Dim sql As String
      Dim s As String
      Dim labnum2find As String
      Dim product2find As String

10    On Error GoTo bConvertGH_Click_Error

20    If g.row = 0 Then Exit Sub

30    g.col = 0
40    If Trim$(g) = "" Then
50      iMsg "Nothing selected", vbCritical
60      If TimedOut Then Unload Me: Exit Sub
70      Exit Sub
80    End If

90    labnum2find = g
100   s = "Lab Number " & labnum2find & vbCrLf
110   g.col = 1
120   s = s & "Name " & g & vbCrLf
130   g.col = 2
140   s = s & "Chart No " & g & vbCrLf
150   s = s & "Convert to Group/Hold ?"
160   Answer = iMsg(s, vbYesNo + vbQuestion)
170   If TimedOut Then Unload Me: Exit Sub
180   If Answer = vbNo Then Exit Sub

190   g.col = 5
200   product2find = ProductBarCodeFor(g)

210   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "'"
220   Set ds = New Recordset
230   RecOpenServerBB 0, ds, sql
240   ds.MoveLast
250   ds!requestfrom = "G"
260   ds.Update
  
270   FillG

280   Exit Sub

bConvertGH_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fWorkListXM", "bConvertGH_Click", intEL, strES, sql

End Sub

Private Sub bdelete_Click()

10    If ConfirmDelete() Then
20      DeleteWorkList Me
30    End If

40    FillG

End Sub

Private Function ConfirmDelete() As Boolean

10    ConfirmDelete = False
20    Answer = iMsg("Remove all entries from Worklist?", vbYesNo + vbQuestion)
30    If TimedOut Then Unload Me: Exit Function
40    If Answer = vbYes Then
50      ConfirmDelete = True
60    End If

End Function

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20        If Not SetFormPrinter() Then Exit Sub
30    Printer.Orientation = vbPRORPortrait

      'Heading
40    Printer.Font.Bold = True
50    Printer.Font.Size = 9
60    Printer.Font.Name = "Courier New"

70    Printer.Print
80    Printer.Print "Cross Match Worklist "
90    Printer.Print "From "; dtFrom.Value; " To "; dtTo.Value
100   For i = 1 To 108
110       Printer.Print "_";
120   Next i
      '****Report body

130   Printer.Print
140   Printer.Print FormatString("Lab No.", 10, "|");
150   Printer.Print FormatString("Patient Name", 30, "|");
160   Printer.Print FormatString("P.I.D.", 10, "|");
170   Printer.Print FormatString("Reqd. by", 16, "|");
180   Printer.Print FormatString("Product", 37, "|")

190   For i = 1 To 108
200       Printer.Print "-";
210   Next i
220   Printer.Print

230   Printer.Font.Bold = False
240   For Y = 1 To g.Rows - 1
250       Printer.Print FormatString(g.TextMatrix(Y, 0), 10, "|");
260       Printer.Print FormatString(g.TextMatrix(Y, 1), 30, "|");
270       Printer.Print FormatString(g.TextMatrix(Y, 2), 10, "|");
280       Printer.Print FormatString(g.TextMatrix(Y, 3), 16, "|");
290       Printer.Print FormatString(g.TextMatrix(Y, 4), 37, "|")
  
300   Next
310   Printer.EndDoc

320   For Each Px In Printers
330     If Px.DeviceName = OriginalPrinter Then
340       Set Printer = Px
350       Exit For
360     End If
370   Next

End Sub

Private Sub FillG()

      Dim s As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "select * from patientdetails where " & _
            "hold = 1 " & _
            "and requestfrom = 'X' " & _
            "and DateTime between '" & Format(dtFrom, "dd/mmm/yyyy") & "' and '" & Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
            "order by labnumber"

60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!LabNumber & vbTab & _
            tb!Name & vbTab & _
            tb!Patnum & vbTab
100     If Not IsNull(tb!daterequired) Then
110       s = s & Format(tb!daterequired, "dd/mm/yyyy")
120     End If
130     s = s & vbTab
140     If Not IsNull(tb!ampm) Then
150       s = s & IIf(tb!ampm, "PM", "AM") & vbTab & _
              ProductWordingFor(tb!BarCode & "")
160     End If
170     g.AddItem s
180     tb.MoveNext
190   Loop

200   If g.Rows > 2 Then
210     g.RemoveItem 1
220   End If

230   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "fWorkListXM", "FillG", intEL, strES, sql


End Sub

Private Sub bTransfer_Click()

      Dim ds As Recordset
      Dim sql As String
      Dim s As String
      Dim labnum2find As String
      Dim product2find As String

10    On Error GoTo bTransfer_Click_Error

20    g.col = 0
30    If g.row = 0 Then Exit Sub
40    If Trim$(g) = "" Then
50      iMsg "Nothing selected", vbExclamation
60      If TimedOut Then Unload Me: Exit Sub
70      Exit Sub
80    End If

90    labnum2find = g
100   s = "Lab Number " & labnum2find & vbCrLf
110   g.col = 1
120   s = s & "Name " & g & vbCrLf
130   g.col = 2
140   s = s & "Chart No " & g & vbCrLf
150   g.col = g.Cols - 1
160   product2find = g
170   s = s & "Transfer to main files?"
180   Answer = iMsg(s, vbQuestion + vbYesNo)
190   If TimedOut Then Unload Me: Exit Sub
200   If Answer = vbNo Then Exit Sub

210   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "' " & _
            "and requestfrom = 'X'"
220   Set ds = New Recordset
230   RecOpenServerBB 0, ds, sql
240   ds!Hold = False
250   ds.Update

260   Dept = XMATCH
270   FillG

280   Exit Sub

bTransfer_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fWorkListXM", "bTransfer_Click", intEL, strES, sql

End Sub




Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Cross Match Worklist " & vbCr
20    strHeading = strHeading & "From " & dtFrom.Value & " To " & dtTo.Value & vbCr
30    strHeading = strHeading & " " & vbCr
40    ExportFlexGrid g, Me, strHeading

End Sub

Private Sub Command1_Click()

10    FillG

End Sub



Private Sub Form_Load()

10    dtTo = Format(Now + 2, "dd/mm/yyyy")
20    dtFrom = DateAdd("d", -1, Now)

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
30        Dept = XMATCH
40        FillG
      '**************************************

End Sub


Private Sub g_DblClick()

      Dim sn As Recordset
      Dim sql As String
      Dim labnum2find As String
      Dim pc2find As String

10    On Error GoTo g_DblClick_Error

20    g.col = 0
30    If g.row = 0 Then Exit Sub
40    If Trim$(g) = "" Then Beep: Exit Sub

50    labnum2find = g
60    g.col = 5
70    pc2find = ProductBarCodeFor(g)

80    sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "' " & _
            "and requestfrom = 'X' " & _
            "and hold = 1 " & _
            "order by datetime"

90    Set sn = New Recordset
100   RecOpenServerBB 0, sn, sql
110   sn.MoveLast

120   With frmxmatch
130     .txtChart = sn!Patnum & ""
140     .txtName = sn!Name & ""
150     .tAddr(0) = sn!Addr1 & ""
160     .tAddr(1) = sn!Addr2 & ""
170     .cWard = sn!Ward & ""
180     .cClinician = sn!Clinician & ""
190     .cConditions = sn!Conditions & ""
200     .cSpecial = sn!specialprod & ""

210     grh2image Trim$(Left$(sn!PrevGroup & "", 2)), sn!previousrh & ""
220     .lSex = sn!Sex & ""
230     If Not IsNull(sn!DoB) Then
240       .tDoB = Format(sn!DoB, "dd/mm/yyyy")
250     Else
260       .tDoB = ""
270     End If
280     .tComment = StripComment(sn!Comment & "")
290     .lblsuggestfg = sn!fgsuggest & ""
300     .lstfg.ListIndex = Group2Index(sn!fGroup & "")

310     .tLabNum = sn!LabNumber & ""
320     .tident = sn!AIDR & ""
330     If Not IsNull(sn!edd) Then
340       .tedd = Format(sn!edd, "dd/mm/yyyy")
350     Else
360       .tedd = ""
370     End If

380     .cmdSave.Enabled = False
390     .bHold.Enabled = False

400   End With

410   Unload Me

420   Exit Sub

g_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "fWorkListXM", "g_DblClick", intEL, strES, sql

End Sub

Private Sub g_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    g.Drag

End Sub

