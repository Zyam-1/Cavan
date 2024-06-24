VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fWorkListGH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group & Hold Worklist"
   ClientHeight    =   7125
   ClientLeft      =   285
   ClientTop       =   1125
   ClientWidth     =   10905
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
   Icon            =   "fWorkListGH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7125
   ScaleWidth      =   10905
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
      Height          =   705
      Left            =   120
      TabIndex        =   6
      Top             =   90
      Width           =   4125
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   345
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   1890
         TabIndex        =   8
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
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   270
         TabIndex        =   9
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
   Begin VB.CommandButton bConvertXM 
      Caption         =   "Convert to Cross Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9300
      TabIndex        =   5
      Top             =   240
      Width           =   1245
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
      Height          =   525
      Left            =   7860
      TabIndex        =   4
      Top             =   240
      Width           =   1245
   End
   Begin VB.CommandButton bdeleteold 
      Appearance      =   0  'Flat
      Caption         =   "Delete &Old"
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
      Left            =   5460
      TabIndex        =   3
      Top             =   90
      Width           =   1125
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
      Height          =   375
      Left            =   5460
      TabIndex        =   2
      Top             =   480
      Width           =   1125
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
      Height          =   375
      Left            =   4410
      TabIndex        =   1
      Top             =   300
      Width           =   885
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
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      DragIcon        =   "fWorkListGH.frx":08CA
      Height          =   5865
      Left            =   90
      TabIndex        =   10
      Top             =   900
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   10345
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
      FormatString    =   $"fWorkListGH.frx":0D0C
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   11
      Top             =   6855
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fWorkListGH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bConvertXM_Click()


      Dim ds As Recordset
      Dim s As String
      Dim labnum2find As String
      Dim sql As String

10    On Error GoTo bConvertXM_Click_Error

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
150   s = s & "Convert to X-Match?"
160   Answer = iMsg(s, vbYesNo + vbQuestion)
170   If TimedOut Then Unload Me: Exit Sub
180   If Answer = vbNo Then Exit Sub

190   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "'"
200   Set ds = New Recordset
210   RecOpenServerBB 0, ds, sql
220   ds.MoveLast
230   ds!requestfrom = "X"
240   ds.Update

250   FillG

260   Exit Sub

bConvertXM_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "fWorkListGH", "bConvertXM_Click", intEL, strES, sql

End Sub


Private Sub bdelete_Click()

10    If ConfirmDelete() Then
20      DeleteWorkList Me
30    End If

40    Dept = GROUP_HOLD
50    FillG

End Sub

Private Function ConfirmDelete() As Boolean

10    ConfirmDelete = False
20    Answer = iMsg("Remove all entries from Worklist?", vbYesNo + vbQuestion)
30    If TimedOut Then Unload Me: Exit Function
40    If Answer = vbYes Then
50      ConfirmDelete = True
60    End If

End Function

Private Sub bdeleteold_Click()

      Dim Y As Integer
      Dim sql As String

      Dim ds As Recordset
      Dim removedate As String

10    On Error GoTo bdeleteold_Click_Error

20    removedate = Format(DateAdd("ww", -1, Now), "dd/mm/yyyy")

30    sql = "Remove all entries before " & removedate
40    Answer = iMsg(sql, vbYesNo + vbQuestion)
50    If TimedOut Then Unload Me: Exit Sub
60    If Answer <> vbYes Then Exit Sub

70    For Y = 1 To g.Rows - 1
80      g.row = Y
90      g.col = 0
100     sql = "select * from PatientDetails where " & _
              "labnumber = '" & g & "' " & _
              "and requestfrom = 'G' " & _
              "and hold = 1 " & _
              "and daterequired < '" & removedate & "'"
110     Set ds = New Recordset
120     RecOpenServerBB 0, ds, sql
130     If Not ds.EOF Then
140       ds.MoveLast
150       ds!Hold = False
160       ds.Update
170     End If
180   Next

190   FillG

200   Exit Sub

bdeleteold_Click_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "fWorkListGH", "bdeleteold_Click", intEL, strES, sql

End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "Group & Hold Worklist " & Format(Now, "dd/mm/yyyy")
40    Printer.Print
50    Printer.Print "Lab No."; Tab(10); "Patient Name"; Tab(35);
60    Printer.Print "P.I.D."; Tab(45); "Reqd. by"; Tab(55); "am/pm"
70    For Y = 1 To g.Rows - 1
80      Printer.Print g.TextMatrix(Y, 0); 'lab#
90      Printer.Print Tab(10); g.TextMatrix(Y, 1); 'name
100     Printer.Print Tab(35); g.TextMatrix(Y, 2); 'pid
110     Printer.Print Tab(45); g.TextMatrix(Y, 3); 'reqd by
120     Printer.Print Tab(55); g.TextMatrix(Y, 4) 'am/pm
130   Next
140   Printer.EndDoc

150   For Each Px In Printers
160     If Px.DeviceName = OriginalPrinter Then
170       Set Printer = Px
180       Exit For
190     End If
200   Next

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
            "and requestfrom = 'G' " & _
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
260   LogError "fWorkListGH", "FillG", intEL, strES, sql


End Sub

Private Sub bTransfer_Click()


      Dim ds As Recordset
      Dim sql As String
      Dim s As String
      Dim labnum2find As String

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
150   s = s & "Transfer to main files?"
160   Answer = iMsg(s, vbYesNo + vbQuestion)
170   If TimedOut Then Unload Me: Exit Sub
180   If Answer = vbNo Then Exit Sub

190   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "' " & _
            "and hold = 1 " & _
            "and requestfrom = 'G'"
200   Set ds = New Recordset
210   RecOpenServerBB 0, ds, sql

220   ds.MoveLast
230   ds!Hold = False
240   ds.Update

250   FillG

260   Exit Sub

bTransfer_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "fWorkListGH", "bTransfer_Click", intEL, strES, sql


End Sub

Private Sub Command1_Click()

10    FillG

End Sub

Private Sub Form_Activate()

10
20    FillG

End Sub

Private Sub Form_Load()

10    dtTo = Format(Now + 2, "dd/mm/yyyy")
20    dtFrom = DateAdd("d", -1, Now)

'*****NOTE
    'FillG might be dependent on many components so for any future
    'update in code try to keep FillG on bottom most line of form load.
30        Dept = GROUP_HOLD
40        FillG
'**************************************

End Sub


Private Sub g_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    g.Drag

End Sub

