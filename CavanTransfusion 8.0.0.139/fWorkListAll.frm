VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fWorkListAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worklists"
   ClientHeight    =   5100
   ClientLeft      =   270
   ClientTop       =   690
   ClientWidth     =   10965
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
   Icon            =   "fWorkListAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   10965
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
      Height          =   525
      Left            =   9540
      TabIndex        =   7
      Top             =   600
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
      Left            =   6690
      TabIndex        =   6
      Top             =   600
      Width           =   1245
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
      Left            =   8130
      TabIndex        =   5
      Top             =   600
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      DragIcon        =   "fWorkListAll.frx":08CA
      Height          =   3285
      Left            =   75
      TabIndex        =   4
      Top             =   1230
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   5794
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
      FormatString    =   $"fWorkListAll.frx":0D0C
   End
   Begin VB.CommandButton bprint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2670
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2670
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   8
      Top             =   4740
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      Caption         =   "Lab No:"
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
      Left            =   630
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label llabno 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lab No:"
      Height          =   285
      Left            =   630
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "fWorkListAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bConvertGH_Click()

      Dim ds As Recordset
      Dim sql As String
      Dim s As String
      Dim labnum2find As String

10    On Error GoTo bConvertGH_Click_Error

20    If g.row = 0 Then Exit Sub

30    g.col = 3
40    If g <> "X-Match" Then Exit Sub

50    g.col = 0
60    If Trim$(g) = "" Then
70      iMsg "Nothing selected", vbCritical
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

110   labnum2find = g
120   s = "Lab Number " & labnum2find & vbCrLf
130   g.col = 1
140   s = s & "Name " & g & vbCrLf
150   g.col = 2
160   s = s & "Chart No " & g & vbCrLf
170   s = s & "Convert to Group/Hold ?"
180   Answer = iMsg(s, vbYesNo + vbQuestion)
190   If TimedOut Then Unload Me: Exit Sub
200   If Answer = vbNo Then Exit Sub

210   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "'"
220   Set ds = New Recordset
230   RecOpenServerBB 0, ds, sql
240   ds!requestfrom = "G"
250   ds.Update
  
260   FillG

270   Exit Sub

bConvertGH_Click_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "fWorkListAll", "bConvertGH_Click", intEL, strES, sql


End Sub

Private Sub bConvertXM_Click()

      Dim mt As Recordset
      Dim s As String
      Dim labnum2find As String
      Dim sql As String

10    On Error GoTo bConvertXM_Click_Error

20    If g.row = 0 Then Exit Sub

30    g.col = 3
40    If g <> "G & H" Then Exit Sub

50    g.col = 0
60    If Trim$(g) = "" Then
70      iMsg "Nothing selected", vbCritical
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

110   labnum2find = g
120   s = "Lab Number " & labnum2find & vbCrLf
130   g.col = 1
140   s = s & "Name " & g & vbCrLf
150   g.col = 2
160   s = s & "Chart No " & g & vbCrLf
170   g.col = 6
180   s = s & "Convert to Cross-Match?"
190   Answer = iMsg(s, vbYesNo + vbQuestion)
200   If TimedOut Then Unload Me: Exit Sub
210   If Answer = vbNo Then Exit Sub

220   sql = "Select * from patientdetails where " & _
            "LabNumber = '" & labnum2find & "'"
230   Set mt = New Recordset
240   RecOpenServerBB 0, mt, sql
250   mt!requestfrom = "X"
260   mt.Update
  
270   FillG

280   Exit Sub

bConvertXM_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fWorkListAll", "bConvertXM_Click", intEL, strES, sql

End Sub

Private Sub bprint_Click()

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "General Worklist " & Format(Now, "dd/mm/yyyy")
40    Printer.Print
50    Printer.Print "Lab No.";
60    Printer.Print Tab(10); "Patient Name";
70    Printer.Print Tab(35); "P.I.D.";
80    Printer.Print Tab(45); "From";
90    Printer.Print Tab(53); "Reqd. by";
100   Printer.Print Tab(65); "Product"

110   With g
120     For Y = 1 To .Rows - 1
130       .row = Y
140       .col = 0
150       Printer.Print .Text; 'lab#
160       .col = 1
170       Printer.Print Tab(10); .Text; 'name
180       .col = 2
190       Printer.Print Tab(35); .Text; 'pid
200       .col = 3
210       Printer.Print Tab(45); .Text; 'from
220       .col = 4
230       Printer.Print Tab(53); .Text; 'reqd by
240       .col = 5
250       Printer.Print Tab(62); .Text; 'am/pm
260       .col = 6
270       Printer.Print Tab(65); .Text 'product
280     Next
290   End With

300   Printer.EndDoc

310   For Each Px In Printers
320     If Px.DeviceName = OriginalPrinter Then
330       Set Printer = Px
340       Exit For
350     End If
360   Next

End Sub

Private Sub FillG()

      Dim s As String
      Dim sn As Recordset
      Dim se As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    lLabNo.Visible = True
30    l.Visible = True

40    g.Rows = 2
50    g.AddItem ""
60    g.RemoveItem 1

70    sql = "select distinct labnumber " & _
            "from patientdetails where hold = 1"
80    Set se = New Recordset
90    RecOpenServerBB 0, se, sql

100   Do While Not se.EOF
110     lLabNo = se!LabNumber
120     lLabNo.Refresh
130     sql = "select * from patientdetails where " & _
              "labnumber = '" & lLabNo & "'"
140     Set sn = New Recordset
150     RecOpenServerBB 0, sn, sql
160     sn.MoveLast
170     If sn!Hold Then
180       s = Trim$(sn!LabNumber) & vbTab & _
              Trim$(sn!Name & "") & vbTab & _
              Trim$(sn!Patnum & "") & vbTab & _
              From2Text(sn!requestfrom & "") & vbTab
190       If Not IsNull(sn!daterequired) Then
200         s = s & Format(sn!daterequired, "dd/mm/yyyy")
210       End If
220       s = s & vbTab & _
              IIf(sn!ampm = 1, "PM", "AM") & vbTab & _
              ProductWordingFor(sn!BarCode & " ")
230       g.AddItem s, 1
240     End If
250     se.MoveNext
260   Loop

270   lLabNo.Visible = False
280   l.Visible = False

290   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "fWorkListAll", "FillG", intEL, strES, sql

End Sub

Private Sub bTransfer_Click()

      Dim ds As Recordset
      Dim sql As String
      Dim s As String
      Dim labnum2find As String
      Dim product2find As String

10    On Error GoTo bTransfer_Click_Error

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
150   g.col = 6
160   product2find = g
170   s = s & "Transfer to main files?"
180   Answer = iMsg(s, vbYesNo + vbQuestion)
190   If TimedOut Then Unload Me: Exit Sub
200   If Answer = vbNo Then Exit Sub

210   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "' and " & _
            "barcode ='" & ProductBarCodeFor(product2find) & "'"
220   Set ds = New Recordset
230   RecOpenServerBB 0, ds, sql
240   Do While Not ds.EOF
250     ds!Hold = False
260     ds.Update
270     ds.MoveNext
280   Loop
  
290   FillG

300   Exit Sub

bTransfer_Click_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "fWorkListAll", "bTransfer_Click", intEL, strES, sql


End Sub



Private Sub Form_Load()


      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************
End Sub

Private Sub g_Click()

      Dim sql As String
      Dim mt As Recordset
      Dim labnum2find As String
      Dim pc2find As String

10    On Error GoTo g_Click_Error

20    g.col = 0
30    If g.row = 0 Then Exit Sub
40    If Trim$(g) = "" Then Beep: Exit Sub

50    labnum2find = g
60    g.col = 6
70    pc2find = ProductBarCodeFor(g)

80    sql = "Select * from patientdetails where " & _
            "LabNumber = '" & labnum2find & "' " & _
            "and BarCode = '" & pc2find & "'"
90    Set mt = New Recordset
100   RecOpenServerBB 0, mt, sql

110   Dept = Val(InStr("XGAD", mt!requestfrom & "")) - 1

120   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "fWorkListAll", "g_Click", intEL, strES, sql

End Sub

