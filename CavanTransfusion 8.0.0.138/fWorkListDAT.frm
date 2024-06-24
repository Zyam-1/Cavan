VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fWorkListDAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D.A.T. Worklist"
   ClientHeight    =   6855
   ClientLeft      =   480
   ClientTop       =   525
   ClientWidth     =   5355
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "fWorkListDAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6855
   ScaleWidth      =   5355
   Begin VB.CommandButton bTransfer 
      Caption         =   "Transfer to Main Files"
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton bdelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   150
      Width           =   1035
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5685
      Left            =   210
      TabIndex        =   3
      Top             =   660
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   10028
      _Version        =   393216
      Cols            =   3
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
      FormatString    =   "<Lab No.              |<Patient Name                  |<P.I.D.        "
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   5
      Top             =   6540
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fWorkListDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

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

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print "D.A.T. Worklist " & Format(Now, "dd/mm/yyyy")
40    Printer.Print
50    Printer.Print "Lab No."; Tab(10); "Patient Name"; Tab(35);
60    Printer.Print "P.I.D."
70    For Y = 1 To g.Rows - 1
80      g.row = Y
90      g.col = 0
100     Printer.Print g; 'lab#
110     g.col = 1
120     Printer.Print Tab(10); g; 'name
130     g.col = 2
140     Printer.Print Tab(35); g 'pid
150   Next
160   Printer.EndDoc

170   For Each Px In Printers
180     If Px.DeviceName = OriginalPrinter Then
190       Set Printer = Px
200       Exit For
210     End If
220   Next

End Sub

Private Sub FillG()

      Dim s As String
      Dim sn As Recordset
      Dim final As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "select distinct labnumber from patientdetails where " & _
            "hold = 1 " & _
            "and requestfrom ='D' " & _
            "order by labnumber"
60    Set sn = New Recordset
70    RecOpenServerBB 0, sn, sql

80    Do While Not sn.EOF
90      sql = "select labnumber, name, patnum, " & _
              "hold from details where " & _
              "requestfrom = 'D' " & _
              "and labnumber = '" & sn!LabNumber & "' " & _
              "order by datetime"
100     Set final = New Recordset
110     RecOpenServerBB 0, final, sql
120     final.MoveLast
130     If final!Hold Then
140       s = final!LabNumber & vbTab
150       s = s & final!Name & vbTab
160       s = s & final!Patnum & vbTab
170       g.AddItem s, 1
180       g.Refresh
190     End If
200     sn.MoveNext
210   Loop
220   g.Refresh

230   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "fWorkListDAT", "FillG", intEL, strES, sql


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
150   s = s & "Transfer to main files?"
160   Answer = iMsg(s, vbYesNo + vbQuestion)
170   If TimedOut Then Unload Me: Exit Sub
180   If Answer = vbNo Then Exit Sub

190   sql = "select * from patientdetails where " & _
            "labnumber = '" & labnum2find & "' " & _
            "order by datetime"
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
290   LogError "fWorkListDAT", "bTransfer_Click", intEL, strES, sql

End Sub



Private Sub Form_Load()

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************
End Sub

Private Sub g_DblClick()

      Dim sn As Recordset
      Dim sql As String
      Dim labnum2find As String
      Dim tempfg As String

10    On Error GoTo g_DblClick_Error

20    g.col = 0
30    If g.row = 0 Then Exit Sub
40    If Trim$(g) = "" Then Beep: Exit Sub

50    labnum2find = g

60    sql = "select * from patientdetails where labnumber ='" & labnum2find & "'"
70    Set sn = New Recordset
80    RecOpenServerBB 0, sn, sql
90    sn.MoveLast
100   With frmxmatch
110   .tLabNum = sn!LabNumber & ""
120   .txtChart = sn!Patnum & ""
130   .txtName = sn!Name & ""
140   .tMaiden = sn!maiden & ""
150   .tAddr(0) = sn!Addr1 & ""
160   .tAddr(1) = sn!Addr2 & ""
170   .tAddr(2) = sn!Addr3 & ""
180   .tAddr(3) = sn!addr4 & ""
190   .cWard.Text = sn!Ward & ""
200   .cClinician.Text = sn!Clinician & ""
210   .cConditions.Text = sn!Conditions & ""
220   .cProcedure.Text = sn!Procedure & ""
230   .cSpecial.Text = sn!specialprod & ""
240   grh2image sn!PrevGroup & "", sn!previousrh & ""

250   .lSex = sn!Sex & ""
260   If Not IsNull(sn!DoB) Then
270     .tDoB = Format(sn!DoB, "dd/mm/yyyy")
280   Else
290     .tDoB = ""
300   End If
310   .tAge = sn!Age & ""
320   .tComment = StripComment(sn!Comment & "")

330   tempfg = sn!fgpattern & ""

340   .lstfg.ListIndex = Group2Index(sn!fGroup & "")
350   .lblsuggestfg = sn!fgsuggest & ""

360   If Not IsNull(sn!edd) Then
370     .tedd = Format(sn!edd, "dd/mm/yyyy")
380   Else
390     .tedd = ""
400   End If

410   .cmdSave.Enabled = False: .bHold.Enabled = False
420   End With
430   Unload Me

440   Exit Sub

g_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "fWorkListDAT", "g_DblClick", intEL, strES, sql

End Sub

