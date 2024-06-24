VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form fWorkListAN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ante-Natal Worklist"
   ClientHeight    =   7155
   ClientLeft      =   2040
   ClientTop       =   2610
   ClientWidth     =   5325
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "fWorkListAN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7155
   ScaleWidth      =   5325
   Begin VB.CommandButton bTransfer 
      Caption         =   "Transfer to Main Files"
      Height          =   615
      Left            =   3870
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5685
      Left            =   240
      TabIndex        =   3
      Top             =   930
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
      FormatString    =   "<Lab No.             |<Patient Name                  |<P.I.D.        "
   End
   Begin VB.CommandButton bdelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   510
      Width           =   1035
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Top             =   270
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   270
      TabIndex        =   5
      Top             =   6870
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fWorkListAN"
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

30    Printer.Print "Ante-Natal Worklist " & Format(Now, "dd/mm/yyyy")
40    Printer.Print
50    Printer.Print "Lab # "; Tab(10); "Name"; Tab(40); "PID"
60    For Y = 1 To g.Rows - 1
70      g.row = Y
80      g.col = 0
90      Printer.Print g;
100     g.col = 1
110     Printer.Print Tab(10); g;
120     g.col = 2
130     Printer.Print Tab(40); g
140   Next
150   Printer.EndDoc

160   For Each Px In Printers
170     If Px.DeviceName = OriginalPrinter Then
180       Set Printer = Px
190       Exit For
200     End If
210   Next

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
            "and requestfrom = 'A' " & _
            "order by labnumber"

60    Set sn = New Recordset
70    RecOpenServerBB 0, sn, sql
80    Do While Not sn.EOF
90      sql = "select labnumber, name, " & _
              "patnum, hold from patientdetails where " & _
              "requestfrom = 'A' " & _
              "and labnumber = '" & sn!LabNumber & "' " & _
              "order by datetime"
100     Set final = New Recordset
110     RecOpenServerBB 0, final, sql
120     final.MoveLast
130     If final!Hold Then
140       s = final!LabNumber & vbTab
150       s = s & final!Name & vbTab
160       s = s & final!Patnum & vbTab
170       g.AddItem s
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
260   LogError "fWorkListAN", "FillG", intEL, strES, sql


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
            "labnumber = '" & labnum2find & "'"
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
290   LogError "fWorkListAN", "bTransfer_Click", intEL, strES, sql


End Sub



Private Sub Form_Load()


      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        FillG
      '**************************************
End Sub

Private Sub g_DblClick()

      Dim mt As Recordset
      Dim sql As String
      Dim labnum2find As String
      Dim n As Integer
      Dim f As Form

10    Set f = frmxmatch

20    g.col = 0
30    If g.row = 0 Then Exit Sub
40    If Trim$(g) = "" Then Beep: Exit Sub

50    labnum2find = g

60    sql = "Select * from PatientDetails where " & _
            "LabNumber = '" & labnum2find & "'"
70    RecOpenServerBB 0, mt, sql

80    f.txtChart = mt("patnum") & ""
90    f.txtName = mt("name") & ""
100   f.tMaiden = mt("maiden") & ""
110   f.tAddr(0) = mt("addr1") & ""
120   f.tAddr(1) = mt("addr2") & ""
130   f.tAddr(2) = mt("addr3") & ""
140   f.tAddr(3) = mt("addr4") & ""
150   f.lwcc(0).Text = mt("ward") & ""
160   f.lwcc(1).Text = mt("clinician") & ""
170   f.lwcc(2).Text = mt("conditions") & ""
180   f.lwcc(3).Text = mt("procedure") & ""
190   f.lwcc(4).Text = mt("specialprod") & ""

200   grh2image mt("prevgroup") & "", mt("previousrh") & ""

210   If Not IsNull(mt!DoB) Then
220     f.txtDoB = Format(mt!DoB, "dd/mm/yyyy")
230   Else
240     f.txtDoB = ""
250   End If
260   f.tComment = StripComment(mt("comment") & "")

270   f.lstfg.ListIndex = Group2Index(mt!fGroup & "")
280   f.lblsuggestfg = mt!fgsuggest & ""

290   f.grdauto.row = 0
300   f.grdauto.col = 0
310   f.grdauto = Left$(mt("autoant") & "", 1)
320   f.grdauto.col = 1
330   f.grdauto = Right$(mt("autoant") & "", 1)

340   f.txtantilot = mt("anti3lot") & ""

350   For n = 0 To 2
360     f.grdanti3c.col = n
370     f.grdanti3e.col = n
380     f.grdanti3c = Mid$(mt("anti3c") & "", n + 1, 1)
390     f.grdanti3e = Mid$(mt("anti3e") & "", n + 1, 1)
400   Next
410   f.lblantibody = mt("anti3reported") & ""
420   f.lident = mt("aids") & ""
430   f.tident = mt("aidr") & ""

440   f.tLabNum = mt("labnumber") & ""
450   f.tptrans = mt("prevtrans") & ""
460   f.tpreaction = mt("prevreact") & ""
470   If Not IsNull(mt!edd) Then
480     f.tedd = Format(mt!edd, "dd/mm/yyyy")
490   Else
500     f.tedd = ""
510   End If
520   f.tppreg = mt("prevpreg") & ""

530   f.cmdSave.Enabled = False: f.bHold.Enabled = False

540   Unload Me

550   Exit Sub

g_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "fWorkListAN", "g_DblClick", intEL, strES, sql

End Sub

Private Sub g_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    g.Drag

End Sub



