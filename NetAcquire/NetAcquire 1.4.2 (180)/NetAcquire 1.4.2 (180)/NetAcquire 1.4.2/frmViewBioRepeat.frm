VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewBioRepeat 
   Caption         =   "NetAcquire - Biochemistry Repeats"
   ClientHeight    =   5640
   ClientLeft      =   4605
   ClientTop       =   2565
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7020
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   1125
      Left            =   5730
      Picture         =   "frmViewBioRepeat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3060
      Width           =   1035
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer to Main File"
      Height          =   1125
      Left            =   5730
      Picture         =   "frmViewBioRepeat.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1770
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1125
      Left            =   5730
      Picture         =   "frmViewBioRepeat.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4350
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4965
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   8758
      _Version        =   393216
      Cols            =   5
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Test                  |<Result  |<Units    |<Date/Time               |<Code  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests in RED will be Transfered"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight Tests to be Transferred"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5640
      TabIndex        =   3
      Top             =   510
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewBioRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private pSampleID As String

Private pDiscipline As String
Private Sub FillG()

      Dim S As String
      Dim bioReps As New BIEResults
      Dim BRs As BIEResults
      Dim BR As BIEResult

32260 On Error GoTo FillG_Error

32270 Set BRs = bioReps.Load(pDiscipline, pSampleID, "Repeats", gDONTCARE, gDONTCARE)

32280 g.Rows = 2
32290 g.AddItem ""
32300 g.RemoveItem 1

32310 If Not BRs Is Nothing Then
32320   For Each BR In BRs
32330     S = BR.ShortName & vbTab
32340     If IsNumeric(BR.Result) Then
32350       S = S & FormatNumber(Val(BR.Result), BR.Printformat, vbTrue)
32360     Else
32370       S = S & BR.Result
32380     End If
32390     S = S & vbTab & BR.Units & vbTab & _
              Format$(BR.RunTime, "dd/MM/yy HH:nn:ss") & vbTab & _
              BR.Code
32400     g.AddItem S
32410   Next
32420 End If

32430 If g.Rows > 2 Then
32440   g.RemoveItem 1
32450 End If

32460 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

32470 intEL = Erl
32480 strES = Err.Description
32490 LogError "frmViewBioRepeat", "FillG", intEL, strES

End Sub

Private Sub cmdCancel_Click()

32500 Unload Me

End Sub


Private Sub cmdDelete_Click()

      Dim sql As String

32510 On Error GoTo cmdDelete_Click_Error

32520 If iMsg("Delete All Repeats?" & vbCrLf & _
              "You will not be able to undo this process!" & vbCrLf & _
              "Continue?", vbQuestion + vbYesNo) = vbYes Then
        
32530   sql = "DELETE FROM " & pDiscipline & "Repeats WHERE " & _
              "SampleID = '" & pSampleID & "'"

32540   Cnxn(0).Execute sql

32550   Unload Me

32560 End If

32570 Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

32580 intEL = Erl
32590 strES = Err.Description
32600 LogError "frmViewBioRepeat", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub cmdTransfer_Click()

      Dim y As Integer
      Dim sql As String
      Dim Code As String
          ' Dim TempResult As String'

32610 On Error GoTo cmdTransfer_Click_Error

32620 g.Col = 0
32630 For y = 1 To g.Rows - 1
32640   g.row = y
32650   If g.CellBackColor = vbRed Then
32660     Code = g.TextMatrix(y, 4)
      '   TempResult = g.TextMatrix(Y, 1)'
32670     sql = "DELETE FROM " & pDiscipline & "Results WHERE " & _
                "SampleID = '" & pSampleID & "' " & _
                "AND Code = '" & Code & "' "
32680     Cnxn(0).Execute sql
          
32690     sql = "INSERT INTO " & pDiscipline & "Results " & _
                "SELECT TOP 1 * FROM " & pDiscipline & "Repeats WHERE " & _
                "SampleID = '" & pSampleID & "' " & _
                "AND Code = '" & Code & "' " & _
                "AND RunTime = '" & Format$(g.TextMatrix(y, 3), "dd/MMM/yyyy HH:nn:ss") & "' "
                
32700     Cnxn(0).Execute sql
          
32710     sql = "DELETE FROM " & pDiscipline & "Repeats WHERE " & _
                "SampleID = '" & pSampleID & "' " & _
                "AND Code = '" & Code & "' AND RunTime = '" & Format$(g.TextMatrix(y, 3), "dd/MMM/yyyy HH:nn:ss") & "' "
32720     Cnxn(0).Execute sql
32730   End If
32740 Next
          
32750 Unload Me

32760 Exit Sub

cmdTransfer_Click_Error:

      Dim strES As String
      Dim intEL As Integer

32770 intEL = Erl
32780 strES = Err.Description
32790 LogError "frmViewBioRepeat", "cmdTransfer_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

32800 If Activated Then Exit Sub

32810 Activated = True

32820 FillG

End Sub

Private Sub Form_Load()

32830 Activated = False
32840 g.ColWidth(4) = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

32850 Activated = False

End Sub

Private Sub g_Click()

      Dim y As Integer

32860 On Error GoTo g_Click_Error

32870 If g.MouseRow = 0 Then Exit Sub

32880 cmdTransfer.Visible = False

32890 g.Col = 0
32900 If g.CellBackColor = vbRed Then
32910   g.CellBackColor = 0
32920 Else
32930   g.CellBackColor = vbRed
32940   cmdTransfer.Visible = True
32950   Exit Sub
32960 End If

32970 For y = 1 To g.Rows - 1
32980   g.row = y
32990   If g.CellBackColor = vbRed Then
33000     cmdTransfer.Visible = True
33010     Exit For
33020   End If
33030 Next

33040 Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

33050 intEL = Erl
33060 strES = Err.Description
33070 LogError "frmViewBioRepeat", "g_Click", intEL, strES

End Sub



Public Property Let SampleID(ByVal sNewValue As String)

33080 pSampleID = sNewValue

End Property


Public Property Let Discipline(ByVal sNewValue As String)

33090 pDiscipline = sNewValue

End Property
