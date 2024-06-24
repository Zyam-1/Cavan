VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCoagRepeats 
   Caption         =   "NetAcquire - Coagulation Repeats"
   ClientHeight    =   2730
   ClientLeft      =   5595
   ClientTop       =   3540
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4230
   Begin VB.CommandButton bDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   525
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   690
      Width           =   1245
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   2670
      TabIndex        =   2
      Top             =   2160
      Width           =   1245
   End
   Begin VB.CommandButton bCopy 
      Caption         =   "Copy to &Result"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2670
      TabIndex        =   1
      Top             =   90
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2625
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4630
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Parameter            |<Result    "
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
End
Attribute VB_Name = "frmCoagRepeats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private mEditForm As Form
Private Sub bcancel_Click()

21150     mEditForm.LoadCoagulation

21160     Unload Me

End Sub


Private Sub bCopy_Click()

          Dim n As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim TestCode As String

21170     On Error GoTo bCopy_Click_Error

21180     For n = 1 To g.Rows - 1
21190         g.row = n
21200         g.Col = 0
21210         If g.CellBackColor = vbYellow Then
21220             TestCode = CoagCodeForTestName(g.TextMatrix(n, 0))
21230             sql = "Select * from CoagResults where " & _
                      "SampleID = '" & mSampleID & "' " & _
                      "and Code = '" & TestCode & "'"
21240             Set tb = New Recordset
21250             RecOpenClient 0, tb, sql
21260             With tb
21270                 If .EOF Then .AddNew
21280                 !Code = TestCode
21290                 !Printed = False
21300                 !Result = g.TextMatrix(g.row, 1)
21310                 !Rundate = Format$(Now, "dd/mmm/yyyy")
21320                 !RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
21330                 !SampleID = mSampleID
21340                 !Valid = 1
21350                 .Update
21360             End With
          
21370             sql = "delete from CoagRepeats where " & _
                      "SampleID = '" & mSampleID & "' " & _
                      "and Code = '" & TestCode & "'"
21380             Set tb = New Recordset
21390             RecOpenClient 0, tb, sql
21400         End If
21410     Next

21420     FillG

21430     bCopy.Enabled = False

21440     Exit Sub

bCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

21450     intEL = Erl
21460     strES = Err.Description
21470     LogError "fCoagRepeats", "bCopy_Click", intEL, strES, sql

End Sub

Private Sub bDelete_Click()
          
          Dim sql As String

21480     On Error GoTo bDelete_Click_Error

21490     StrEvent = "Delete Coag Repeats"
21500     LogEvent StrEvent, "fCoagRepeats", "bDelete_Click"

21510     sql = "delete from CoagRepeats where " & _
              "SampleID = '" & mSampleID & "'"
21520     Cnxn(0).Execute sql

21530     mEditForm.LoadCoagulation

21540     Unload Me

21550     Exit Sub

bDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

21560     intEL = Erl
21570     strES = Err.Description
21580     LogError "fCoagRepeats", "bDelete_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

21590     FillG

End Sub
Private Sub FillG()

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim s As String

21600     Set CRs = CRs.Load(mSampleID, gDONTCARE, gDONTCARE, "Repeats")

21610     g.Rows = 2
21620     g.AddItem ""
21630     g.RemoveItem 1

21640     For Each CR In CRs
21650         If CR.InUse Then
21660             s = CR.TestName & vbTab & CR.Result
21670             g.AddItem s
21680         End If
21690     Next
        
21700     If g.Rows > 2 Then
21710         g.RemoveItem 1
21720     End If

End Sub

Private Sub g_Click()

21730     If g.MouseRow = 0 Then Exit Sub

21740     g.Col = 0
21750     If g.CellBackColor = vbYellow Then
21760         g.CellBackColor = 0
21770         g.Col = 1
21780         g.CellBackColor = 0
21790     Else
21800         g.CellBackColor = vbYellow
21810         g.Col = 1
21820         g.CellBackColor = vbYellow
21830     End If

21840     bCopy.Enabled = True

End Sub



Public Property Let SampleID(ByVal sNewValue As String)

21850     mSampleID = sNewValue

End Property

Public Property Let EditForm(ByVal f As Form)

21860     Set mEditForm = f

End Property
