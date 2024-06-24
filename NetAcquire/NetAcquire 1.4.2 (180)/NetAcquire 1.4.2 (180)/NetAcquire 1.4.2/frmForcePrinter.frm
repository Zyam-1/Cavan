VERSION 5.00
Begin VB.Form frmForcePrinter 
   Caption         =   "NetAcquire "
   ClientHeight    =   4425
   ClientLeft      =   2085
   ClientTop       =   2415
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   825
      Left            =   810
      Picture         =   "frmForcePrinter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   4110
      Picture         =   "frmForcePrinter.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3300
      Width           =   1275
   End
   Begin VB.ListBox lAvailable 
      Height          =   2325
      IntegralHeight  =   0   'False
      Left            =   600
      TabIndex        =   0
      Top             =   750
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmForcePrinter.frx":0CD4
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   150
      Width           =   4755
   End
End
Attribute VB_Name = "frmForcePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FromForm As Form

Private Sub bcancel_Click()

41600     Unload Me

End Sub

Private Sub cmdSave_Click()

41610     FromForm.PrintToPrinter = lAvailable

41620     Unload Me

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

41630     On Error GoTo Form_Load_Error

41640     lAvailable.Clear
41650     lAvailable.AddItem "Automatic Selection"
41660     lAvailable.AddItem ""

41670     sql = "Select * from InstalledPrinters"
41680     Set tb = New Recordset
41690     RecOpenServer 0, tb, sql

41700     Do While Not tb.EOF
41710         lAvailable.AddItem tb!PrinterName & ""
41720         tb.MoveNext
41730     Loop

41740     Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

41750     intEL = Erl
41760     strES = Err.Description
41770     LogError "frmForcePrinter", "Form_Load", intEL, strES, sql


End Sub



Public Property Let From(ByVal frmNewValue As Form)

41780     Set FromForm = frmNewValue

End Property

Private Sub lAvailable_DblClick()

41790     FromForm.PrintToPrinter = lAvailable

41800     Unload Me

End Sub


