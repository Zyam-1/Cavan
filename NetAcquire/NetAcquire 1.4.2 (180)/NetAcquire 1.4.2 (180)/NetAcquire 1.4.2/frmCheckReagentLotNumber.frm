VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckReagentLotNumber 
   Caption         =   "NetAcquire - Card / Reagent Lot Number"
   ClientHeight    =   2415
   ClientLeft      =   3120
   ClientTop       =   2430
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   3030
      Picture         =   "frmCheckReagentLotNumber.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1470
      Width           =   1425
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   675
      Left            =   1320
      Picture         =   "frmCheckReagentLotNumber.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1470
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   690
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   326172673
      CurrentDate     =   38408
   End
   Begin VB.TextBox txtLotNumber 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   690
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Barcode Scan or Type"
      Height          =   225
      Left            =   270
      TabIndex        =   6
      Top             =   240
      Width           =   3345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry"
      Height          =   195
      Left            =   3810
      TabIndex        =   3
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Lot Number of Monospot Card / Reagent"
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   480
      Width           =   3360
   End
End
Attribute VB_Name = "frmCheckReagentLotNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pAnalyte As String
Private pSampleID As String
Public Property Let Analyte(ByVal strNewValue As String)

13250     pAnalyte = strNewValue

End Property
Public Property Let SampleID(ByVal lngNewValue As Long)

13260     pSampleID = lngNewValue

End Property

Private Sub cmdCancel_Click()

13270     txtLotNumber = ""

13280     Me.Hide

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim tb As Recordset

13290     On Error GoTo cmdSave_Click_Error

13300     If Trim$(txtLotNumber) = "" Then
13310         Exit Sub
13320     End If

13330     If DateDiff("d", dtExpiry, Now) > 0 Then
13340         iMsg "Lot Expired"
13350         Exit Sub
13360     End If

13370     If dtExpiry = Format$(Now, "dd/mm/yyyy") Then
13380         If iMsg("Expiry Today!" & vbCrLf & "Is this correct?", vbQuestion + vbYesNo) = vbNo Then
13390             Exit Sub
13400         End If
13410     End If

13420     sql = "Select * from ReagentLotNumbers where " & _
              "Analyte = 'xxx'"
13430     Set tb = New Recordset
13440     RecOpenServer 0, tb, sql
13450     tb.AddNew
13460     tb!LotNumber = txtLotNumber
13470     tb!Expiry = Format$(dtExpiry, "dd/mmm/yyyy")
13480     tb!Analyte = pAnalyte
13490     tb!EntryDateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
13500     tb!SampleID = pSampleID
13510     tb.Update

13520     Me.Hide

13530     Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

13540     intEL = Erl
13550     strES = Err.Description
13560     LogError "frmCheckReagentLotNumber", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()

          Dim sql As String
          Dim tb As Recordset

13570     On Error GoTo Form_Activate_Error

13580     lblTitle = "Lot Number of " & pAnalyte & " Card / Reagent"

13590     sql = "Select top 1 LotNumber, Expiry from ReagentLotNumbers where " & _
              "Analyte = '" & pAnalyte & "' " & _
              "Order by EntryDateTime desc"
13600     Set tb = New Recordset
13610     RecOpenServer 0, tb, sql
13620     If Not tb.EOF Then
13630         txtLotNumber = tb!LotNumber & ""
13640         dtExpiry = Format$(tb!Expiry, "dd/mm/yyyy")
13650     Else
13660         txtLotNumber = ""
13670         dtExpiry = Format$(Now, "dd/mm/yyyy")
13680     End If

13690     Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

13700     intEL = Erl
13710     strES = Err.Description
13720     LogError "frmCheckReagentLotNumber", "Form_Activate", intEL, strES, sql


End Sub


Public Property Get LotNumber() As String

13730     LotNumber = txtLotNumber

End Property

Public Property Get Expiry() As String

13740     Expiry = dtExpiry

End Property


