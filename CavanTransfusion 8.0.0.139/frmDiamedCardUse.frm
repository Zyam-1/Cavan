VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiamedCardUse 
   Caption         =   "NetAcquire"
   ClientHeight    =   3570
   ClientLeft      =   3615
   ClientTop       =   345
   ClientWidth     =   5970
   Icon            =   "frmDiamedCardUse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   915
      Left            =   4440
      Picture         =   "frmDiamedCardUse.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "bCancel"
      Top             =   2370
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save F2"
      Height          =   915
      Left            =   3300
      Picture         =   "frmDiamedCardUse.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2370
      Width           =   1005
   End
   Begin VB.Frame fraABO 
      Caption         =   "ABO/D + Reverse Grouping Card"
      Height          =   2025
      Left            =   150
      TabIndex        =   7
      Top             =   180
      Width           =   2715
      Begin VB.TextBox txtABOBatch 
         Height          =   288
         Left            =   90
         TabIndex        =   0
         Top             =   720
         Width           =   2490
      End
      Begin VB.TextBox txtABOExpiry 
         Height          =   288
         Left            =   90
         TabIndex        =   2
         Top             =   1290
         Width           =   1236
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lot Number"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lexpired 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Card Expired"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   1620
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame fraLISS 
      Caption         =   "LISS/Coombs"
      Height          =   2025
      Left            =   3030
      TabIndex        =   4
      Top             =   180
      Width           =   2715
      Begin VB.TextBox txtLISSBatch 
         Height          =   288
         Left            =   90
         TabIndex        =   1
         Top             =   720
         Width           =   2520
      End
      Begin VB.TextBox txtLISSExpiry 
         Height          =   288
         Left            =   90
         TabIndex        =   3
         Top             =   1290
         Width           =   1236
      End
      Begin VB.Label lexpired 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Card Expired"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   1620
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lot Number"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date:"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   1110
         Width           =   870
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   14
      Top             =   3390
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmDiamedCardUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUseABO As Boolean
Private mCardType As CardTypes

Public Enum CardTypes
    cdPatientDetailCard = 0
    cdCrossMatchCard = 1
End Enum

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim strCardType As String

10    On Error GoTo cmdSave_Click_Error


20    If txtABOBatch = "" And txtABOExpiry = "" And txtLISSBatch = "" And txtLISSExpiry = "" Then
    
30        Unload Me
40        Exit Sub
50    End If

60    Select Case mCardType
          Case 0: strCardType = "S"
70        Case 1: strCardType = "X"
80        Case Else: strCardType = ""
90    End Select

100   sql = "Select * from CardValidation where " & _
            "SampleID = '" & frmxmatch.tLabNum & "' And CardType = '" & strCardType & "'"
110   Set tb = New Recordset
120   RecOpenServerBB 0, tb, sql

130   If tb.EOF Then tb.AddNew
  
140   tb!ABOBatch = txtABOBatch
150   tb!LISSBatch = txtLISSBatch
160   If IsDate(txtABOExpiry) Then
170     tb!ABOExpiry = Format$(txtABOExpiry, "dd/mmm/yyyy")
180   End If
190   If IsDate(txtLISSExpiry) Then
200     tb!LISSExpiry = Format$(txtLISSExpiry, "dd/mmm/yyyy")
210   End If
220   tb!SampleID = frmxmatch.tLabNum
230   tb!CardType = strCardType
240   tb!DateTime = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
250   tb!Operator = UserName
260   tb.Update

270   Unload Me

280   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmDiamedCardUse", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

10    If mUseABO Then
20      fraABO.Enabled = True
30    Else
40      fraABO.Enabled = False
50    End If

End Sub

Private Sub txtABOBatch_LostFocus()


      '*************SAMPLE BARCODE VALUE***********
      '5074218020904223737        (19 digits)
      '
      'Batch Number is: 50742.18.02       (first 9 digits)
      'Expiry 4 / 2009                    (month = digit 10,11 and year = digit 12,13)
      'Last 6 digits is the card number and is not logged.
      '********************************************

10    If Len(txtABOBatch) = 19 Then
20      txtABOExpiry = Mid$(txtABOBatch, 10, 4)
30      txtABOExpiry = "28/" & _
                      Right$(txtABOExpiry, 2) & "/" & _
                      "20" & Left$(txtABOExpiry, 2)
40      txtABOBatch = Left$(txtABOBatch, 5) & "." & _
                   Mid$(txtABOBatch, 6, 2) & "." & _
                   Mid$(txtABOBatch, 8, 2)
50    End If

End Sub


Private Sub txtLISSBatch_LostFocus()

10    If Len(txtLISSBatch) = 19 Then
20      txtLISSExpiry = Mid$(txtLISSBatch, 10, 4)
30      txtLISSExpiry = "28/" & _
                      Right$(txtLISSExpiry, 2) & "/" & _
                      "20" & Left$(txtLISSExpiry, 2)
40      txtLISSBatch = Left$(txtLISSBatch, 5) & "." & _
                   Mid$(txtLISSBatch, 6, 2) & "." & _
                   Mid$(txtLISSBatch, 8, 2)
50    End If

End Sub



Public Property Let UseABO(ByVal bNewValue As Boolean)

10    mUseABO = bNewValue

End Property

Public Property Let CardType(ByVal intNewValue As CardTypes)
10        mCardType = intNewValue
    
End Property
