VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchStatus 
   Caption         =   "Batch Status Amendment"
   ClientHeight    =   5430
   ClientLeft      =   1470
   ClientTop       =   1530
   ClientWidth     =   6585
   Icon            =   "frmBatchStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6585
   Begin VB.TextBox tCurrent 
      Height          =   285
      Left            =   4800
      TabIndex        =   18
      Top             =   1650
      Width           =   1125
   End
   Begin VB.ComboBox cGroup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5100
      TabIndex        =   15
      Text            =   "cGroup"
      Top             =   750
      Width           =   825
   End
   Begin VB.TextBox tVolume 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   12
      Text            =   "200"
      Top             =   1230
      Width           =   555
   End
   Begin VB.TextBox txtBatchNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   780
      Width           =   2775
   End
   Begin VB.ComboBox cmbProduct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Text            =   "cmbProduct"
      Top             =   270
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   3465
      Picture         =   "frmBatchStatus.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4230
      Width           =   1215
   End
   Begin VB.CommandButton bupdate 
      Caption         =   "&Update"
      Height          =   705
      Left            =   1905
      Picture         =   "frmBatchStatus.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtReason 
      Height          =   1065
      Left            =   1260
      TabIndex        =   0
      Top             =   3000
      Width           =   3885
   End
   Begin MSComCtl2.DTPicker dtReceived 
      Height          =   315
      Left            =   1350
      TabIndex        =   8
      Top             =   2040
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92536833
      CurrentDate     =   37294
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   315
      Left            =   1350
      TabIndex        =   9
      Top             =   1650
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92536833
      CurrentDate     =   37294
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   30
      TabIndex        =   19
      Top             =   5190
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current Stock"
      Height          =   195
      Left            =   3780
      TabIndex        =   17
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   4650
      TabIndex        =   16
      Top             =   780
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Unit Volume"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   14
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ml"
      Height          =   195
      Left            =   1950
      TabIndex        =   13
      Top             =   1290
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Received"
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   2130
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Expiry Date"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   1710
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      Height          =   195
      Left            =   750
      TabIndex        =   6
      Top             =   330
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Reason for Amendment"
      Height          =   195
      Index           =   0
      Left            =   1350
      TabIndex        =   3
      Top             =   2790
      Width           =   1665
   End
End
Attribute VB_Name = "frmBatchStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bupdate_Click()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo bupdate_Click_Error

20    If Left$(txtBatchNumber, 1) = "A" And Right$(txtBatchNumber, 1) = "B" Then
30      txtBatchNumber = Mid$(txtBatchNumber, 2, 9)
40    End If

50    If DateDiff("d", Now, dtExpiry) < 0 Then
60      If UserMemberOf = "Managers" Then
70        Answer = iMsg("Expired Expiry Date entered!" & vbCrLf & _
             "Check Expiry Date." & dtExpiry & vbCrLf & _
             "Are you sure the expiry date is correct?", vbYesNo + vbQuestion)
80        If TimedOut Then Unload Me: Exit Sub
90        If Answer = vbNo Then
100   iMsg "Changes not saved.", vbInformation
110   If TimedOut Then Unload Me: Exit Sub
120   Exit Sub
130       End If
140     Else
150       iMsg "Expired Expiry Date entered!" & vbCrLf & "You do not have authorisation to perform this task.", vbCritical
160       If TimedOut Then Unload Me: Exit Sub
170       Exit Sub
180     End If
190   End If
  
200   If Trim$(txtReason) = "" Then
210     iMsg "Enter the reason for the amendment.", vbCritical
220     If TimedOut Then Unload Me: Exit Sub
230     txtReason.SetFocus
240     Exit Sub
250   End If

260   sql = "Select * from BatchProductList where " & _
            "BatchNumber = '" & txtBatchNumber & "'"
270   Set tb = New Recordset
280   RecOpenServerBB 0, tb, sql

290   If tb.EOF Then
300     iMsg "No details found!", vbExclamation
310     If TimedOut Then Unload Me: Exit Sub
320     Exit Sub
330   End If

340   Answer = iMsg("Confirmation Required." & vbCrLf & _
              "Current stock of " & txtBatchNumber & " = " & tCurrent, vbQuestion + vbYesNo)
350   If TimedOut Then Unload Me: Exit Sub
360   If Answer = vbNo Then
370     iMsg "Changes not saved"
380     If TimedOut Then Unload Me: Exit Sub
390     Exit Sub
400   End If

410   Do While Not tb.EOF
420     tb!DateExpiry = Format(dtExpiry, "dd/mmm/yyyy")
430     tb!DateReceived = Format(dtReceived, "dd/mmm/yyyy")
440     tb!UnitVolume = tVolume
450     tb!Group = cGroup
460     tb!Product = cmbProduct
470     tb!CurrentStock = Val(tCurrent)
480     tb.Update
490     tb.MoveNext
500   Loop

510   sql = "Insert into BatchDetails " & _
            "(BatchNumber, UserCode, Date, Bottles, Product, Event, Expiry, Comment) VALUES " & _
            "('" & txtBatchNumber & "', " & _
            " '" & AddTicks(UserCode) & "', " & _
            " '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "', " & _
            " '" & tCurrent & "', " & _
            " '" & cmbProduct & "', " & _
            " 'E', " & _
            " '" & Format(dtExpiry, "dd/MMM/yyyy") & "'," & _
            " '" & AddTicks(txtReason) & "')"
520   CnxnBB(0).Execute sql

530   sql = "Select * from IncidentLog WHERE 0 = 1"
540   Set tb = New Recordset
550   RecOpenServerBB 0, tb, sql
560   tb.AddNew
570   tb!Incident = txtReason
580   tb!DateTime = Format(Now, "dd/mmm/yyyy hh:mm:ss")
590   tb!Technician = UserName
600   tb.Update

610   cmbProduct = ""
620   txtBatchNumber = ""
630   tVolume = ""
640   cGroup = ""
650   tCurrent = ""
660   txtReason = ""

670   Exit Sub

bupdate_Click_Error:

      Dim strES As String
      Dim intEL As Integer

680   intEL = Erl
690   strES = Err.Description
700   LogError "frmBatchStatus", "bupdate_Click", intEL, strES, sql


End Sub

Private Sub cmbProduct_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub Form_Load()

10    FillcmbProduct

20    With cGroup
30      .Clear
40      .AddItem "O"
50      .AddItem "A"
60      .AddItem "B"
70      .AddItem "AB"
80      .ListIndex = -1
90    End With

End Sub


Private Sub FillcmbProduct()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillcmbProduct_Error

20    cmbProduct.Clear
30    sql = "Select * from Lists where " & _
            "ListType = 'B' " & _
            "order by ListOrder"
40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
60    Do While Not tb.EOF
70      cmbProduct.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   Exit Sub

FillcmbProduct_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmBatchStatus", "FillcmbProduct", intEL, strES, sql


End Sub

Private Sub txtBatchNumber_LostFocus()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo txtBatchNumber_LostFocus_Error

20    If Left$(txtBatchNumber, 1) = "A" And Right$(txtBatchNumber, 1) = "B" Then
30      txtBatchNumber = Mid$(txtBatchNumber, 2, 9)
40    End If

50    sql = "Select * from BatchProductList where " & _
            "BatchNumber = '" & txtBatchNumber & "'"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql

80    If tb.EOF Then
90      iMsg "No details found!", vbExclamation
100     If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120   End If

130   dtExpiry = Format(tb!DateExpiry, "dd/mm/yyyy")
140   dtReceived = Format(tb!DateReceived, "dd/mm/yyyy")
150   tVolume = tb!UnitVolume & ""
160   cGroup = tb!Group & ""
170   tCurrent = IIf(IsNull(tb!CurrentStock), 0, Format(tb!CurrentStock))

180   Exit Sub

txtBatchNumber_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmBatchStatus", "txtBatchNumber_LostFocus", intEL, strES, sql

End Sub


