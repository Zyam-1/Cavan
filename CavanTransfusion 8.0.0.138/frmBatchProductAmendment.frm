VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchProductAmendment 
   Caption         =   "Batch Status Amendment"
   ClientHeight    =   5820
   ClientLeft      =   1470
   ClientTop       =   1530
   ClientWidth     =   8880
   Icon            =   "frmBatchProductAmendment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   8880
   Begin VB.Frame Frame2 
      Height          =   2955
      Left            =   240
      TabIndex        =   11
      Top             =   2250
      Width           =   6765
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   1065
         Left            =   5220
         Picture         =   "frmBatchProductAmendment.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "cmdSave"
         Top             =   1620
         Width           =   1245
      End
      Begin VB.TextBox txtReason 
         Height          =   1065
         Left            =   330
         TabIndex        =   14
         Top             =   1620
         Width           =   4725
      End
      Begin VB.TextBox txtVolume 
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
         Left            =   4350
         TabIndex        =   13
         Text            =   "200"
         Top             =   795
         Width           =   705
      End
      Begin VB.ComboBox cmbGroup 
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
         Left            =   360
         TabIndex        =   12
         Text            =   "cmbGroup"
         Top             =   780
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtExpiry 
         Height          =   315
         Left            =   2310
         TabIndex        =   15
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You may change the Group, Unit Volume or Expiry Date for this Batch"
         Height          =   285
         Left            =   750
         TabIndex        =   20
         Top             =   0
         Width           =   5325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Reason for Amendment"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   19
         Top             =   1380
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date"
         Height          =   195
         Left            =   2700
         TabIndex        =   18
         Top             =   570
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Unit Volume (ml)"
         Height          =   195
         Index           =   1
         Left            =   4110
         TabIndex        =   17
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   690
         TabIndex        =   16
         Top             =   570
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6765
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
         Left            =   1560
         TabIndex        =   4
         Text            =   "cmbProduct"
         Top             =   300
         Width           =   4575
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
         Left            =   1560
         TabIndex        =   3
         Top             =   810
         Width           =   4305
      End
      Begin VB.Label lblDateReceived 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "88/88/8888"
         Height          =   285
         Left            =   4650
         TabIndex        =   10
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lblCurrentStock 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Batch Number"
         Height          =   195
         Left            =   510
         TabIndex        =   7
         Top             =   870
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date Received"
         Height          =   195
         Left            =   3510
         TabIndex        =   6
         Top             =   1365
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Current Stock"
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   1365
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   7290
      Picture         =   "frmBatchProductAmendment.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "cmdCancel"
      Top             =   4470
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmBatchProductAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdSave_Click()

Dim BP As BatchProduct
Dim BPs As New BatchProducts
Dim EventCode As String

10    On Error GoTo cmdSave_Click_Error

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
100         iMsg "Changes not saved.", vbInformation
110         If TimedOut Then Unload Me: Exit Sub
120         Exit Sub
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

260   BPs.LoadSpecificBatch txtBatchNumber

270   If BPs.Count = 0 Then
280     iMsg "No details found!", vbExclamation
290     If TimedOut Then Unload Me: Exit Sub
300     Exit Sub
310   End If

320   For Each BP In BPs
    
330     EventCode = BP.EventCode
340     BP.DateExpiry = Format(dtExpiry, "dd/mmm/yyyy")
350     BP.UnitVolume = txtVolume
360     BP.UnitGroup = cmbGroup
370     BP.EventCode = "E"
380     BP.Comment = txtReason
390     BPs.Update BP
400     BP.EventCode = EventCode
410     BP.Comment = ""
420     BPs.Update BP
430   Next

440   cmbProduct = ""
450   txtBatchNumber = ""
460   txtVolume = ""
470   cmbGroup = ""
480   txtReason = ""

490   Exit Sub

cmdSave_Click_Error:

Dim strES As String
Dim intEL As Integer

500   intEL = Erl
510   strES = Err.Description
520   LogError "frmBatchProductAmendment", "cmdSave_Click", intEL, strES

End Sub

Private Sub cmbProduct_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub Form_Load()

10    FillcmbProduct

20    With cmbGroup
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
130   LogError "frmBatchProductAmendment", "FillcmbProduct", intEL, strES, sql

End Sub

Private Sub txtBatchNumber_LostFocus()

      Dim BPs As New BatchProducts

10    On Error GoTo txtBatchNumber_LostFocus_Error

20    If Left$(txtBatchNumber, 1) = "A" And Right$(txtBatchNumber, 1) = "B" Then
30      txtBatchNumber = Mid$(txtBatchNumber, 2, 9)
40    End If

50    BPs.LoadSpecificBatch txtBatchNumber

60    If BPs.Count = 0 Then
70      iMsg "No details found!", vbExclamation
80      If TimedOut Then Unload Me: Exit Sub
90      Exit Sub
100   End If

110   If BPs.Item(1).Product <> cmbProduct Then
120     iMsg "Batch number " & txtBatchNumber & " is not " & cmbProduct, vbExclamation
130     If TimedOut Then Unload Me: Exit Sub
140     Exit Sub
150   End If

160   dtExpiry = Format(BPs.Item(1).DateExpiry, "dd/mm/yyyy")
170   lblDateReceived = Format(BPs.Item(1).DateReceived, "dd/mm/yyyy")
180   txtVolume = BPs.Item(1).UnitVolume
190   cmbGroup = BPs.Item(1).UnitGroup

200   Exit Sub

txtBatchNumber_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmBatchProductAmendment", "txtBatchNumber_LostFocus", intEL, strES

End Sub


