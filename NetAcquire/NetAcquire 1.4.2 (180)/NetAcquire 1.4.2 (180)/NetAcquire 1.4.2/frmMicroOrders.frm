VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMicroOrders 
   Caption         =   "NetAcquire"
   ClientHeight    =   5460
   ClientLeft      =   3150
   ClientTop       =   975
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6840
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   3540
      MaxLength       =   6
      TabIndex        =   28
      Top             =   870
      Width           =   1545
   End
   Begin VB.TextBox txtDoB 
      Height          =   285
      Left            =   3540
      MaxLength       =   10
      TabIndex        =   27
      Top             =   510
      Width           =   1545
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   870
      MaxLength       =   30
      TabIndex        =   26
      Tag             =   "tName"
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtChart 
      Height          =   300
      Left            =   870
      MaxLength       =   8
      TabIndex        =   25
      Top             =   870
      Width           =   1545
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   870
      MaxLength       =   12
      TabIndex        =   22
      Top             =   510
      Width           =   1545
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   5310
      Picture         =   "frmMicroOrders.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1860
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   5310
      Picture         =   "frmMicroOrders.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Requests"
      Height          =   2115
      Left            =   3180
      TabIndex        =   13
      Top             =   1740
      Width           =   1905
      Begin VB.CheckBox chkUrine 
         Caption         =   "Urinary HCG"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1710
         Width           =   1215
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Specific Gravity"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1470
         Width           =   1455
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Fat Globules"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Bence Jones"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1230
         Width           =   1245
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Pregnancy"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   570
         Width           =   1155
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "C && S"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Faecal Requests"
      Height          =   3435
      Left            =   870
      TabIndex        =   0
      Top             =   1740
      Width           =   2145
      Begin VB.CheckBox chkFaecal 
         Caption         =   "S/S Screen"
         Height          =   255
         Index           =   11
         Left            =   180
         TabIndex        =   12
         Top             =   3090
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   11
         Top             =   1200
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check10"
         Height          =   195
         Index           =   4
         Left            =   420
         TabIndex        =   10
         Top             =   1200
         Width           =   225
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check9"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Toxin A"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   8
         Top             =   2040
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C.Difficile"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   630
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Coli 0157"
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   6
         Top             =   2310
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Adeno Virus"
         Height          =   255
         Index           =   10
         Left            =   180
         TabIndex        =   5
         Top             =   2820
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Rota Virus"
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   4
         Top             =   2550
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "O/P"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Width           =   645
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   825
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "E/P Coli"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   1
         Top             =   1770
         Width           =   915
      End
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   210
      Left            =   1860
      TabIndex        =   23
      Top             =   300
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   370
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtSampleID"
      BuddyDispid     =   196613
      OrigLeft        =   1920
      OrigTop         =   540
      OrigRight       =   2160
      OrigBottom      =   1020
      Max             =   99999
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   450
      TabIndex        =   32
      Top             =   930
      Width           =   375
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   3240
      TabIndex        =   31
      Top             =   900
      Width           =   270
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "D.o.B"
      Height          =   195
      Index           =   0
      Left            =   3090
      TabIndex        =   30
      Top             =   540
      Width           =   405
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   29
      Top             =   1230
      Width           =   420
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   930
      TabIndex        =   24
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Long
      Dim n As Integer

41980 On Error GoTo LoadDetails_Error

      '20    SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)
41990 SampleIDWithOffset = Val(txtSampleID)

42000 For n = 0 To 11
42010   chkFaecal(n) = 0
42020 Next

42030 For n = 0 To 5
42040   chkUrine(n) = 0
42050 Next
      '+++ Junaid 20-05-2024
      '90    sql = "Select Patname, DoB, Chart, Sex " & _
      '            "from Demographics where " & _
      '            "SampleID = '" & SampleIDWithOffset & "' "
42060 sql = "Select Patname, DoB, Chart, Sex " & _
            "from Demographics where " & _
            "SampleID = '" & Trim(txtSampleID.Text) & "' "
      '--- Junaid
42070 Set tb = New Recordset
42080 RecOpenServer 0, tb, sql
42090 If Not tb.EOF Then
42100   txtDoB = tb!DoB & ""
42110   txtChart = tb!Chart & ""
42120   Select Case UCase$(Left$(tb!Sex & "", 1))
          Case "M": txtSex = "Male"
42130     Case "F": txtSex = "Female"
42140     Case Else: txtSex = ""
42150   End Select
42160   txtName = tb!PatName & ""
42170 Else
42180   txtDoB = ""
42190   txtChart = ""
42200   txtSex = ""
42210   txtName = ""
42220 End If
      '+++ Junaid 20-05-2024
      '260   sql = "Select Faecal, Urine " & _
      '            "from MicroRequests where " & _
      '            "SampleID = '" & SampleIDWithOffset & "' "
42230 sql = "Select Faecal, Urine " & _
            "from MicroRequests where " & _
            "SampleID = '" & Trim(txtSampleID.Text) & "' "
      '--- Junaid
42240 Set tb = New Recordset
42250 RecOpenServer 0, tb, sql
42260 If Not tb.EOF Then

42270   For n = 0 To 11
42280     If tb!Faecal And 2 ^ n Then
42290       chkFaecal(n) = 1
42300     End If
42310   Next
        
42320   For n = 0 To 7
42330     If tb!Urine And 2 ^ n Then
42340       chkUrine(n) = 1
42350     End If
42360   Next

42370 End If
        
42380 cmdSave.Enabled = False

42390 Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

42400 intEL = Erl
42410 strES = Err.Description
42420 LogError "frmMicroOrders", "LoadDetails", intEL, strES, sql


End Sub

Private Sub SaveDetails()

      Dim lngF As Long
      Dim lngU As Long
      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Long

42430 On Error GoTo SaveDetails_Error

42440 lngF = 0
42450 For n = 0 To 11
42460   If chkFaecal(n) Then
42470     lngF = lngF + 2 ^ n
42480   End If
42490 Next

42500 lngU = 0
42510 For n = 0 To 7
42520   If chkUrine(n) Then
42530     lngU = lngU + 2 ^ n
42540   End If
42550 Next

42560 If lngU + lngF = 0 Then
42570   iMsg "Nothing to Save!", vbExclamation
42580   cmdSave.Enabled = False
42590   Exit Sub
42600 End If

      '190   SampleIDWithOffset = Val(txtSampleID) + sysOptMicroOffset(0)
42610 SampleIDWithOffset = Val(txtSampleID)
      '+++ Junaid 20-05-2024
      '200   sql = "Select * from MicroRequests where " & _
      '            "SampleID = '" & SampleIDWithOffset & "'"
42620 sql = "Select * from MicroRequests where " & _
            "SampleID = '" & Trim(txtSampleID.Text) & "'"
      '--- Junaid
42630 Set tb = New Recordset
42640 RecOpenServer 0, tb, sql
42650 If tb.EOF Then
42660   tb.AddNew
42670 End If
      '+++ Junaid 20-05-2024
      '260   tb!SampleID = SampleIDWithOffset
42680 tb!SampleID = Trim(txtSampleID.Text)
      '--- Junaid
42690 tb!RequestDate = Format(Now, "dd/mmm/yyyy hh:mm")
42700 tb!Faecal = lngF
42710 tb!Urine = lngU
42720 tb.Update
        
42730 cmdSave.Enabled = False

42740 Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

42750 intEL = Erl
42760 strES = Err.Description
42770 LogError "frmMicroOrders", "SaveDetails", intEL, strES, sql


End Sub

Private Sub chkFaecal_Click(Index As Integer)

42780 cmdSave.Enabled = True

End Sub

Private Sub chkUrine_Click(Index As Integer)

42790 If Index = 6 Then
42800   chkUrine(7) = 0
42810 ElseIf Index = 7 Then
42820   chkUrine(6) = 0
42830 End If

42840 cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

42850 If cmdSave.Enabled Then
42860   If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
42870     Exit Sub
42880   End If
42890 End If

42900 Unload Me

End Sub


Private Sub cmdSave_Click()

42910 SaveDetails

End Sub

Private Sub Form_Load()

42920 LoadDetails

End Sub

Private Sub txtsampleid_LostFocus()

42930 LoadDetails

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

42940 LoadDetails

End Sub

