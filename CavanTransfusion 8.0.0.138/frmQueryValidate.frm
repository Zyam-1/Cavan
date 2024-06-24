VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmQueryValidate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComment 
      Height          =   765
      Left            =   1440
      TabIndex        =   18
      Top             =   3330
      Width           =   6765
   End
   Begin VB.Frame fraDateTime 
      Caption         =   "Date/Time Transfusion START"
      Height          =   1695
      Left            =   1440
      TabIndex        =   15
      Top             =   4230
      Width           =   3285
      Begin ComCtl2.UpDown udM 
         Height          =   375
         Left            =   2820
         TabIndex        =   26
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtM"
         BuddyDispid     =   196634
         OrigLeft        =   2940
         OrigTop         =   360
         OrigRight       =   3180
         OrigBottom      =   735
         Max             =   59
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtM 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "59"
         Top             =   300
         Width           =   360
      End
      Begin VB.TextBox txtH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "23"
         Top             =   300
         Width           =   360
      End
      Begin ComCtl2.UpDown udH 
         Height          =   375
         Left            =   2145
         TabIndex        =   23
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtH"
         BuddyDispid     =   196633
         OrigLeft        =   2130
         OrigTop         =   300
         OrigRight       =   2370
         OrigBottom      =   645
         Max             =   23
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   345
         Left            =   300
         TabIndex        =   16
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
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
         Format          =   81395713
         CurrentDate     =   41101
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "> 30 minutes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   270
         TabIndex        =   22
         Top             =   750
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblTimeRemovedFromLab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "88:88"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2070
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblDateRemovedFromLab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "88/88/8888"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   900
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblRemovedFromLab 
         Caption         =   "Removed From Lab"
         Height          =   405
         Left            =   150
         TabIndex        =   19
         Top             =   1110
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "No - Cancel"
      Height          =   975
      Left            =   6900
      Picture         =   "frmQueryValidate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes - Continue"
      Height          =   975
      Left            =   5070
      Picture         =   "frmQueryValidate.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4740
      Width           =   1275
   End
   Begin VB.ComboBox cmbReason 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   2550
      Width           =   6795
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   450
      TabIndex        =   14
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   660
   End
   Begin VB.Label lblPatientName 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5100
      TabIndex        =   13
      Top             =   1170
      Width           =   3090
   End
   Begin VB.Label lblPatientNameTitle 
      AutoSize        =   -1  'True
      Caption         =   "Patient Name"
      Height          =   195
      Left            =   4035
      TabIndex        =   12
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label lblChart 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblChartTitle 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label lblCurrentStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5085
      TabIndex        =   9
      Top             =   660
      Width           =   3120
   End
   Begin VB.Label lblUnitNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   630
      Width           =   2385
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Current Status"
      Height          =   195
      Left            =   4005
      TabIndex        =   7
      Top             =   690
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   690
      Width           =   885
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Enter Reason for Return"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   2340
      Width           =   6780
   End
   Begin VB.Label lblNewStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4545
      TabIndex        =   1
      Top             =   1680
      Width           =   3690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This Product will be marked as:-"
      Height          =   195
      Left            =   2205
      TabIndex        =   0
      Top             =   1740
      Width           =   2250
   End
End
Attribute VB_Name = "frmQueryValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Prompt As String
Private m_Options As Collection

Private m_CurrentStatus As String
Private m_NewStatus As String

Private m_UnitNumber As String
Private m_Chart As String
Private m_PatientName As String

Private m_RetVal As Boolean
Private m_Reason As String

Private m_ShowDatePicker As Boolean
Private m_DatePickerCaption As String
Private m_DateTimeReturn As String

Private m_DateTimeRemovedFromLab As String

Public Property Get retval() As Boolean

10    retval = m_RetVal

End Property

Public Property Get DateTimeReturn() As String

10    DateTimeReturn = m_DateTimeReturn

End Property


Public Property Get Comment() As String

10    Comment = Trim$(txtComment)

End Property


Public Property Get Reason() As String

10    Reason = m_Reason

End Property

Public Property Get Options() As Collection

10    Set Options = m_Options

End Property


Public Property Let CurrentStatus(ByVal sNewValue As String)

10    m_CurrentStatus = sNewValue

End Property
Public Property Let Prompt(ByVal sNewValue As String)

10    m_Prompt = sNewValue

End Property

Public Property Let NewStatus(ByVal sNewValue As String)

10    m_NewStatus = sNewValue

End Property

Public Property Let Chart(ByVal sNewValue As String)

10    m_Chart = sNewValue

End Property


Private Sub SetHighLight()

      Dim TxStart As Date
      Dim RFL As Date

10    On Error GoTo SetHighLight_Error

20    lblWarning.Visible = False

30    TxStart = dtDate.Value & " " & txtH & ":" & txtM

40    If DateDiff("n", Now, TxStart) > 0 Then
50      dtDate = Format$(Now, "dd/MM/yyyy")
60      txtH = Format$(Now, "HH")
70      txtM = Format$(Now, "nn")
80      TxStart = dtDate.Value & " " & txtH & ":" & txtM
90    End If
  
100   If Not IsDate(lblDateRemovedFromLab) Or Not IsDate(lblTimeRemovedFromLab) Then Exit Sub
110   RFL = lblDateRemovedFromLab & " " & lblTimeRemovedFromLab
 
120   If lblCurrentStatus = "Removed Pending Transfusion" And _
         lblNewStatus = "Transfused" And _
         IsDate(lblDateRemovedFromLab) And _
         IsDate(lblTimeRemovedFromLab) Then
130     If DateDiff("n", RFL, TxStart) > 30 Then
140       lblWarning.Visible = True
150     ElseIf DateDiff("n", TxStart, RFL) > 0 Then
160       dtDate = Format$(RFL, "dd/MM/yyyy")
170       txtH = Format$(RFL, "HH")
180       txtM = Format$(RFL, "nn")
190     End If
200   End If

210   Exit Sub

SetHighLight_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmQueryValidate", "SetHighLight", intEL, strES

End Sub

Public Property Let ShowDatePicker(ByVal bNewValue As Boolean)

10    m_ShowDatePicker = bNewValue

End Property



Public Property Let PatientName(ByVal sNewValue As String)

10    m_PatientName = sNewValue

End Property

Public Property Let DatePickerCaption(ByVal sNewValue As String)

10    m_DatePickerCaption = sNewValue

End Property


Public Property Let UnitNumber(ByVal sNewValue As String)

10    m_UnitNumber = sNewValue

End Property
Public Property Let Options(ByVal cNewValue As Collection)

10    Set m_Options = cNewValue

End Property

Private Sub cmbReason_Validate(Cancel As Boolean)

10    m_Reason = cmbReason

End Sub


Private Sub cmdCancel_Click()

10    m_Reason = ""
20    m_RetVal = False
30    Me.Hide

End Sub


Private Sub cmdYes_Click()

10    If cmbReason.Visible = True And cmbReason = "" Then
20      iMsg lblPrompt, vbCritical, , vbRed, 10
30      Exit Sub
40    End If

50    m_DateTimeReturn = Format$(dtDate, "dd/MMM/yyyy") & " " & Format$(txtH & ":" & txtM, "HH:mm")
  
60    m_RetVal = True
70    Me.Hide

End Sub


Private Sub dtDate_CloseUp()

10    SetHighLight

End Sub


Private Sub dtDate_LostFocus()

10    SetHighLight

End Sub


Private Sub Form_Activate()

      Dim s As Variant

10    lblPrompt.Visible = False
20    lblChart.Visible = False
30    lblPatientName.Visible = False
40    lblChartTitle.Visible = False
50    lblPatientNameTitle.Visible = False
60    cmbReason.Visible = False
70    fraDateTime.Visible = False

80    If m_Prompt <> "" Then
90      lblPrompt.Caption = m_Prompt
100     cmbReason.Clear
110     For Each s In m_Options
120       cmbReason.AddItem s
130     Next
140     lblPrompt.Visible = True
150     cmbReason.Visible = True
160   End If

170   lblCurrentStatus = m_CurrentStatus
180   lblNewStatus = m_NewStatus
190   lblUnitNumber = m_UnitNumber
200   If m_Chart <> "" Then
210     lblChart = m_Chart
220     lblChart.Visible = True
230     lblChartTitle.Visible = True
240   End If
250   If m_PatientName <> "" Then
260     lblPatientName = m_PatientName
270     lblPatientName.Visible = True
280     lblPatientNameTitle.Visible = True
290   End If

300   If m_ShowDatePicker = True Then
310     If lblCurrentStatus = "Removed Pending Transfusion" And _
           lblNewStatus = "Transfused" And _
           IsDate(lblDateRemovedFromLab) And _
           IsDate(lblTimeRemovedFromLab) Then
320       dtDate = lblDateRemovedFromLab
330       txtH = Format$(lblTimeRemovedFromLab, "HH")
340       txtM = Format$(lblTimeRemovedFromLab, "nn")
350     Else
360       dtDate = Format$(Now, "dd/MM/yyyy")
370       txtH = Format$(Now, "HH")
380       txtM = Format$(Now, "nn")
390     End If
400     fraDateTime.Caption = m_DatePickerCaption
410     fraDateTime.Visible = True
420   End If

430   SetHighLight

End Sub

Public Property Get DateTimeRemovedFromLab() As String

10      DateTimeRemovedFromLab = m_DateTimeRemovedFromLab

End Property

Public Property Let DateTimeRemovedFromLab(ByVal DateTimeRemovedFromLab As String)

10    m_DateTimeRemovedFromLab = DateTimeRemovedFromLab

20    If IsDate(m_DateTimeRemovedFromLab) And m_CurrentStatus = "Removed Pending Transfusion" Then
30      lblDateRemovedFromLab = Format(m_DateTimeRemovedFromLab, "dd/MM/yyyy")
40      lblTimeRemovedFromLab = Format(m_DateTimeRemovedFromLab, "HH:nn")
50      lblDateRemovedFromLab.Visible = True
60      lblTimeRemovedFromLab.Visible = True
70      lblRemovedFromLab.Visible = True
80    End If

End Property

Private Sub txtH_KeyPress(KeyAscii As Integer)

10    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
20      KeyAscii = 0
30    End If

End Sub


Private Sub txtH_LostFocus()

10    SetHighLight

End Sub


Private Sub txtM_KeyPress(KeyAscii As Integer)

10    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
20      KeyAscii = 0
30    End If

End Sub


Private Sub txtM_LostFocus()

10    SetHighLight

End Sub


Private Sub udH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    SetHighLight

End Sub


Private Sub udM_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    SetHighLight

End Sub


