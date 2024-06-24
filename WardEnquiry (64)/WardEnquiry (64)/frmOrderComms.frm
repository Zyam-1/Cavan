VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrderComms 
   Caption         =   "NetAcquire"
   ClientHeight    =   5985
   ClientLeft      =   915
   ClientTop       =   1095
   ClientWidth     =   10440
   Icon            =   "frmOrderComms.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10440
   Begin VB.Frame fraDisplay 
      Caption         =   "Display Panels"
      Height          =   495
      Left            =   6990
      TabIndex        =   20
      Top             =   1530
      Width           =   1545
      Begin VB.OptionButton optShortName 
         Alignment       =   1  'Right Justify
         Caption         =   "Short"
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optLongName 
         Caption         =   "Long"
         Height          =   195
         Left            =   780
         TabIndex        =   21
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.OptionButton optOrder 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Tests to be Ordered"
      Height          =   225
      Left            =   330
      TabIndex        =   19
      Top             =   1800
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.OptionButton optServiceSet 
      Caption         =   "by Service Set"
      Height          =   225
      Left            =   2610
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   1110
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Request Lab Services"
      Height          =   1275
      Left            =   8940
      Picture         =   "frmOrderComms.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstOrder 
      Columns         =   5
      Height          =   3570
      Left            =   330
      MultiSelect     =   1  'Simple
      TabIndex        =   15
      Top             =   2040
      Width           =   8205
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   300
      TabIndex        =   1
      Top             =   0
      Width           =   9825
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lblSampleID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1035
         TabIndex        =   23
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label lblName 
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
         Left            =   2775
         TabIndex        =   14
         Top             =   210
         Width           =   3540
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1035
         TabIndex        =   13
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7275
         TabIndex        =   12
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4515
         TabIndex        =   11
         Top             =   510
         Width           =   5100
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   2325
         TabIndex        =   10
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblChartTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Chart"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   3915
         TabIndex        =   7
         Top             =   540
         Width           =   570
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8895
         TabIndex        =   6
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2775
         TabIndex        =   5
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   8580
         TabIndex        =   4
         Top             =   210
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2445
         TabIndex        =   3
         Top             =   540
         Width           =   270
      End
      Begin VB.Label lblDemogComment 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   840
         Width           =   9465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1275
      Left            =   8940
      Picture         =   "frmOrderComms.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4350
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   300
      TabIndex        =   17
      Top             =   1230
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmOrderComms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillOrderList(ByVal ByServiceSet As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean
      Dim LongOrShort As String

10    On Error GoTo FillOrderList_Error

20    lstOrder.Clear

30    LongOrShort = IIf(optLongName, "LongName", "ShortName")

40    sql = "Select " & LongOrShort & " from ocOrderPanel " & _
            "where ClinicianOrWard "
50    If ByServiceSet Then
60      sql = sql & "= '" & UserName & "'"
70    Else
80      sql = sql & "= '' or ClinicianOrWard is null "
90    End If
100   sql = sql & " order by ListOrder"
110   Set tb = New Recordset
120   RecOpenServer 0, tb, sql
130   Do While Not tb.EOF
140     Found = False
150     For n = 0 To lstOrder.ListCount - 1
160       If lstOrder.List(n) = tb(LongOrShort) Then
170         Found = True
180         Exit For
190       End If
200     Next
210     If Not Found Then
220       lstOrder.AddItem tb(LongOrShort) & ""
230     End If
240     tb.MoveNext
250   Loop
  
260   If ByServiceSet And lstOrder.ListCount = 0 Then
270     iMsg "No Service Set found." & vbCrLf & _
             "If you wish to be assigned a Service Set," & vbCrLf & _
             "Please contact your Laboratory Manager.", vbInformation
280   End If

290   Exit Sub

FillOrderList_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmOrderComms", "FillOrderList", intEL, strES, sql

End Sub

Private Sub FillServiceSetOrder()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean

10    lstOrder.Clear

20    sql = "Select * from ServiceSetOrder " & _
            "where ClinicianOrWard = '" & UserName & "' " & _
            "order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60      Do While Not tb.EOF
70        Found = False
80        For n = 0 To lstOrder.ListCount - 1
90          If UCase$(tb!OrderPanel & "") = UCase$(lstOrder.List(n)) Then
100           Found = True
110           Exit For
120         End If
130       Next
140       If Not Found Then
150         lstOrder.AddItem tb!OrderPanel & ""
160       End If
170       tb.MoveNext
180     Loop
190   Else
  
200     iMsg "No Service Set found." & vbCrLf & _
             "If you wish to be assigned a Service Set," & vbCrLf & _
             "Please contact your Laboratory Manager.", vbInformation

210   End If

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub Form_Load()

10    PBar.Max = LogOffDelaySecs
20    PBar = 0

30      lblChartTitle = "Chart"

40    FillOrderList False

End Sub


Private Sub optLongName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    FillOrderList optServiceSet = True

End Sub


Private Sub optOrder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    FillOrderList 0

End Sub


Private Sub optServiceSet_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

10    FillOrderList True

End Sub


Private Sub optShortName_Click()

10    FillOrderList optServiceSet = True

End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10    PBar = PBar + 1
  
20    If PBar = PBar.Max Then
30      LogOffNow = True
40      Unload Me
50    End If

End Sub


