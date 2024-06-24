VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMergeClinsWards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5550
   ClientLeft      =   2535
   ClientTop       =   795
   ClientWidth     =   6585
   Icon            =   "frmMergeClinsWards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraReplace 
      Height          =   2325
      Left            =   3450
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txtReplace 
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
         Left            =   270
         TabIndex        =   8
         Top             =   1050
         Width           =   2355
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Yes"
         Height          =   675
         Left            =   900
         Picture         =   "frmMergeClinsWards.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "with"
         Height          =   195
         Left            =   1260
         TabIndex        =   10
         Top             =   840
         Width           =   285
      End
      Begin VB.Label lblReplace 
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
         Height          =   285
         Left            =   270
         TabIndex        =   9
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Replace All occurrances of"
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   210
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   4320
      Picture         =   "frmMergeClinsWards.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4260
      Width           =   1155
   End
   Begin VB.ListBox lstSource 
      Height          =   4740
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   1035
      Left            =   4020
      TabIndex        =   0
      Top             =   150
      Width           =   1995
      Begin VB.OptionButton optSource 
         Caption         =   "Wards"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   2
         Top             =   630
         Width           =   975
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Clinicians"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   540
         TabIndex        =   1
         Top             =   330
         Width           =   1275
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   390
      TabIndex        =   11
      Top             =   5280
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmMergeClinsWards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

      Dim Source As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillList_Error

20    Source = IIf(optSource(0), "Clinician", "Ward")
  
30    lstSource.Clear

40    sql = "Select distinct " & Source & " as Source " & _
            "From PatientDetails " & _
            "where " & Source & " <> '' " & _
            "and " & Source & " is not null " & _
            "Order by " & Source
50    Set tb = New Recordset
60    RecOpenServerBB 0, tb, sql
70    Do While Not tb.EOF
80      lstSource.AddItem tb!Source
90      tb.MoveNext
100   Loop

110   Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmMergeClinsWards", "FillList", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdReplace_Click()

      Dim s As String
      Dim sql As String
      Dim Source As String
      Dim Records As Long

10    On Error GoTo cmdReplace_Click_Error

20    Source = IIf(optSource(0), "Clinician", "Ward")

30    txtReplace = Replace(txtReplace, "'", " ")

40    s = "If you continue, you will replace" & vbCrLf & _
          "all occurrances of" & vbCrLf & _
          lblReplace & " with" & vbCrLf & _
          txtReplace & "." & vbCrLf & _
          "You will not be able to undo this change!" & vbCrLf & _
          "Do you want to proceed?"
50    Answer = iMsg(s, vbQuestion + vbYesNo, , vbRed)
60    If TimedOut Then Unload Me: Exit Sub

70    If Answer = vbYes Then
80      s = iBOX("Enter your Password", , , True)
90      If TimedOut Then Unload Me: Exit Sub
100     If s <> TechnicianPasswordForName(UserName) Then
110       iMsg "Invalid Password." & vbCrLf & "Operation cancelled!", vbInformation
120       If TimedOut Then Unload Me: Exit Sub
130       Exit Sub
140     End If
150     s = "Replace " & lblReplace & " with" & vbCrLf & _
            txtReplace & "." & vbCrLf & _
            "Are you sure?"
160     Answer = iMsg(s, vbQuestion + vbYesNo, , vbRed)
170     If TimedOut Then Unload Me: Exit Sub
180     If Answer = vbYes Then
190       LogReasonWhy "Replaced " & lblReplace & " with " & txtReplace, "XM"
200       sql = "Update PatientDetails " & _
                "Set " & Source & " = '" & txtReplace & "' " & _
                "where " & Source & " = '" & AddTicks(lblReplace) & "'"
210       CnxnBB(0).Execute sql, Records
220       iMsg Format$(Records) & " Record(s) affected."
230       If TimedOut Then Unload Me: Exit Sub
240       FillList
250       fraReplace.Visible = False
260     Else
270       iMsg "Operation cancelled!", vbInformation
280       If TimedOut Then Unload Me: Exit Sub
290     End If
300   Else
310     iMsg "Operation cancelled!", vbInformation
320     If TimedOut Then Unload Me: Exit Sub
330   End If

340   Exit Sub

cmdReplace_Click_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmMergeClinsWards", "cmdReplace_Click", intEL, strES, sql

    
End Sub

Private Sub Form_Load()

10    optSource(0) = False
20    optSource(1) = False

End Sub

Private Sub lstSource_Click()

10    fraReplace.Visible = True
20    lblReplace = lstSource
30    txtReplace = lstSource

End Sub

Private Sub optSource_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10    If Index = 0 Then
20      If optSource(0) Then
30        optSource(0).ForeColor = vbRed
40        optSource(0).Font.Bold = True
50        optSource(1).ForeColor = vbBlack
60        optSource(1).Font.Bold = False
70      End If
80    Else
90      If optSource(1) Then
100       optSource(1).ForeColor = vbRed
110       optSource(1).Font.Bold = True
120       optSource(0).ForeColor = vbBlack
130       optSource(0).Font.Bold = False
140     End If
150   End If

160   FillList

End Sub


