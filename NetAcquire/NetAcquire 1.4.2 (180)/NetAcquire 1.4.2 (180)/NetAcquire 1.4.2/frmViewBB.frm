VERSION 5.00
Begin VB.Form frmViewBB 
   Caption         =   "NetAcquire"
   ClientHeight    =   4080
   ClientLeft      =   1725
   ClientTop       =   1410
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6210
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   4470
      TabIndex        =   1
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label lGroup 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1860
      TabIndex        =   18
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label lAnti3Reported 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   17
      Top             =   1500
      Width           =   1770
   End
   Begin VB.Label lAIDr 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   16
      Top             =   1890
      Width           =   1770
   End
   Begin VB.Label lProcedure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   15
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lConditions 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   14
      Top             =   2700
      Width           =   3855
   End
   Begin VB.Label lComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lSampleComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   12
      Top             =   3540
      Width           =   3855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Antibody ID"
      Height          =   195
      Left            =   900
      TabIndex        =   11
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Antibodies Reported"
      Height          =   195
      Left            =   285
      TabIndex        =   10
      Top             =   1530
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Conditions"
      Height          =   195
      Left            =   990
      TabIndex        =   9
      Top             =   2730
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      Height          =   195
      Left            =   990
      TabIndex        =   8
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample Comment"
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   3570
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   1065
      TabIndex        =   6
      Top             =   3150
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   1290
      TabIndex        =   5
      Top             =   930
      Width           =   435
   End
   Begin VB.Label lName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   4
      Top             =   510
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   1410
      TabIndex        =   3
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "frmViewBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

32050 Unload Me

End Sub


Private Sub Form_Activate()

      Dim sql As String
      Dim tb As Recordset

32060 On Error GoTo Form_Activate_Error

32070 sql = "Select * from PatientDetails where " & _
            "PatNum = '" & lchart & "'"
32080 Set tb = New Recordset
32090 RecOpenClientBB 0, tb, sql
32100 If tb.EOF Then
32110   lname = "No Record in Blood Bank"
32120 Else
32130   lname = tb!Name & ""
32140   lProcedure = tb!Procedure & ""
32150   lConditions = tb!conditions & ""
32160   lGroup = tb!fgroup & ""
32170   lAnti3Reported = tb!anti3reported & ""
32180   lcomment = tb!Comment & ""
32190   lAIDr = tb!aidr & ""
32200   lSampleComment = tb!samplecomment & ""
32210 End If

32220 Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

32230 intEL = Erl
32240 strES = Err.Description
32250 LogError "fViewBB", "Form_Activate", intEL, strES, sql


End Sub

