VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewMedibridge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Reports"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   600
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   9300
      Picture         =   "frmViewMedibridge.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4980
      Width           =   1605
   End
   Begin RichTextLib.RichTextBox rtbResult 
      Height          =   4005
      Left            =   60
      TabIndex        =   10
      Top             =   870
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7064
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmViewMedibridge.frx":066A
   End
   Begin VB.Label lblSampleID 
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
      Left            =   1230
      TabIndex        =   9
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   510
      TabIndex        =   8
      Top             =   180
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Report Time"
      Height          =   195
      Left            =   7440
      TabIndex        =   7
      Top             =   180
      Width           =   870
   End
   Begin VB.Label lblMessageTime 
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
      Left            =   8370
      TabIndex        =   6
      Top             =   150
      Width           =   2415
   End
   Begin VB.Label lblSex 
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
      Left            =   5910
      TabIndex        =   5
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label lblDoB 
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
      Left            =   3420
      TabIndex        =   4
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label lblPatName 
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
      Left            =   3420
      TabIndex        =   3
      Top             =   150
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   5610
      TabIndex        =   2
      Top             =   540
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   3030
      TabIndex        =   1
      Top             =   510
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2970
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmViewMedibridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private Sub cmdCancel_Click()

33940 Unload Me

End Sub

Private Sub Form_Activate()

      Dim tb As Recordset
      Dim sql As String

33950 On Error GoTo Form_Activate_Error

33960 sql = "Select * from Demographics where " & _
            "SampleID = '" & pSampleID & "'"
33970 Set tb = New Recordset
33980 RecOpenServer 0, tb, sql
33990 If Not tb.EOF Then
34000   lblSampleID = Val(tb!SampleID)
34010   lblPatName = tb!PatName & ""
34020   lblDoB = tb!DoB & ""
34030   Select Case UCase$(Left$(Trim$(tb!Sex & ""), 1))
          Case "M": lblSex = "Male"
34040     Case "F": lblSex = "Female"
34050     Case "U": lblSex = "Unknown"
34060     Case Else: lblSex = "Not Given"
34070   End Select
34080 End If

34090 rtbResult.SelText = ""

34100 sql = "Select * from MedibridgeResults where " & _
            "SampleID = '" & pSampleID & "'"
34110 Set tb = New Recordset
34120 RecOpenServer 0, tb, sql
34130 Do While Not tb.EOF
34140   lblMessageTime = tb!MessageTime
          
34150   With rtbResult
34160     .SelIndent = 0
34170     .SelColor = vbBlue
34180     .SelBold = False
34190     .SelText = "Request: "
34200     .SelBold = False
34210     .SelText = .SelText & tb!Request & vbCrLf
34220     .SelColor = vbBlack
34230     .SelBold = True
34240     .SelIndent = 200
34250     .SelText = .SelText & tb!Result & ""
34260   End With

34270   tb.MoveNext
34280 Loop

34290 sql = "Select * from ExtResults where " & _
            "SampleID = '" & pSampleID & "'"
34300 Set tb = New Recordset
34310 RecOpenServer 0, tb, sql
34320 Do While Not tb.EOF

34330   With rtbResult
34340     .SelIndent = 0
34350     .SelColor = vbBlue
          '.SelBold = False
          '.SelText = "Analyte: "
34360     .SelBold = True
34370     .SelText = .SelText & tb!Analyte & ": "
34380     .SelColor = vbBlack
34390     .SelBold = True
34400     .SelIndent = 200
34410     If Trim$(tb!Result & "") <> "" Then
34420       .SelText = .SelText & tb!Result & " " & tb!Units & ""
34430     Else
34440       .SelText = .SelText & "Not yet Available."
34450     End If
34460     .SelText = .SelText & vbCrLf
34470   End With

34480   tb.MoveNext
34490 Loop

34500 Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

34510 intEL = Erl
34520 strES = Err.Description
34530 LogError "frmViewMedibridge", "Form_Activate", intEL, strES, sql


End Sub


Public Property Let SampleID(ByRef NewValue As String)

34540 If HospName(0) = "Monaghan" Then
34550   pSampleID = Val(NewValue) ' + sysOptMicroOffset(0)
34560 Else
34570   pSampleID = Val(NewValue)
34580 End If

End Property
