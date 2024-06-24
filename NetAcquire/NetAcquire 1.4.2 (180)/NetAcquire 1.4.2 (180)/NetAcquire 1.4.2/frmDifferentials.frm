VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDifferentials 
   Caption         =   "NetAcquire - Differentials"
   ClientHeight    =   7350
   ClientLeft      =   645
   ClientTop       =   1380
   ClientWidth     =   5400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   5400
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5235
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   9234
      _Version        =   393216
      Rows            =   21
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   0
      FormatString    =   "^Key |<Cells                            |^Count % |^Count # "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancel without saving"
      Height          =   525
      Left            =   390
      TabIndex        =   3
      Top             =   6240
      Width           =   1785
   End
   Begin VB.CommandButton bSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   3600
      TabIndex        =   2
      Top             =   6240
      Width           =   1245
   End
   Begin VB.Label lWBC 
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
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   330
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "WBC"
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   390
      Width           =   375
   End
End
Attribute VB_Name = "frmDifferentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private total As Integer

Private Sub bcancel_Click()

34620     Unload Me

End Sub

Private Sub bSave_Click()

          Dim n As Integer

34630     With frmEditAll
34640         .tNeutP = g.TextMatrix(1, 2)
34650         .tNeutA = g.TextMatrix(1, 3)
34660         .tLymP = g.TextMatrix(2, 2)
34670         .tLymA = g.TextMatrix(2, 3)
34680         .tMonoP = g.TextMatrix(3, 2)
34690         .tMonoA = g.TextMatrix(3, 3)
34700         .tEosP = g.TextMatrix(4, 2)
34710         .tEosA = g.TextMatrix(4, 3)
34720         .tBasP = g.TextMatrix(5, 2)
34730         .tBasA = g.TextMatrix(5, 3)
34740         .tnrbcA = g.TextMatrix(6, 2)
34750         .tnrbcP = g.TextMatrix(6, 3)
          
34760         For n = 7 To 20
34770             If Trim$(g.TextMatrix(n, 2)) <> "" Then
34780                 .txtHaemComment = .txtHaemComment & g.TextMatrix(n, 1) & ":" & g.TextMatrix(n, 3) & " "
34790             End If
34800         Next
34810     End With

34820     Unload Me

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

          Dim n As Integer

34830     If total = 100 Then
34840         KeyAscii = 0
34850         Beep
34860         Exit Sub
34870     End If

34880     For n = 1 To 20
34890         If UCase$(Chr(KeyAscii)) = g.TextMatrix(n, 0) Then
34900             If n <> 6 Then
34910                 g.TextMatrix(n, 2) = Val(g.TextMatrix(n, 2)) + 1
34920                 g.TextMatrix(n, 3) = Format$(Val(g.TextMatrix(n, 2)) * Val(lWBC) / 100, "0.0")
34930             Else 'NRBC's
34940                 g.TextMatrix(n, 3) = Val(g.TextMatrix(n, 3)) + 1
34950             End If
34960             Exit For
34970         End If
34980     Next
34990     KeyAscii = 0

35000     total = 0
35010     For n = 1 To 20
35020         If n <> 6 Then
35030             total = total + Val(g.TextMatrix(n, 2))
35040         End If
35050     Next
35060     Debug.Print total
35070     If total >= 100 Then
35080         Beep
35090     End If

End Sub

Private Sub Form_Load()

          Dim n As Integer

35100     For n = 0 To 19
35110         g.TextMatrix(n + 1, 0) = GetSetting("NetAcquire", "DiffKey", Format$(n))
35120     Next

35130     For n = 0 To 5
35140         g.TextMatrix(n + 1, 1) = Choose(n + 1, "Neutrophils", "Lymphocytes", _
                  "Monocytes", "Eosinophils", _
                  "Basophils", "NRBC's")
35150     Next

35160     For n = 6 To 19
35170         g.TextMatrix(n + 1, 1) = GetSetting("NetAcquire", "DiffCell", Format$(n))
35180     Next

35190     total = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

35200     total = 0

End Sub


Private Sub g_Click()

          Dim n As Integer

35210     If g.MouseRow = 0 Then Exit Sub

35220     Select Case g.Col
              Case 0: 'key
35230             g = UCase$(iBOX("New Key?", , g))
35240         Case 1: 'Wording
35250             If g.row < 7 Then Exit Sub
35260             g = iBOX("New Cell?", , g)
35270     End Select

35280     For n = 0 To 19
35290         SaveSetting "NetAcquire", "DiffKey", Format$(n), g.TextMatrix(n + 1, 0)
35300     Next
35310     For n = 6 To 19
35320         SaveSetting "NetAcquire", "DiffCell", Format$(n), UCase$(g.TextMatrix(n + 1, 1))
35330     Next

End Sub

