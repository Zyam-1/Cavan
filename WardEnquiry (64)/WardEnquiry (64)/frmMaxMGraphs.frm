VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frmMaxMGraphs 
   Caption         =   "NetAcquire"
   ClientHeight    =   5520
   ClientLeft      =   3900
   ClientTop       =   2910
   ClientWidth     =   7605
   Icon            =   "frmMaxMGraphs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7605
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   2
      Left            =   4050
      TabIndex        =   6
      Text            =   "Plt"
      Top             =   3150
      Width           =   495
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Text            =   "WBC"
      Top             =   3150
      Width           =   555
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Text            =   "RBC"
      Top             =   210
      Width           =   525
   End
   Begin VB.PictureBox gDF1 
      AutoRedraw      =   -1  'True
      Height          =   2550
      Left            =   4710
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   3
      Top             =   180
      Width           =   2550
   End
   Begin MSChart20Lib.MSChart gPla 
      Height          =   2505
      Left            =   3780
      OleObjectBlob   =   "frmMaxMGraphs.frx":0ECA
      TabIndex        =   0
      Top             =   2880
      Width           =   3765
   End
   Begin MSChart20Lib.MSChart gWBC 
      Height          =   2505
      Left            =   -120
      OleObjectBlob   =   "frmMaxMGraphs.frx":278C
      TabIndex        =   1
      Top             =   2880
      Width           =   4065
   End
   Begin MSChart20Lib.MSChart gRBC 
      Height          =   3045
      Left            =   -150
      OleObjectBlob   =   "frmMaxMGraphs.frx":475C
      TabIndex        =   2
      Top             =   -60
      Width           =   5025
   End
End
Attribute VB_Name = "frmMaxMGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mSampleID As String
Private mCN As Integer

Public Property Let SampleID(ByVal sNewValue As String)

10    mSampleID = sNewValue

End Property


Public Property Let UseCnxn(ByVal Cn As Integer)

10    mCN = Cn

End Property



Private Sub DrawGraphs()

      Dim tb As Recordset
      Dim sql As String

      Dim gDataRBC(1 To 256, 1 To 1) As Variant
      Dim gDataWBC(1 To 256, 1 To 3) As Variant
      Dim gDataPLa(1 To 256, 1 To 2) As Variant
      Dim PLTH As String
      Dim PLTF As String
      Dim gV As String
      Dim gC As String
      Dim gS As String
      Dim RBCH As String
      Dim DF1 As String
      Dim VL(1 To 5) As Integer
      Dim x As Integer
      Dim y As Integer
      Dim n As Integer
      Dim P As Integer
      Dim strPla As String

10    On Error GoTo DrawGraphs_Error

20    sql = "Select * from HaemResults where " & _
            "SampleID = '" & mSampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient mCN, tb, sql
50    If tb.EOF Then
60      Exit Sub
70    End If
80    PLTH = "00000000000000000000000000" & Left$(tb!gplth & String$(512, "0"), 512)
90    PLTF = Left$(tb!gPLTF & String$(512, "0"), 512)
100   gV = Left$("0" & tb!gV & String$(512, "0"), 512)
110   gC = Left$("0" & tb!gC & String$(512, "0"), 512)
120   gS = Left$("0" & tb!gS & String$(512, "0"), 512)
130   RBCH = Left$(tb!gRBCH & String$(512, "0"), 512)
140   DF1 = Left$(tb!DF1 & String$(4096, "0"), 4096)

150   For n = 1 To 5
160     VL(n) = IIf(IsNull(tb("Val" & Format(n))), 0, tb("Val" & Format(n))) / 4
170   Next

180   gDF1.ScaleMode = vbUser
190   gDF1.ScaleHeight = 63
200   gDF1.ScaleWidth = 63
210   gDF1.Cls
220   gDF1.DrawWidth = 3
230   For y = 0 To 63
240     For x = 0 To 63
250       P = Val("&h" & Mid$(DF1, (y + 1) + 64 * x, 1))
260       Select Case P
            Case 1:        gDF1.PSet (x, 63 - y), vbBlue
270         Case 2 To 4:   gDF1.PSet (x, 63 - y), vbGreen
280         Case 5 To 10:  gDF1.PSet (x, 63 - y), vbRed
290         Case 11 To 15: gDF1.PSet (x, 63 - y), vbYellow
300       End Select
310     Next
320   Next

330   gDF1.DrawWidth = 1
340   gDF1.Line (0, 63 - VL(3))-(VL(1), 63 - VL(3))
350   gDF1.Line (VL(1), 63 - VL(4))-(63, 63 - VL(4))
360   gDF1.Line (0, 63 - VL(5))-(63, 63 - VL(5))
370   gDF1.Line (VL(1), 0)-(VL(1), 63 - VL(5))
380   gDF1.Line (VL(2), 0)-(VL(2), 63 - VL(4))

390   For n = 1 To 512 Step 2
400     gDataPLa(Int(n / 2) + 1, 1) = Val("&h" & Mid$(PLTH, n, 2))
410     gDataPLa(Int(n / 2) + 1, 2) = Val("&h" & Mid$(PLTF, n, 2))
420     gDataWBC(Int(n / 2) + 1, 1) = Val("&h" & Mid$(gV, n, 2))
430     gDataWBC(Int(n / 2) + 1, 2) = Val("&h" & Mid$(gC, n, 2))
440     gDataWBC(Int(n / 2) + 1, 3) = Val("&h" & Mid$(gS, n, 2))
450     gDataRBC(Int(n / 2) + 1, 1) = Val("&h" & Mid$(RBCH, n, 2))
460   Next

470   gRBC.ChartData = gDataRBC
480   gWBC.ChartData = gDataWBC
490   gPla.ChartData = gDataPLa
500   strPla = Format(tb!plt & "", "####")

510   gPla.Plot.SeriesCollection(1).Pen.Width = 1
520   gPla.Plot.SeriesCollection(2).Pen.Width = 1
530   gWBC.Plot.SeriesCollection(1).Pen.Width = 1
540   gWBC.Plot.SeriesCollection(2).Pen.Width = 1
550   gWBC.Plot.SeriesCollection(3).Pen.Width = 1
560   gRBC.Plot.SeriesCollection(1).Pen.Width = 1

570   If Val(strPla) < 100 Then
580     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
590     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 250
600   Else
610     gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
620   End If

630   Exit Sub

DrawGraphs_Error:

      Dim strES As String
      Dim intEL As Integer

640   intEL = Erl
650   strES = Err.Description
660   LogError "frmMaxMGraphs", "DrawGraphs", intEL, strES, sql


End Sub

Private Sub Form_Activate()

10    DrawGraphs
20    SingleUserUpdateLoggedOn UserName

End Sub

Private Sub Form_Click()

10    Unload Me

End Sub


Private Sub gDF1_Click()

10    Unload Me

End Sub

Private Sub gPla_Click()

10    Unload Me

End Sub

Private Sub gRBC_Click()

10    Unload Me

End Sub

Private Sub gWBC_Click()

10    Unload Me

End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

10    Unload Me

End Sub


