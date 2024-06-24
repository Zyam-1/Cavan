VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmHaemGraphs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Haematology Graphs"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   5220
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   13380
   Begin MSChart20Lib.MSChart gRBC 
      Height          =   2955
      Left            =   330
      OleObjectBlob   =   "frmHaemGraphs.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -60
      Width           =   4275
   End
   Begin MSChart20Lib.MSChart gWBC 
      Height          =   2955
      Left            =   4410
      OleObjectBlob   =   "frmHaemGraphs.frx":18C2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -60
      Width           =   4305
   End
   Begin MSChart20Lib.MSChart gPla 
      Height          =   2955
      Left            =   8640
      OleObjectBlob   =   "frmHaemGraphs.frx":3892
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -60
      Width           =   4365
   End
End
Attribute VB_Name = "frmHaemGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private Sub LoadResults()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim gDataRBC(1 To 64, 1 To 1) As Variant
          Dim gDataWBC(1 To 64, 1 To 3) As Variant
          Dim gDataPLa(1 To 64, 1 To 1) As Variant
          Dim PltVal As Single

5700      On Error GoTo LoadResults_Error

5710      sql = "Select * from HaemResults where " & _
              "SampleID = '" & mSampleID & "'"
5720      Set tb = New Recordset
5730      RecOpenClient 0, tb, sql
5740      If tb.EOF Then
5750          Exit Sub
5760      End If

5770      For n = 1 To 64
5780          gDataRBC(n, 1) = Asc(Mid$(tb!gRBC & String$(64, 1), n, 1))
5790          gDataWBC(n, 1) = Asc(Mid$(tb!gwb1 & String$(64, 1), n, 1))
5800          gDataWBC(n, 2) = Asc(Mid$(tb!gwb2 & String$(64, 1), n, 1))
5810          gDataWBC(n, 3) = Asc(Mid$(tb!gwic & String$(64, 1), n, 1))
5820          gDataPLa(n, 1) = Asc(Mid$(tb!gplt & String$(64, 1), n, 1))
5830      Next

5840      gRBC.ChartData = gDataRBC
5850      gWBC.ChartData = gDataWBC
5860      gPla.ChartData = gDataPLa
5870      PltVal = Val(tb!plt & "")
5880      If PltVal < 100 Then
5890          gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
5900          gPla.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 250
5910      Else
5920          gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
5930      End If

5940      Exit Sub

LoadResults_Error:

          Dim strES As String
          Dim intEL As Integer

5950      intEL = Erl
5960      strES = Err.Description
5970      LogError "fHaemGraphs", "LoadResults", intEL, strES, sql

End Sub
Public Property Let SampleID(ByVal sNewValue As String)

5980      mSampleID = sNewValue

End Property

Private Sub Form_Activate()

5990      LoadResults

End Sub

Private Sub Form_Click()

6000      Unload Me

End Sub


Private Sub gPla_Click()

6010      Unload Me

End Sub

Private Sub gRBC_Click()

6020      Unload Me

End Sub

Private Sub gWBC_Click()

6030      Unload Me

End Sub

