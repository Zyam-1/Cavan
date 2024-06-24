VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManageSensitivities 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Sensitivities"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdCurrent 
      Height          =   1305
      Left            =   900
      TabIndex        =   6
      Top             =   510
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   11340
      Picture         =   "frmManageSensitivities.frx":008A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5100
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdArc 
      Height          =   2625
      Left            =   900
      TabIndex        =   7
      Top             =   1860
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":06F4
   End
   Begin MSFlexGridLib.MSFlexGrid grdRepeat 
      Height          =   1305
      Left            =   900
      TabIndex        =   8
      Top             =   4560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":079F
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   540
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Repeats"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   4620
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   1890
      Width           =   540
   End
End
Attribute VB_Name = "frmManageSensitivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

35680 Unload Me

End Sub


Private Sub Form_Activate()

35690 FillCurrent
35700 FillRepeat
35710 FillArchive

End Sub

Private Sub FillCurrent()

      Dim s As String
      Dim Sxs As New Sensitivities
      Dim sx As Sensitivity

35720 On Error GoTo FillCurrent_Error

35730 With grdCurrent
35740   .Rows = 2
35750   .AddItem ""
35760   .RemoveItem 1

35770   Sxs.Load Val(lblSampleID) ' + sysOptMicroOffset(0)
35780   For Each sx In Sxs
35790     s = sx.IsolateNumber & vbTab & _
              sx.AntibioticCode & vbTab & _
              sx.Result & vbTab & _
              sx.Report & vbTab & _
              sx.CPOFlag & vbTab & _
              Format(sx.Rundate, "dd/MM/yy") & vbTab & _
              Format(sx.RunDateTime, "dd/MM/yy HH:mm") & vbTab & _
              sx.RSI & vbTab & _
              sx.UserCode & vbTab & _
              sx.Forced & vbTab & _
              sx.Secondary & vbTab & _
              sx.Valid & vbTab & _
              sx.AuthoriserCode & ""
35800     .AddItem s
35810   Next
        
35820   If .Rows > 2 Then
35830     .RemoveItem 1
35840   End If

35850 End With

35860 Exit Sub

FillCurrent_Error:

      Dim strES As String
      Dim intEL As Integer

35870 intEL = Erl
35880 strES = Err.Description
35890 LogError "frmManageSensitivities", "FillCurrent", intEL, strES
        
End Sub

Private Sub FillRepeat()

      Dim Sxs As New Sensitivities
      Dim sx As Sensitivity
      Dim s As String

35900 On Error GoTo FillRepeat_Error

35910 With grdRepeat
35920   .Rows = 2
35930   .AddItem ""
35940   .RemoveItem 1

35950   Sxs.LoadRepeats Val(lblSampleID) ' + sysOptMicroOffset(0)
35960   For Each sx In Sxs
35970     s = sx.IsolateNumber & vbTab & _
              sx.AntibioticCode & vbTab & _
              sx.Result & vbTab & _
              sx.Report & vbTab & _
              sx.CPOFlag & vbTab & _
              Format(sx.Rundate, "dd/MM/yy") & vbTab & _
              Format(sx.RunDateTime, "dd/MM/yy HH:mm") & vbTab & _
              sx.RSI & vbTab & _
              sx.UserCode & vbTab & _
              sx.Forced & vbTab & _
              sx.Secondary & vbTab & _
              sx.Valid & vbTab & _
              sx.AuthoriserCode & ""
35980     .AddItem s
35990   Next
        
36000   If .Rows > 2 Then
36010     .RemoveItem 1
36020   End If
36030 End With

36040 Exit Sub

FillRepeat_Error:

      Dim strES As String
      Dim intEL As Integer

36050 intEL = Erl
36060 strES = Err.Description
36070 LogError "frmManageSensitivities", "FillRepeat", intEL, strES
        
End Sub


Private Sub FillArchive()

      Dim Sxs As New Sensitivities
      Dim sx As Sensitivity
      Dim s As String

36080 On Error GoTo FillArchive_Error

36090 With grdArc
36100   .Rows = 2
36110   .AddItem ""
36120   .RemoveItem 1

36130   Sxs.LoadArchive Val(lblSampleID) ' + sysOptMicroOffset(0)
36140   For Each sx In Sxs
36150     s = sx.IsolateNumber & vbTab & _
              sx.AntibioticCode & vbTab & _
              sx.Result & vbTab & _
              sx.Report & vbTab & _
              sx.CPOFlag & vbTab & _
              Format(sx.Rundate, "dd/MM/yy") & vbTab & _
              Format(sx.RunDateTime, "dd/MM/yy HH:mm") & vbTab & _
              sx.RSI & vbTab & _
              sx.UserCode & vbTab & _
              sx.Forced & vbTab & _
              sx.Secondary & vbTab & _
              sx.Valid & vbTab & _
              sx.AuthoriserCode & vbTab & _
              sx.ArchivedBy & vbTab & _
              Format(sx.ArchiveDateTime, "dd/MM/yy HH:mm")
36160     .AddItem s
36170   Next
        
36180   If .Rows > 2 Then
36190     .RemoveItem 1
36200   End If
36210 End With

36220 Exit Sub

FillArchive_Error:

      Dim strES As String
      Dim intEL As Integer

36230 intEL = Erl
36240 strES = Err.Description
36250 LogError "frmManageSensitivities", "FillArchive", intEL, strES
        
End Sub



