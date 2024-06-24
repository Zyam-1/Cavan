VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhoneLogHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Transfusion Phone Log History"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   Icon            =   "frmPhoneLogHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBoth 
      Caption         =   "Both"
      Height          =   255
      Left            =   570
      TabIndex        =   8
      Top             =   720
      Width           =   1905
   End
   Begin VB.OptionButton optCallOut 
      Caption         =   "Dialed Calls"
      Height          =   255
      Left            =   570
      TabIndex        =   7
      Top             =   435
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.OptionButton optCallIn 
      Caption         =   "Received Calls"
      Height          =   255
      Left            =   570
      TabIndex        =   6
      Top             =   150
      Width           =   1905
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   11520
      Picture         =   "frmPhoneLogHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4830
      Width           =   1485
   End
   Begin VB.TextBox txtSampleID 
      Height          =   315
      Left            =   2580
      TabIndex        =   1
      Top             =   600
      Width           =   1680
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdPhoneLog 
      Height          =   3585
      Left            =   150
      TabIndex        =   3
      Top             =   1170
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   11
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmPhoneLogHistory.frx":0F34
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   5
      Top             =   5580
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image imgCallIn 
      Height          =   240
      Left            =   9060
      Picture         =   "frmPhoneLogHistory.frx":0FFF
      Top             =   390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCallOut 
      Height          =   240
      Left            =   7800
      Picture         =   "frmPhoneLogHistory.frx":1374
      Top             =   510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   7470
      Picture         =   "frmPhoneLogHistory.frx":16EE
      Top             =   90
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phone Log History for Sample ID"
      Height          =   390
      Left            =   2580
      TabIndex        =   4
      Top             =   210
      Width           =   1680
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPhoneLogHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Public Property Let SampleID(ByVal strNewValue As String)

10    pSampleID = strNewValue

End Property

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdSearch_Click()

10    If Trim$(txtSampleID) = "" Then Exit Sub

20    FillG

End Sub




Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim strRecord As String
      Dim intN As Integer

10    On Error GoTo FillG_Error

20    With grdPhoneLog
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60      .Visible = False
  
70      If optCallIn Then
80        sql = "SELECT * FROM PhoneLog WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "And ReasonForCall IS NOT NULL " & _
              "ORDER BY DateTime DESC"
90      ElseIf optCallOut Then
100       sql = "SELECT * FROM PhoneLog WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "And Discipline IS NOT NULL " & _
              "ORDER BY DateTime DESC"
110     ElseIf optBoth Then
120       sql = "SELECT * FROM PhoneLog WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "ORDER BY DateTime DESC"
130     End If
  

140     Set tb = New Recordset
150     RecOpenServerBB 0, tb, sql
160     Do While Not tb.EOF
170       strRecord = Format$(tb!DateTime, "dd/mm/yy HH:nn:ss")
180       For intN = 0 To 6
190         strRecord = strRecord & vbTab
200       Next
210       strRecord = strRecord & tb!ReasonForCall & "" & vbTab & _
                      tb!PhonedTo & vbTab & _
                      tb!Comment & vbTab & _
                      tb!PhonedBy & ""
220       .AddItem strRecord
  
230       .Row = .Rows - 1
240       For intN = 2 To 6
250         If InStr(tb!Discipline, Mid$("GADKT", intN, 1)) Then
260           .Col = intN
270           Set .CellPicture = imgSquareTick.Picture
280           .CellPictureAlignment = flexAlignCenterCenter
290         End If
300       Next
310       If IsNull(tb!Discipline) Then
320           .Col = 1
330           Set .CellPicture = imgCallIn.Picture
340           .CellPictureAlignment = flexAlignCenterCenter
350       ElseIf IsNull(tb!ReasonForCall) Then
360           .Col = 1
370           Set .CellPicture = imgCallOut.Picture
380           .CellPictureAlignment = flexAlignCenterCenter
390       End If
400       tb.MoveNext

410     Loop

420     If .Rows > 2 Then
430       .RemoveItem 1
440     End If
450     .Visible = True
460   End With

470   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmPhoneLogHistory", "FillG", intEL, strES, sql

End Sub


Private Sub Form_Load()
      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
10        txtSampleID = pSampleID

20        FillG
      '**************************************
End Sub

