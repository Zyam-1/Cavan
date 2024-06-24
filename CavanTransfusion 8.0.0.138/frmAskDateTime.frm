VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAskDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Transfusion"
   ClientHeight    =   2010
   ClientLeft      =   3960
   ClientTop       =   4485
   ClientWidth     =   7515
   ControlBox      =   0   'False
   Icon            =   "frmAskDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInvalid 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   480
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   6420
      Picture         =   "frmAskDateTime.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   990
      Width           =   885
   End
   Begin MSMask.MaskEdBox tTime 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   930
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   930
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   110362625
      CurrentDate     =   36979
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   705
      Left            =   6420
      Picture         =   "frmAskDateTime.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   885
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   600
      TabIndex        =   7
      Top             =   1740
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblInvalid 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invalid Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1380
      TabIndex        =   6
      Top             =   510
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Time HH:MM"
      Height          =   195
      Left            =   3270
      TabIndex        =   1
      Top             =   990
      Width           =   945
   End
   Begin VB.Label lPrompt 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   150
      Width           =   5805
   End
End
Attribute VB_Name = "frmAskDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DateTime As String

Private m_Prompt As String
Public Property Let Prompt(ByVal strPrompt As String)

10    m_Prompt = strPrompt

End Property

Private Sub cmdCancel_Click()

10    m_DateTime = ""
20    Me.Hide

End Sub

Private Sub cmdSave_Click()

10    If tTime <> "__:__" Then
20        If Not IsDate(tTime) Then
30          m_DateTime = ""
40          lblInvalid.Visible = True
50          lblInvalid.Refresh
60          tmrInvalid.Enabled = True
70          Exit Sub
80        End If
90    End If

100   If tTime <> "__:__" Then
110       m_DateTime = Format(dtDate, "dd/mmm/yyyy") & " " & Format(tTime, "hh:mm")
120   Else
130       m_DateTime = Format(dtDate, "dd/mmm/yyyy")
140   End If

150   Me.Hide

End Sub



Public Property Get DateTime() As String

10    DateTime = m_DateTime

End Property

Public Property Let DateTime(ByVal EntryDateTime As String)

10    m_DateTime = EntryDateTime

End Property


Private Sub Form_Activate()

If Not IsDate(m_DateTime) Then
  m_DateTime = Format(Now, "dd/MM/yyyy HH:nn")
End If

10    dtDate = Format(m_DateTime, "dd/MMM/yyyy")
20    tTime.Mask = ""
30    tTime.Text = Format(m_DateTime, "hh:nn")
40    tTime.Mask = "##:##"
50    lPrompt = m_Prompt

End Sub

Private Sub tmrInvalid_Timer()

10    lblInvalid.Visible = False
20    lblInvalid.Refresh
30    tmrInvalid.Enabled = False

End Sub


