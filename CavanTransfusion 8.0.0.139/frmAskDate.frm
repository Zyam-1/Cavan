VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAskDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   3045
   ClientLeft      =   1515
   ClientTop       =   1275
   ClientWidth     =   3675
   ControlBox      =   0   'False
   Icon            =   "frmAskDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2370
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   92536834
      TrailingForeColor=   -2147483647
      CurrentDate     =   38305
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   1
      Top             =   2760
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmAskDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

End Sub

Public Property Get DisplayDate() As String

10    DisplayDate = mvDate

End Property

Public Property Let DisplayDate(ByVal OpeningDate As String)

10    If IsDate(OpeningDate) Then
20      mvDate = OpeningDate
30    Else
40      mvDate = Now
50    End If

End Property

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

10    Me.Hide

End Sub

