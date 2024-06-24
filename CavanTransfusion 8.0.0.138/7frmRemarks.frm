VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remarks Entry"
   ClientHeight    =   3570
   ClientLeft      =   1485
   ClientTop       =   1425
   ClientWidth     =   7245
   ControlBox      =   0   'False
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "7frmRemarks.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   7245
   Begin VB.TextBox txtRemarks 
      Height          =   2595
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6945
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy and Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2670
      TabIndex        =   0
      Top             =   2880
      Width           =   1605
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   150
      TabIndex        =   2
      Top             =   3360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()

10    Me.Hide

End Sub


Public Property Get Comment() As String

10    Comment = txtRemarks

End Property

Public Property Let Comment(ByVal strComment As String)

10    txtRemarks = strComment

End Property

Public Property Let Heading(ByVal strHeading As String)

10    Me.Caption = strHeading

End Property


