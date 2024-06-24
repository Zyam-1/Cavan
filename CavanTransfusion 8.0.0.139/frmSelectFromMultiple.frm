VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmSelectFromMultiple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "          Cancel         Do not select any"
      Height          =   1365
      HelpContextID   =   10090
      Left            =   12150
      Picture         =   "frmSelectFromMultiple.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3930
      Width           =   1485
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Product"
      Height          =   1065
      Left            =   12150
      Picture         =   "frmSelectFromMultiple.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2130
      Width           =   1515
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3645
      Left            =   420
      TabIndex        =   0
      Top             =   1620
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmSelectFromMultiple.frx":1794
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   390
      TabIndex        =   7
      Top             =   5310
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2685
      Picture         =   "frmSelectFromMultiple.frx":1841
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Unit of interest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   4
      Top             =   810
      Width           =   2925
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Multiple Products have been found for this Unit Number."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4650
      TabIndex        =   3
      Top             =   450
      Width           =   6360
   End
   Begin VB.Label lblUnitNumber 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
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
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   360
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Number"
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   450
      Width           =   885
   End
End
Attribute VB_Name = "frmSelectFromMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_ProductList As Products
Private m_SelectedProduct As Product

Private Sub cmdCancel_Click()

Set m_SelectedProduct = Nothing
Me.Hide

End Sub

Public Property Let ProductList(ByVal Ps As Products)

Set m_ProductList = Ps

End Property

Public Property Get SelectedProduct() As Product

  Set SelectedProduct = m_SelectedProduct

End Property



Private Sub cmdSelect_Click()

Dim Y As Integer

Set m_SelectedProduct = Nothing

g.col = 0
For Y = 1 To g.Rows - 1
  g.row = Y
  If g.CellBackColor = vbYellow Then
    Set m_SelectedProduct = m_ProductList(Y)
    Exit For
  End If
Next

If Not m_SelectedProduct Is Nothing Then
  Me.Hide
End If

End Sub


Private Sub Form_Activate()

Dim p As Product
Dim s As String

g.Rows = 2
g.AddItem ""
g.RemoveItem 1

For Each p In m_ProductList
  lblUnitNumber.Caption = p.ISBT128
  s = ProductWordingFor(p.BarCode) & vbTab & _
      Format(p.DateExpiry, "dd/mm/yyyy HH:mm") & vbTab & _
      SupplierNameFor(p.Supplier) & vbTab & _
      gEVENTCODES(p.PackEvent).Text
  g.AddItem s
Next

If g.Rows > 2 Then
  g.RemoveItem 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set m_SelectedProduct = Nothing

End Sub


Private Sub g_Click()

Dim Y As Integer
Dim X As Integer
Dim ySave As Integer

If g.MouseRow = 0 Then Exit Sub

ySave = g.row

g.col = 0
For Y = 1 To g.Rows - 1
  g.row = Y
  If g.CellBackColor = vbYellow Then
    For X = 0 To g.Cols - 1
      g.col = X
      g.CellBackColor = 0
    Next
    Exit For
  End If
Next

g.row = ySave
For X = 0 To g.Cols - 1
  g.col = X
  g.CellBackColor = vbYellow
Next

End Sub


