VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrinterSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Selection"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmPrinterSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbPrinterList 
      Height          =   315
      Index           =   2
      Left            =   195
      TabIndex        =   14
      Text            =   "cmbPrinterList"
      Top             =   2595
      Width           =   5775
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Index           =   2
      Left            =   6030
      Picture         =   "frmPrinterSelect.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2490
      Width           =   795
   End
   Begin MSComCtl2.UpDown udAuto 
      Height          =   225
      Left            =   2790
      TabIndex        =   10
      Top             =   3705
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      _Version        =   393216
      Value           =   2
      Alignment       =   0
      BuddyControl    =   "lblAutoPrintLabels"
      BuddyDispid     =   196614
      OrigLeft        =   4860
      OrigTop         =   2400
      OrigRight       =   5100
      OrigBottom      =   3165
      Max             =   4
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65537
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   6030
      Picture         =   "frmPrinterSelect.frx":194C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   3705
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Index           =   1
      Left            =   6030
      Picture         =   "frmPrinterSelect.frx":29CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1590
      Width           =   795
   End
   Begin VB.ComboBox cmbPrinterList 
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Text            =   "cmbPrinterList"
      Top             =   1710
      Width           =   5775
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Index           =   0
      Left            =   6030
      Picture         =   "frmPrinterSelect.frx":3A50
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   795
   End
   Begin VB.ComboBox cmbPrinterList 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Text            =   "cmbPrinterList"
      Top             =   840
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   4065
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblTitle 
      Caption         =   "Select printers for use on this computer: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   15
      Top             =   180
      Width           =   6885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "PDF Printer"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   2370
      Width           =   810
   End
   Begin VB.Label lblAutoPrintLabels 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   3390
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "labels after Saving"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3300
      TabIndex        =   9
      Top             =   3405
      Width           =   1665
   End
   Begin VB.Label label3 
      AutoSize        =   -1  'True
      Caption         =   "Automatically print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1110
      TabIndex        =   8
      Top             =   3405
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label Printer"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1470
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Form Printer"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   630
      Width           =   840
   End
End
Attribute VB_Name = "frmPrinterSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPrinterList_KeyPress(Index As Integer, KeyAscii As Integer)

10  If Index < 2 Then
20      KeyAscii = 0
30  End If

End Sub


Private Sub cmdExit_Click()

10  Unload Me

End Sub

Private Sub cmdSave_Click(Index As Integer)

    Dim pname As String

10  pname = Choose(Index + 1, "Form", "Label", "PDF")

20  SaveUserOptionSetting "Transfusion" & pname, cmbPrinterList(Index), UCase$(vbGetComputerName())

30  TransfusionForm = cmbPrinterList(0)
40  TransfusionLabel = cmbPrinterList(1)
50  TransfusionPDF = cmbPrinterList(2)

60  iMsg pname & " printer saved!"

End Sub


Private Sub Form_Load()

    Dim Px As Printer

10  lblTitle = "Select printers for use on this computer: " & vbGetComputerName()

20  cmbPrinterList(0).Clear
30  cmbPrinterList(1).Clear
40  cmbPrinterList(2).Clear

50  For Each Px In Printers
60      cmbPrinterList(0).AddItem Px.DeviceName
70      cmbPrinterList(1).AddItem Px.DeviceName
80      cmbPrinterList(2).AddItem Px.DeviceName
90  Next

100 If TransfusionForm = "" Then
110     cmbPrinterList(0).ListIndex = -1
120 Else
130     cmbPrinterList(0) = TransfusionForm
140 End If

150 If TransfusionLabel = "" Then
160     cmbPrinterList(1).ListIndex = -1
170 Else
180     cmbPrinterList(1) = TransfusionLabel
190 End If

200 If TransfusionPDF = "" Then
210     cmbPrinterList(2).ListIndex = -1
220 Else
230     cmbPrinterList(2) = TransfusionPDF
240 End If

250 udAuto.Value = Val(GetOptionSetting("AutoPrintLabels", "2"))
260 lblAutoPrintLabels.Caption = Format$(udAuto.Value)

End Sub


Private Sub udAuto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10  SaveOptionSetting "AutoPrintLabels", lblAutoPrintLabels.Caption

End Sub


