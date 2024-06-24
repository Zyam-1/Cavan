VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrinters 
   Caption         =   "NetAcquire - Printers"
   ClientHeight    =   8475
   ClientLeft      =   615
   ClientTop       =   1155
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10110
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6330
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Printer"
      Height          =   1515
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      Begin VB.ListBox lAvailable 
         Height          =   1185
         IntegralHeight  =   0   'False
         Left            =   4620
         TabIndex        =   13
         Top             =   240
         Width           =   4965
      End
      Begin VB.TextBox tMappedTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         MaxLength       =   7
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox tPrinterName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1050
         Width           =   3495
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   2970
         TabIndex        =   2
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lCopy 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy"
         Height          =   285
         Left            =   4140
         TabIndex        =   15
         Top             =   1140
         Width           =   480
      End
      Begin VB.Image iCopy 
         Height          =   480
         Left            =   3690
         Picture         =   "frmPrinters.frx":0CD4
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Available Printers"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mapped To"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Printer Name"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   870
         Width           =   915
      End
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5250
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6525
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1770
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11509
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Mapped To |<Printer Name                                                                    "
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6870
      Picture         =   "frmPrinters.frx":1780
      Top             =   2910
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Click on Specific Printer Name to Edit"
      Height          =   375
      Left            =   7350
      TabIndex        =   12
      Top             =   2970
      Width           =   1545
   End
   Begin VB.Label lCurrent 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   1980
      Width           =   2925
   End
   Begin VB.Label Label3 
      Caption         =   "Current Default Printer"
      Height          =   195
      Left            =   6990
      TabIndex        =   10
      Top             =   1770
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private Sub CopyToName()

      Dim n As Integer
      Dim Found As Boolean

38410 For n = 0 To lAvailable.ListCount - 1
38420   If lAvailable.Selected(n) Then
38430     tPrinterName = lAvailable.List(n)
38440     lAvailable.Selected(n) = False
38450     Found = True
38460     Exit For
38470   End If
38480 Next

38490 If Not Found Then
38500   MsgBox "Make a selection from the available printers.", vbInformation
38510 End If

End Sub

Private Sub bAdd_Click()

38520 tMappedTo = Trim$(UCase$(tMappedTo))
38530 tPrinterName = Trim$(UCase$(tPrinterName))

38540 If tMappedTo = "" Then
38550   Exit Sub
38560 End If
38570 If tPrinterName = "" Then
38580   Exit Sub
38590 End If

38600 g.AddItem tMappedTo & vbTab & tPrinterName

38610 tMappedTo = ""
38620 tPrinterName = ""

38630 bsave.Enabled = True

End Sub

Private Sub bcancel_Click()

38640 Unload Me

End Sub


Private Sub bSave_Click()

      Dim Y As Integer
      Dim sql As String
      Dim tb As Recordset

38650 On Error GoTo bSave_Click_Error

38660 For Y = 1 To g.Rows - 1
        
38670   sql = "Select * from Printers where " & _
              "MappedTo = '" & g.TextMatrix(Y, 0) & "'"
38680   Set tb = New Recordset
38690   RecOpenServer 0, tb, sql
38700   If tb.EOF Then
38710     tb.AddNew
38720   End If

38730   tb!MappedTo = UCase$(g.TextMatrix(Y, 0))
38740   tb!PrinterName = UCase$(g.TextMatrix(Y, 1))
38750   tb.Update

38760 Next

38770 FillG

38780 tMappedTo = ""
38790 tPrinterName = ""
38800 tMappedTo.SetFocus
38810 bsave.Enabled = False

38820 Exit Sub

bSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

38830 intEL = Erl
38840 strES = Err.Description
38850 LogError "fPrinters", "bSave_Click", intEL, strES, sql


End Sub


Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

38860 On Error GoTo FillG_Error

38870 g.Rows = 2
38880 g.AddItem ""
38890 g.RemoveItem 1

38900 sql = "Select * from Printers"
38910 Set tb = New Recordset
38920 RecOpenClient 0, tb, sql
38930 Do While Not tb.EOF
38940   s = tb!MappedTo & vbTab & tb!PrinterName & ""
38950   g.AddItem s
38960   tb.MoveNext
38970 Loop

38980 If g.Rows > 2 Then
38990   g.RemoveItem 1
39000 End If

39010 Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

39020 intEL = Erl
39030 strES = Err.Description
39040 LogError "fPrinters", "FillG", intEL, strES, sql


End Sub


Private Sub Form_Activate()

39050 If Activated Then Exit Sub

39060 FillG

39070 Activated = True

End Sub

Private Sub Form_Load()

      Dim Px As Printer

39080 lCurrent = Printer.DeviceName

39090 lAvailable.Clear
39100 For Each Px In Printers
39110   lAvailable.AddItem Px.DeviceName
39120 Next

39130 Activated = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

39140 If bsave.Enabled Then
39150   If MsgBox("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
39160     Cancel = True
39170   End If
39180 End If

End Sub

Private Sub g_Click()

      Dim OldName As String
      Dim NewName As String

39190 If g.MouseRow = 0 Then Exit Sub
39200 If g.MouseCol = 0 Then Exit Sub

39210 OldName = g.TextMatrix(g.row, 1)
39220 NewName = iBOX("PROCEED WITH CAUTION" & vbCrLf & vbCrLf & "New Printer Name?", , OldName)
39230 If Trim$(NewName) = "" Then
39240   Exit Sub
39250 End If

39260 If MsgBox("Change " & vbCrLf & OldName & vbCrLf & "to" & vbCrLf & NewName, vbQuestion + vbYesNo) = vbNo Then Exit Sub

39270 g.TextMatrix(g.row, 1) = NewName
39280 bsave.Enabled = True

End Sub


Private Sub iCopy_Click()

39290 CopyToName

End Sub
Private Sub lCopy_Click()

39300 CopyToName

End Sub


