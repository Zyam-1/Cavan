VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPrinters 
   Caption         =   "NetAcquire - Printers"
   ClientHeight    =   8475
   ClientLeft      =   615
   ClientTop       =   1155
   ClientWidth     =   10110
   Icon            =   "fPrinters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10110
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   7995
      Picture         =   "fPrinters.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   8010
      Picture         =   "fPrinters.frx":0F34
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
         Picture         =   "fPrinters.frx":159E
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8010
      Picture         =   "fPrinters.frx":19E0
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   660
      TabIndex        =   16
      Top             =   8310
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6870
      Picture         =   "fPrinters.frx":204A
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
Attribute VB_Name = "fPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CopyToName()

      Dim n As Integer
      Dim Found As Boolean

10    For n = 0 To lAvailable.ListCount - 1
20      If lAvailable.Selected(n) Then
30        tPrinterName = lAvailable.List(n)
40        lAvailable.Selected(n) = False
50        Found = True
60        Exit For
70      End If
80    Next

90    If Not Found Then
100     iMsg "Make a selection from the available printers.", vbInformation
110     If TimedOut Then Unload Me: Exit Sub
120   End If

End Sub

Private Sub bAdd_Click()

10    tMappedTo = Trim$(UCase$(tMappedTo))
20    tPrinterName = Trim$(UCase$(tPrinterName))

30    If tMappedTo = "" Then
40      Exit Sub
50    End If
60    If tPrinterName = "" Then
70      Exit Sub
80    End If

90    g.AddItem tMappedTo & vbTab & tPrinterName

100   tMappedTo = ""
110   tPrinterName = ""

120   cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()

10    Unload Me

End Sub




Private Sub cmdSave_Click()

      Dim Y As Integer
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmdSave_Click_Error

20    For Y = 1 To g.Rows - 1
  
30      sql = "Select * from Printers where " & _
              "MappedTo = '" & g.TextMatrix(Y, 0) & "'"
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If tb.EOF Then
70        tb.AddNew
80      End If

90      tb!MappedTo = UCase$(g.TextMatrix(Y, 0))
100     tb!PrinterName = UCase$(g.TextMatrix(Y, 1))
110     tb.Update

120   Next

130   FillG

140   tMappedTo = ""
150   tPrinterName = ""
160   tMappedTo.SetFocus
170   cmdSave.Enabled = False

180   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "fPrinters", "cmdSave_Click", intEL, strES, sql


End Sub


Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from Printers"
60    Set tb = New Recordset
70    RecOpenClient 0, tb, sql
80    Do While Not tb.EOF
90      s = tb!MappedTo & vbTab & tb!PrinterName & ""
100     g.AddItem s
110     tb.MoveNext
120   Loop

130   If g.Rows > 2 Then
140     g.RemoveItem 1
150   End If

160   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fPrinters", "FillG", intEL, strES, sql


End Sub




Private Sub Form_Load()

      Dim Px As Printer

10    lCurrent = Printer.DeviceName

20    lAvailable.Clear
30    For Each Px In Printers
40      lAvailable.AddItem Px.DeviceName
50    Next

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
60        FillG
      '**************************************

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If TimedOut Then Exit Sub
20    If cmdSave.Enabled Then
30      Answer = iMsg("Cancel without saving?", vbQuestion + vbYesNo)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = vbNo Then
60        Cancel = True
70      End If
80    End If

End Sub

Private Sub g_Click()

      Dim OldName As String
      Dim NewName As String

10    If g.MouseRow = 0 Then Exit Sub
20    If g.MouseCol = 0 Then Exit Sub

30    OldName = g.TextMatrix(g.Row, 1)
40    NewName = iBOX("PROCEED WITH CAUTION" & vbCrLf & vbCrLf & "New Printer Name?", , OldName)
50    If TimedOut Then Unload Me: Exit Sub
60    If Trim$(NewName) = "" Then
70      Exit Sub
80    End If

90    Answer = iMsg("Change " & vbCrLf & OldName & vbCrLf & "to" & vbCrLf & NewName, vbQuestion + vbYesNo)
100   If TimedOut Then Unload Me: Exit Sub
110   If Answer = vbNo Then Exit Sub

120   g.TextMatrix(g.Row, 1) = NewName
130   cmdSave.Enabled = True

End Sub


Private Sub iCopy_Click()

10    CopyToName

End Sub
Private Sub lCopy_Click()

10    CopyToName

End Sub


