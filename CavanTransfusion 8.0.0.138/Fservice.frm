VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fservice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Information"
   ClientHeight    =   6735
   ClientLeft      =   930
   ClientTop       =   1095
   ClientWidth     =   8400
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
   Icon            =   "Fservice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7020
      Picture         =   "Fservice.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   360
      Width           =   1065
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7020
      Picture         =   "Fservice.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1620
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3015
      Left            =   300
      TabIndex        =   19
      Top             =   3240
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   5318
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
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"Fservice.frx":123E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   18
      Top             =   2190
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5220
      TabIndex        =   17
      Top             =   330
      Width           =   1065
   End
   Begin VB.TextBox txtNextDue 
      DataField       =   "nextdue"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      TabIndex        =   16
      Top             =   2460
      Width           =   3135
   End
   Begin VB.TextBox txtComments 
      DataField       =   "comments"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtServicedBy 
      DataField       =   "servicedby"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtLastServiced 
      DataField       =   "lastservice"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   8
      TabIndex        =   14
      Top             =   1860
      Width           =   3135
   End
   Begin VB.TextBox txtInterval 
      DataField       =   "interval"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtSupplier 
      DataField       =   "supplier"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1260
      Width           =   3135
   End
   Begin VB.TextBox txtSerial 
      DataField       =   "serial"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   11
      Top             =   330
      Width           =   3135
   End
   Begin VB.TextBox txtInstrument 
      DataField       =   "instrument"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   10
      Top             =   810
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Top             =   2760
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   315
      TabIndex        =   20
      Top             =   6420
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   6960
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1125
      TabIndex        =   9
      Top             =   2820
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Next Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1185
      TabIndex        =   8
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Serviced By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   990
      TabIndex        =   7
      Top             =   2220
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Date Last Serviced"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   6
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Service Interval"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   705
      TabIndex        =   5
      Top             =   1620
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1245
      TabIndex        =   4
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Serial Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   390
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Instrument"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   2
      Top             =   870
      Width           =   735
   End
End
Attribute VB_Name = "fservice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()
 
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.FontSize = 9
50    Printer.Orientation = vbPRORPortrait

      '****Report heading
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "                                     Service Information"

      '****Report body


90    For i = 1 To 108
100       Printer.Print "_";
110   Next i
120   Printer.Print
130   Printer.Print FormatString("Last Service", 30, "|");
140   Printer.Print FormatString("Service By", 10, "|");
150   Printer.Print FormatString("Next Due", 18, "|");
160   Printer.Print FormatString("Comment", 46, "|")
170   Printer.Font.Bold = False
180   For i = 1 To 108
190       Printer.Print "-";
200   Next i
210   Printer.Print
220   For Y = 1 To g.Rows - 1
230       Printer.Print FormatString(g.TextMatrix(Y, 0), 30, "|");
240       Printer.Print FormatString(g.TextMatrix(Y, 1), 10, "|");
250       Printer.Print FormatString(g.TextMatrix(Y, 2), 18, "|");
260       Printer.Print FormatString(g.TextMatrix(Y, 3), 46, "|")
 
270   Next


280   Printer.EndDoc



290   For Each Px In Printers
300     If Px.DeviceName = OriginalPrinter Then
310       Set Printer = Px
320       Exit For
330     End If
340   Next
End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmdSave_Click_Error

20    If Trim$(txtNextDue) <> "" And Not IsDate(txtNextDue) Then
30      iMsg "Invalid Next Due Date!", vbExclamation
40      If TimedOut Then Unload Me: Exit Sub
50      txtNextDue = ""
60      Exit Sub
70    End If

80    If Trim$(txtLastServiced) <> "" And Not IsDate(txtLastServiced) Then
90      iMsg "Invalid Last Serviced Date!", vbExclamation
100     If TimedOut Then Unload Me: Exit Sub
110     txtLastServiced = ""
120     Exit Sub
130   End If

140   sql = "SELECT * FROM Servicedetails WHERE " & _
            "serial = 'x'"
150   Set tb = New Recordset
160   RecOpenServerBB 0, tb, sql
170   tb.AddNew
180   tb!serial = txtSerial
190   tb!Instrument = txtInstrument
200   tb!Supplier = txtSupplier
210   tb!Interval = txtInterval
220   tb!LastService = Format(txtLastServiced, "dd/mmm/yyyy")
230   tb!ServicedBy = txtServicedBy
240   tb!NextDue = Format(txtNextDue, "dd/mmm/yyyy")
250   tb!Comments = txtComments
260   tb.Update

270   cmdSave.Enabled = False

280   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fservice", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub cmdSearch_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo cmdSearch_Click_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "SELECT * FROM Servicedetails WHERE " & _
            "serial = '" & txtSerial & "' " & _
            "Order by LastService desc"
60    Set tb = New Recordset
70    RecOpenServerBB 0, tb, sql
80    If tb.EOF Then
90      iMsg "Serial Number not known!", vbExclamation
100     If TimedOut Then Unload Me: Exit Sub
110     txtInstrument = ""
120     txtSupplier = ""
130     txtInterval = ""
140     Exit Sub
150   Else
160     txtInstrument = tb!Instrument & ""
170     txtSupplier = tb!Supplier & ""
180     txtInterval = tb!Interval & ""
190   End If
200   Do While Not tb.EOF
210     s = tb!LastService & vbTab & _
            tb!ServicedBy & vbTab & _
            tb!NextDue & vbTab & _
            tb!Comments & ""
220     g.AddItem s
230     tb.MoveNext
240   Loop

250   If g.Rows > 2 Then
260     g.RemoveItem 1
270   End If

280   Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "fservice", "cmdSearch_Click", intEL, strES, sql


End Sub

Private Sub cmdXL_Click()
       Dim strHeading As String

10    strHeading = "Service Information" & vbCr
20    strHeading = strHeading & " " & vbCr
30    ExportFlexGrid g, Me, strHeading
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


Private Sub txtInstrument_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


Private Sub txtInterval_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub

Private Sub txtLastServiced_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


Private Sub txtNextDue_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


Private Sub txtServicedBy_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub

Private Sub txtSupplier_KeyPress(KeyAscii As Integer)

10    cmdSave.Enabled = True

End Sub


