VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmXMDaily 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Cross-Match Report"
   ClientHeight    =   6045
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   13140
   ControlBox      =   0   'False
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
   Icon            =   "FPXMDAIL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   13140
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
      Height          =   585
      Left            =   5160
      Picture         =   "FPXMDAIL.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11970
      Picture         =   "FPXMDAIL.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4080
      Picture         =   "FPXMDAIL.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   147324929
      CurrentDate     =   36963
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4905
      Left            =   15
      TabIndex        =   1
      Top             =   690
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   8652
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"FPXMDAIL.frx":18A8
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   15
      TabIndex        =   6
      Top             =   5640
      Width           =   12990
      _ExtentX        =   22913
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
      Left            =   6420
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prepare daily report for"
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1590
   End
End
Attribute VB_Name = "frmXMDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub FillG()

      Dim sn As Recordset
      Dim tb As Recordset
      Dim sql As String
      Dim strFromTime As String
      Dim strToTime As String
      Dim s As String
      Dim labnumprinted As Integer

10    On Error GoTo FillG_Error

20    Grid1.Rows = 2
30    Grid1.AddItem ""
40    Grid1.RemoveItem 1

50    strFromTime = Format(dtDate, "dd/mmm/yyyy") & " 00:00:00"
60    strToTime = Format(dtDate, "dd/mmm/yyyy") & " 23:59:59"

70    sql = "select * from patientdetails where " & _
            "datetime between '" & _
            strFromTime & "' and '" & strToTime & "'"

80    Set sn = New Recordset
90    RecOpenServerBB 0, sn, sql

100   If sn.EOF Then
      '  imsg "No records found!",  vbInformation
      '  If TimedOut Then Unload Me: Exit Sub
110     Exit Sub
120   End If

130   Do While Not sn.EOF
140     labnumprinted = False
150     sql = "Select * from Product where " & _
              "LabNumber = '" & sn!LabNumber & "' " & _
              "and event = 'X' order by Counter desc"
160     Set tb = New Recordset
170     RecOpenServerBB 0, tb, sql
180     Do While Not tb.EOF
190       If Not labnumprinted Then
200         s = sn!LabNumber & vbTab & _
                sn!Name & vbTab & _
                sn!Patnum & vbTab & _
                sn!Ward & vbTab & _
                sn!fGroup & vbTab
210       Else
220         s = vbTab & vbTab & vbTab & vbTab & vbTab
230       End If
240       s = s & tb!ISBT128 & "" & vbTab
250       s = s & Bar2Group(tb!GroupRh & "") & vbTab
260       If Not labnumprinted Then
270         s = s & ProductWordingFor(tb!BarCode & "") & vbTab
280         s = s & tb!Operator & ""
290         labnumprinted = True
300       End If
310       Grid1.AddItem s
320       tb.MoveNext
330     Loop
340     sn.MoveNext
350   Loop

360   If Grid1.Rows > 2 Then
370     Grid1.RemoveItem 1
380   End If


390   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "fpxmdaily", "FillG", intEL, strES, sql


End Sub


Private Sub cmdPrint_Click()

      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim Y As Integer

10    OriginalPrinter = Printer.DeviceName

20    If Not SetFormPrinter() Then Exit Sub

30    Printer.FontName = "Courier New"
40    Printer.Font.Size = 9
50    Printer.Orientation = vbPRORPortrait
60    Printer.Font.Bold = True
70    Printer.Print
80    Printer.Print "Daily Crossmatch Report for " & Format(dtDate, "long date")
90    Printer.Print

100   For Y = 1 To Grid1.Rows - 1
          'grid1.Row = Y
          'if its parent record then write patient info in paragraph format.
110       If Grid1.TextMatrix(Y, 0) <> "" Then
              'patient info goes here
120           Printer.Print "____________________________________________________________________________________________________________"
130           Printer.Font.Italic = True
140           Printer.Print
150           Printer.Font.Bold = False
160           Printer.Print "Lab#: ";
170           Printer.Font.Bold = True
180           Printer.Print Grid1.TextMatrix(Y, 0); " "; 'lab no
190           Printer.Font.Bold = False
200           Printer.Print "Name: ";
210           Printer.Font.Bold = True 'Patient Name
220           Printer.Print Grid1.TextMatrix(Y, 1)
230           Printer.Font.Bold = False
240           Printer.Print "Pat.ID: ";
250           Printer.Font.Bold = True 'Patient ID
260           Printer.Print Grid1.TextMatrix(Y, 2)
270           Printer.Font.Bold = False
280           Printer.Print "Location: ";
290           Printer.Font.Bold = True 'Location
300           Printer.Print Grid1.TextMatrix(Y, 3); " ";
310           Printer.Font.Bold = False
320           Printer.Print "Group: ";
330           Printer.Font.Bold = True 'Group
340           Printer.Print Grid1.TextMatrix(Y, 4); " ";
350           Printer.Font.Bold = False
360           Printer.Font.Italic = False
370           Printer.Print
380           Printer.Print
  
390           Printer.Print FormatString(" ", 25);
400           Printer.Print FormatString("Unit", 16, "|");
410           Printer.Print FormatString("Prod Group", 10, "|");
420           Printer.Print FormatString("Product", 50, "|");
430           Printer.Print FormatString("Op.", 5, "|")
  
440           Printer.Print FormatString(" ", 25); "-------------------------------------------------------------------------------"
450           Printer.Print FormatString(" ", 25);
460           Printer.Print FormatString(Grid1.TextMatrix(Y, 5), 16, "|"); 'Unit
470           Printer.Print FormatString(Grid1.TextMatrix(Y, 6), 10, "|"); 'Product group
480           Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 50, "|"); 'product
490           Printer.Print FormatString(Grid1.TextMatrix(Y, 8), 5, "|") 'operator
500       Else
  
              'units detail goes here
510           Printer.Print FormatString(" ", 25);
520           Printer.Print FormatString(Grid1.TextMatrix(Y, 5), 16, "|"); 'Unit
530           Printer.Print FormatString(Grid1.TextMatrix(Y, 6), 10, "|"); 'product group
540           Printer.Print FormatString(Grid1.TextMatrix(Y, 7), 50, "|"); 'product
550           Printer.Print FormatString(Grid1.TextMatrix(Y, 8), 5, "|") 'operator
560       End If
  
  
  
570   Next
580   Printer.EndDoc

590   For Each Px In Printers
600     If Px.DeviceName = OriginalPrinter Then
610       Set Printer = Px
620       Exit For
630     End If
640   Next

End Sub


Private Sub cmdXL_Click()
      Dim strHeading As String
10    strHeading = "Daily Crossmatch Report for " & Format(dtDate, "long date") & vbCr
20    strHeading = strHeading & " " & vbCr
30    ExportFlexGrid Grid1, Me, strHeading

End Sub

Private Sub dtDate_CloseUp()

10    FillG

End Sub



Private Sub Form_Load()

10    dtDate = Format(Now, "dd/mmm/yyyy")

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
20        FillG
      '**************************************

End Sub

