VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnitHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit History"
   ClientHeight    =   6075
   ClientLeft      =   270
   ClientTop       =   390
   ClientWidth     =   11265
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
   Icon            =   "7frmUnitHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
   ScaleWidth      =   11265
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
      Height          =   735
      Left            =   5760
      Picture         =   "7frmUnitHistory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   90
      Width           =   825
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3645
      Left            =   120
      TabIndex        =   11
      Top             =   2115
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"7frmUnitHistory.frx":0BD4
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
      Height          =   735
      Left            =   8760
      Picture         =   "7frmUnitHistory.frx":0C91
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   825
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4770
      Picture         =   "7frmUnitHistory.frx":12FB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   825
   End
   Begin VB.TextBox txtUnitNumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   1290
      TabIndex        =   2
      Top             =   60
      Width           =   3270
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
      Height          =   735
      Left            =   10290
      Picture         =   "7frmUnitHistory.frx":173D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   5805
      Width           =   11025
      _ExtentX        =   19447
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
      Left            =   6990
      TabIndex        =   14
      Top             =   330
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lProduct 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   1305
      Width           =   10965
   End
   Begin VB.Label lantigen 
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
      Height          =   405
      Left            =   150
      TabIndex        =   10
      Top             =   1665
      Width           =   10965
   End
   Begin VB.Label lblExpiry 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   510
      Width           =   1875
   End
   Begin VB.Label lsupplier 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Top             =   885
      Width           =   3255
   End
   Begin VB.Label lgroup 
      Alignment       =   2  'Center
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
      Left            =   3540
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Expiry "
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
      Left            =   600
      TabIndex        =   8
      Top             =   540
      Width           =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
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
      Left            =   420
      TabIndex        =   7
      Top             =   945
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ISBT-128"
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
      Left            =   360
      TabIndex        =   1
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmUnitHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mProduct As String
Dim mUnitNumber As String
Public Property Let UnitNumber(ByVal Value As String)

10    mUnitNumber = Value
20    txtUnitNumber = Value

End Property

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim n As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

      '*****NOTE:
          'print format is standarized to 86 characters. a fix length is
          'given to each column. if something is to be written in center of
          'line then center point of line would be 86/2
      '********************************************

30    Printer.Orientation = vbPRORLandscape
40    Printer.Font.Size = 10
50    Printer.Font.Bold = True
60    Printer.Font.Name = "Courier New"
70    Printer.Print FormatString("Unit History", 146, , AlignCenter)
80    Printer.Print
90    Printer.Print FormatString("Unit Number:", 15, , AlignRight); txtUnitNumber;
100   Printer.Print
110   Printer.Print FormatString("Product:", 15, , AlignRight); lProduct;
120   If Trim$(lantigen) <> "" Then
130       Printer.Print
140       Printer.Print FormatString(" ", 15); "("; lantigen; ")"
150   End If
      'Printer.Print
160   Printer.Print FormatString("Group:", 15, , AlignRight); lgroup;
170   Printer.Print
180   Printer.Print FormatString("Expiry:", 15, , AlignRight); lblExpiry;
190   Printer.Print
200   Printer.Print
      'If Grid1.Cols = 6 Then Printer.Print "Pack"; 'multipack
210   Printer.Font.Size = 9
220   For i = 1 To 152
230       Printer.Print "-";
240   Next i
250   Printer.Print
260   Printer.Print FormatString("Event", 25, "|", AlignCenter);
270   Printer.Print FormatString("Time", 16, "|", AlignCenter);
280   Printer.Print FormatString("Pat ID", 10, "|", AlignCenter);
290   Printer.Print FormatString("Patient Name", 20, "|", AlignCenter);
300   Printer.Print FormatString("Op", 5, "|", AlignCenter);
310   Printer.Print FormatString("Notes", 26, "|", AlignCenter);
320   Printer.Print FormatString("Event Start", 21, "|", AlignCenter);
330   Printer.Print FormatString("Event End", 21, "|", AlignCenter)
340   For i = 1 To 152
350       Printer.Print "-";
360   Next i
370   Printer.Print
380   Printer.Font.Bold = False
390   For n = 1 To Grid1.Rows - 1
400     Grid1.Row = n
  

410     Grid1.Col = 0   'event
420     Printer.Print FormatString(Grid1.Text, 25, "|");
430     Grid1.Col = Grid1.Col + 1 'time
440     Printer.Print FormatString(Format(Grid1.Text, "dd/MM/yyyy hh:mm"), 16, "|");
450     Grid1.Col = Grid1.Col + 1 'patid
460     Printer.Print FormatString(Grid1.Text, 10, "|");
470     Grid1.Col = Grid1.Col + 1 'patient name
480     Printer.Print FormatString(Grid1.Text, 20, "|");
490     Grid1.Col = Grid1.Col + 1 'operator
500     Printer.Print FormatString(Grid1.Text, 5, "|");
510     Grid1.Col = Grid1.Col + 1 'notes
520     Printer.Print FormatString(Grid1.Text, 26, "|");
530     Grid1.Col = Grid1.Col + 1 'notes
540     Printer.Print FormatString(Grid1.Text, 21, "|");
550     Grid1.Col = Grid1.Col + 1 'notes
560     Printer.Print FormatString(Grid1.Text, 21, "|")
570   Next

580   Printer.EndDoc

590   For Each Px In Printers
600     If Px.DeviceName = OriginalPrinter Then
610       Set Printer = Px
620       Exit For
630     End If
640   Next

End Sub

Private Sub cmdSearch_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim strBarCode As String
Dim Ps As New Products
Dim p As Product
Dim f As Form

10    On Error GoTo cmdSearch_Click_Error
      
20    txtUnitNumber = Replace(txtUnitNumber, "'", "")

30    Grid1.Rows = 2
40    Grid1.AddItem ""
50    Grid1.RemoveItem 1

60    If mUnitNumber <> "" And mProduct <> "" Then
70      txtUnitNumber.Locked = True
80      Ps.Load mUnitNumber, ProductBarCodeFor(mProduct)
90    End If
100   If Ps.Count = 0 Then
110     Ps.LoadLatestByUnitNumberISBT128 txtUnitNumber
120     If Ps.Count = 0 Then
130       iMsg "Unit Number not found."
140       If TimedOut Then Unload Me: Exit Sub
150       txtUnitNumber = ""
160       Exit Sub
170     ElseIf Ps.Count > 1 Then 'multiple products found
180       Set f = New frmSelectFromMultiple
190       f.ProductList = Ps
200       f.Show 1
210       Set p = f.SelectedProduct
220       Unload f
230       Set f = Nothing
240     Else
250       Set p = Ps.Item(1)
260     End If
270   Else
280     Set p = Ps.Item(1)
290   End If
300   If Not p Is Nothing Then
310     strBarCode = p.BarCode
320     lblExpiry = Format(p.DateExpiry, "dd/mm/yyyy HH:mm")

330     lProduct = ProductWordingFor(strBarCode)
340     If p.Checked Then
350       s = "Group Checked" & vbTab & _
              Format(p.RecordDateTime, "dd/mm/yy hh:mm:ss") & vbTab & _
              vbTab & _
              vbTab & _
              p.UserName & vbTab & _
              p.Notes & "" & vbTab & _
              p.EventStart & "" & vbTab & _
              p.EventEnd
360           Grid1.AddItem s
370           Grid1.AddItem ""
380     End If
    
390     lgroup = Bar2Group(p.GroupRh)
400     lsupplier = SupplierNameFor(p.Supplier)
    
410   Set Ps = New Products
420   Ps.Load p.ISBT128, p.BarCode

430     For Each p In Ps
440       lantigen = p.Screen
450       s = gEVENTCODES(p.PackEvent).Text & vbTab & _
              Format(p.RecordDateTime, "dd/MM/yy hh:mm:ss") & vbTab & _
              p.Chart & vbTab & _
              p.PatName & vbTab & _
              p.UserName & vbTab & _
              p.Notes & vbTab
460       If IsDate(p.EventStart) Then
470         s = s & Format$(p.EventStart, "dd/MM/yy HH:nn:ss") & vbTab
480       End If
490       If IsDate(p.EventEnd) Then
500         s = s & Format$(p.EventEnd, "dd/MM/yy HH:nn:ss")
510       End If
520       Grid1.AddItem s
530     Next
540   End If

550   sql = "Select * From PartialPacks " & _
            "Where Number = '" & txtUnitNumber & "' " & _
            "order by Counter Desc"
560   Set tb = New Recordset
570   RecOpenServerBB 0, tb, sql
580   If Not (tb.EOF And tb.BOF) Then
590       Do While Not tb.EOF
600           s = gEVENTCODES(tb!Event).Text & vbTab & _
                  Format(tb!DateTime, "dd/MM/yy hh:mm:ss") & vbTab & _
                  tb!Patid & vbTab & _
                  tb!PatName & vbTab & _
                  tb!Operator & vbTab & _
                  "" & vbTab
610               Grid1.AddItem s
620               tb.MoveNext
630       Loop
640   End If

650   If Grid1.Rows = 2 Then
660     iMsg "No Record!", vbInformation
670     If TimedOut Then Unload Me: Exit Sub
680     Exit Sub
690   Else
700     Grid1.Col = 1
710     Grid1.Sort = 9
720   End If

730   If Grid1.Rows > 2 Then
740     Grid1.RemoveItem 1
750   End If

760   Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

770   intEL = Erl
780   strES = Err.Description
790   LogError "frmUnitHistory", "cmdSearch_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
  
      Dim strHeading As String
10    strHeading = "Product History" & vbCr & _
                  "" & vbCr & _
                  "Unit Number:  " & txtUnitNumber & vbCr & _
                  "Product:      " & lProduct.Caption & vbCr & _
                  "( " & lantigen & " )" & vbCr & _
                  "Group:        " & lgroup.Caption & vbCr & _
                  "Expiry:       " & lblExpiry & vbCr & _
                  "" & vbCr

20    ExportFlexGrid Grid1, Me, strHeading

End Sub






Private Sub Grid1_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
          Dim d1 As String
          Dim d2 As String
    
10        If Not IsDate(Grid1.TextMatrix(Row1, Grid1.Col)) Then
20         Cmp = 0
30         Exit Sub
40        End If
    
50        If Not IsDate(Grid1.TextMatrix(Row2, Grid1.Col)) Then
60         Cmp = 0
70         Exit Sub
80        End If
    
90        d1 = Format(Grid1.TextMatrix(Row1, Grid1.Col), "dd/mmm/yyyy hh:mm:ss")
100       d2 = Format(Grid1.TextMatrix(Row2, Grid1.Col), "dd/mmm/yyyy hh:mm:ss")
110       Cmp = Sgn(DateDiff("s", d1, d2))
     
End Sub



Private Sub txtUnitNumber_Change()

10    lantigen = ""
20    lblExpiry = ""
30    lgroup = ""
40    lProduct = ""
50    lsupplier = ""
60    Grid1.Rows = 2
70    Grid1.AddItem ""
80    Grid1.RemoveItem 1

End Sub

Private Sub txtUnitNumber_LostFocus()
Dim s As String

10    If Trim$(txtUnitNumber) = "" Then Exit Sub

20      txtUnitNumber = UCase(txtUnitNumber)
30      If Left$(txtUnitNumber, 1) = "=" Then 'Barcode scanning entry
40        s = ISOmod37_2(Mid$(txtUnitNumber, 2, 13))
50        txtUnitNumber = Mid$(txtUnitNumber, 2, 13) & " " & s
60      End If


End Sub


Public Property Let ProductName(ByVal Value As String)

10    mProduct = Value
20    lProduct = Value

End Property


Public Property Let Expiry(ByVal Expiry As String)

10      lblExpiry = Expiry

End Property
