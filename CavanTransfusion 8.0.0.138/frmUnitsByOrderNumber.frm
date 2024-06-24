VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnitsByOrderNumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Search Units By Order Number"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrderNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1305
      Width           =   2475
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1125
      Left            =   7770
      Picture         =   "frmUnitsByOrderNumber.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   690
      Width           =   1125
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   1125
      Left            =   2700
      Picture         =   "frmUnitsByOrderNumber.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1125
      Left            =   10680
      Picture         =   "frmUnitsByOrderNumber.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   690
      Width           =   1125
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export"
      Height          =   1125
      Left            =   6510
      Picture         =   "frmUnitsByOrderNumber.frx":2C5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3585
      Left            =   120
      TabIndex        =   5
      Top             =   1890
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   5550
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Order Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Units For Selected Order Number"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   270
      Left            =   3120
      TabIndex        =   8
      Top             =   225
      Width           =   5250
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8970
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   135
      Top             =   150
      Width           =   11640
   End
End
Attribute VB_Name = "frmUnitsByOrderNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SortOrder As Boolean


Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub cmdPrint_Click()

10    If Grid.Rows = 1 Then
20        iMsg "Nothing to print", vbInformation
30        If TimedOut Then Unload Me: Exit Sub
40        Exit Sub
50    End If

      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

60    OriginalPrinter = Printer.DeviceName

70    If Not SetFormPrinter() Then Exit Sub

80    Printer.Print
90    Printer.FontName = "Courier New"
100   Printer.FontSize = 10
110   Printer.Font.Bold = True

      '****Report heading

120   Printer.Print FormatString(Label1.Caption, 140, , AlignCenter)
130   Printer.Print FormatString("Unit Number: " & txtOrderNumber, 140, AlignCenter)


140   Printer.Font.Size = 9

      '****Report body
150   For i = 1 To 152
160       Printer.Print "-";
170   Next i
180   Printer.Print

190   Printer.Print FormatString("", 0, "|");
200   Printer.Print FormatString("Unit", 16, "|", AlignCenter);
210   Printer.Print FormatString("DateTime", 20, "|", AlignCenter);
220   Printer.Print FormatString("Product", 35, "|", AlignCenter);
230   Printer.Print FormatString("Expiry", 10, "|", AlignCenter);
240   Printer.Print FormatString("Group", 6, "|", AlignCenter);
250   Printer.Print FormatString("Supplier", 25, "|", AlignCenter);

260   For i = 1 To 152
270       Printer.Print "-";
280   Next i
290   Printer.Print
 
300   Printer.Font.Bold = False
  
310   For Y = 1 To Grid.Rows - 1
320       Printer.Print FormatString("", 0, "|");
330       Printer.Print FormatString(Grid.TextMatrix(Y, 0), 16, "|", AlignCenter);
340       Printer.Print FormatString(Grid.TextMatrix(Y, 1), 20, "|", AlignLeft);
350       Printer.Print FormatString(Grid.TextMatrix(Y, 2), 35, "|", AlignLeft);
360       Printer.Print FormatString(Grid.TextMatrix(Y, 3), 10, "|", AlignLeft);
370       Printer.Print FormatString(Grid.TextMatrix(Y, 4), 6, "|", AlignCenter);
380       Printer.Print FormatString(Grid.TextMatrix(Y, 5), 25, "|", AlignLeft)
    
390   Next

400   Printer.EndDoc

410   For Each Px In Printers
420     If Px.DeviceName = OriginalPrinter Then
430       Set Printer = Px
440       Exit For
450     End If
460   Next

End Sub

Private Sub cmdSearch_Click()

10    If txtOrderNumber = "" Then
20        iMsg "Please enter order number first", vbInformation
30        If TimedOut Then Unload Me: Exit Sub
40        txtOrderNumber.SetFocus
50        Exit Sub
60    End If

70    FillG

End Sub

Private Sub cmdXL_Click()

10    If Grid.Rows = 1 Then
20        iMsg "Nothing to export", vbInformation
30        If TimedOut Then Unload Me: Exit Sub
40        Exit Sub
50    End If
      Dim strHeading As String

60    strHeading = Label1 & vbCr
70    strHeading = strHeading & "Order Number: " & txtOrderNumber & vbCr

80    strHeading = strHeading & " " & vbCr
90    ExportFlexGrid Grid, Me, strHeading

End Sub



Private Sub Grid_Click()

10    If Grid.MouseRow = 0 Then
20      If InStr(Grid.TextMatrix(0, Grid.Col), "Date") <> 0 Then
30        Grid.Sort = 9
40      Else
50        If SortOrder Then
60          Grid.Sort = flexSortGenericAscending
70        Else
80          Grid.Sort = flexSortGenericDescending
90        End If
100     End If
110     SortOrder = Not SortOrder
120     Exit Sub
130   End If


End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
       Dim d1 As String
      Dim d2 As String

10    If Not IsDate(Grid.TextMatrix(Row1, Grid.Col)) Then
20      Cmp = 0
30      Exit Sub
40    End If

50    If Not IsDate(Grid.TextMatrix(Row2, Grid.Col)) Then
60      Cmp = 0
70      Exit Sub
80    End If

90    d1 = Format(Grid.TextMatrix(Row1, Grid.Col), "dd/mmm/yyyy hh:mm:ss")
100   d2 = Format(Grid.TextMatrix(Row2, Grid.Col), "dd/mmm/yyyy hh:mm:ss")

110   If SortOrder Then
120     Cmp = Sgn(DateDiff("D", d1, d2))
130   Else
140     Cmp = Sgn(DateDiff("D", d2, d1))
150   End If
End Sub

Private Sub InitGrid()

10    With Grid
20        .Rows = 2: .FixedRows = 1
30        .Cols = 6: .FixedCols = 0
40        .Rows = 1
          '.AddItem ""
          '.RemoveItem 1
          '^Unit Number  |^DateTime   |^Product|^Expiry|^Group|^Supplier|^Screen|
50        .ColWidth(0) = 1500: .TextMatrix(0, 0) = "Unit #": .ColAlignment(0) = flexAlignCenterCenter
60        .ColWidth(1) = 1900: .TextMatrix(0, 1) = "DateTime": .ColAlignment(1) = flexAlignLeftCenter
70        .ColWidth(2) = 4000: .TextMatrix(0, 2) = "Product": .ColAlignment(2) = flexAlignLeftCenter
80        .ColWidth(3) = 1400: .TextMatrix(0, 3) = "Expiry Date": .ColAlignment(3) = flexAlignLeftCenter
90        .ColWidth(4) = 800: .TextMatrix(0, 4) = "Group": .ColAlignment(4) = flexAlignCenterCenter
100       .ColWidth(5) = 2000: .TextMatrix(0, 5) = "Supplier": .ColAlignment(5) = flexAlignLeftCenter
110   End With

End Sub


Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    InitGrid

30    sql = "Select * From Product Where OrderNumber = '" & txtOrderNumber & "' And Event = 'C' Order By [DateTime]"

40    Set tb = New Recordset
50    RecOpenServerBB 0, tb, sql
  
60    If tb.EOF Then
70        iMsg "No Records found.", vbExclamation
80        If TimedOut Then Unload Me: Exit Sub
90        Exit Sub
100   Else
110       Grid.Visible = False
120       Do While Not tb.EOF
130           s = tb!ISBT128 & "" & vbTab & _
                  tb!DateTime & "" & vbTab & _
                  ProductWordingFor(tb!BarCode & "") & vbTab & _
                  Format(tb!DateExpiry, "dd/mm/yyyy HH:mm") & "" & vbTab & _
                  Bar2Group(tb!GroupRh & "") & vbTab & _
                  SupplierCodeFor(tb!Supplier & "")

140           Grid.AddItem s
150           tb.MoveNext
160       Loop
    
170       tb.Close
180   End If
190   Grid.Visible = True

200   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

210   Grid.Visible = True
220   intEL = Erl
230   strES = Err.Description
240   LogError "frmSearch", "FillG", intEL, strES, sql
250   Grid.Visible = True

End Sub

