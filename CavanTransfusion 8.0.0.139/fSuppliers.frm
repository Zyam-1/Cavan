VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fSuppliers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Suppliers"
   ClientHeight    =   6060
   ClientLeft      =   360
   ClientTop       =   945
   ClientWidth     =   11370
   ControlBox      =   0   'False
   Icon            =   "fSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   9660
      TabIndex        =   22
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   8925
      Begin VB.TextBox tFAX 
         Height          =   285
         Left            =   7200
         TabIndex        =   13
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox teMail 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   1410
         Width           =   3945
      End
      Begin VB.TextBox tListOrder 
         Height          =   285
         Left            =   8040
         TabIndex        =   11
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox tAddr3 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   1110
         Width           =   3945
      End
      Begin VB.TextBox tAddr2 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   810
         Width           =   3945
      End
      Begin VB.TextBox tPhone 
         Height          =   285
         Left            =   7200
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox tBarCode 
         Height          =   285
         Left            =   7200
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox tCode 
         Height          =   285
         Left            =   5460
         TabIndex        =   6
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox tAddr1 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   510
         Width           =   3945
      End
      Begin VB.TextBox tSupplier 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   210
         Width           =   3945
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   5160
         TabIndex        =   3
         Top             =   1350
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "List Order"
         Height          =   195
         Left            =   7320
         TabIndex        =   21
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "eMail"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   6810
         TabIndex        =   19
         Top             =   810
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         Height          =   195
         Left            =   6660
         TabIndex        =   18
         Top             =   510
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bar Code"
         Height          =   195
         Left            =   6480
         TabIndex        =   16
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3585
      Left            =   120
      TabIndex        =   1
      Top             =   2010
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      ForeColorFixed  =   16711680
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      GridLines       =   2
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   9660
      TabIndex        =   0
      Top             =   870
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   23
      Top             =   5730
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "fSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()
  
      Dim s As String
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    g.Visible = False

60    sql = "Select * from Supplier order by ListOrder"
70    Set tb = New Recordset
80    RecOpenServerBB 0, tb, sql
90    Do While Not tb.EOF
100     s = tb!Supplier & vbTab & _
            tb!BarCode & vbTab & _
            tb!code & vbTab & _
            tb!Addr1 & vbTab & _
            tb!Addr2 & vbTab & _
            tb!Addr3 & vbTab & _
            tb!Phone & vbTab & _
            tb!FAX & vbTab & _
            tb!eMail & vbTab & _
            tb!ListOrder
110     g.AddItem s
120     tb.MoveNext
130   Loop

140   g.Visible = True
150   If g.Rows > 2 Then
160     g.RemoveItem 1
170   End If

180   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "fSuppliers", "FillG", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdadd_Click()
  
      Dim s As String


10    s = tSupplier & vbTab & _
          tBarCode & vbTab & _
          tCode & vbTab & _
          taddr1 & vbTab & _
          tAddr2 & vbTab & _
          tAddr3 & vbTab & _
          tPhone & vbTab & _
          tFAX & vbTab & _
          teMail & vbTab & _
          tListOrder
20    g.AddItem s

30    tSupplier = ""
40    tBarCode = ""
50    tCode = ""
60    taddr1 = ""
70    tAddr2 = ""
80    tAddr3 = ""
90    tPhone = ""
100   tFAX = ""
110   teMail = ""
120   tListOrder = ""

130   tSupplier.SetFocus

End Sub

Private Sub cmdSave_Click()

      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmdSave_Click_Error

20    For n = 1 To g.Rows - 1
30      If g.TextMatrix(n, 0) <> "" Then
40        sql = "Select * from Supplier where " & _
              "BarCode = '" & g.TextMatrix(n, 1) & "'"
50        Set tb = New Recordset
60        RecOpenServerBB 0, tb, sql
70        If tb.EOF Then
80          tb.AddNew
90        End If
100       tb!Supplier = g.TextMatrix(n, 0)
110       tb!BarCode = g.TextMatrix(n, 1)
120       tb!code = g.TextMatrix(n, 2)
130       tb!Addr1 = g.TextMatrix(n, 3)
140       tb!Addr2 = g.TextMatrix(n, 4)
150       tb!Addr3 = g.TextMatrix(n, 5)
160       tb!Phone = g.TextMatrix(n, 6)
170       tb!FAX = g.TextMatrix(n, 7)
180       tb!eMail = g.TextMatrix(n, 8)
190       tb!ListOrder = n
200       tb.Update
210     End If
220   Next

230   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "fSuppliers", "cmdSave_Click", intEL, strES, sql


End Sub




Private Sub Form_Load()

      Dim n As Integer

10    g.Row = 0
20    For n = 0 To 9
30      g.Col = n
40      g = Choose(n + 1, "Supplier", "BarCode", _
                          "Code", "Addr1", "Addr2", _
                          "Addr3", "Phone", "FAX", _
                          "eMail", "List Order")
50      g.ColWidth(n) = Choose(n + 1, 2600, 900, _
                          500, 1000, 1000, _
                          1000, 900, 900, _
                          1000, 850)
60      g.ColAlignment(n) = 1
  
  
70    Next

      '*****NOTE
          'FillG might be dependent on many components so for any future
          'update in code try to keep FillG on bottom most line of form load.
80        FillG
      '**************************************

End Sub


Private Sub g_Click()

      Static SortOrder As Boolean

10    If g.MouseRow = 0 Then
20      If SortOrder Then
30        g.Sort = flexSortGenericAscending
40      Else
50        g.Sort = flexSortGenericDescending
60      End If
70      SortOrder = Not SortOrder
80      Exit Sub
90    End If

100   g.Col = 0
110   tSupplier = g
120   g.Col = 1
130   tBarCode = g
140   g.Col = 2
150   tCode = g
160   g.Col = 3
170   taddr1 = g
180   g.Col = 4
190   tAddr2 = g
200   g.Col = 5
210   tAddr3 = g
220   g.Col = 6
230   tPhone = g
240   g.Col = 7
250   tFAX = g
260   g.Col = 8
270   teMail = g
280   g.Col = 9
290   tListOrder = g

300   g.RemoveItem g.Row

End Sub

Private Sub tBarCode_Change()

10    If Trim$(tSupplier) <> "" And Trim$(tBarCode) <> "" Then
20      cmdAdd.Visible = True
30    Else
40      cmdAdd.Visible = False
50    End If

End Sub

Private Sub tSupplier_Change()

10    If Trim$(tSupplier) <> "" And Trim$(tBarCode) <> "" Then
20      cmdAdd.Visible = True
30    Else
40      cmdAdd.Visible = False
50    End If

End Sub


