VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExtAddress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6345
   ClientLeft      =   1950
   ClientTop       =   1980
   ClientWidth     =   7605
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   7605
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3735
      Left            =   240
      TabIndex        =   13
      Top             =   2430
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
   End
   Begin VB.TextBox tcode 
      Height          =   285
      Left            =   5580
      MaxLength       =   5
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox taddr 
      Height          =   285
      Index           =   3
      Left            =   930
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1140
      Width           =   3915
   End
   Begin VB.TextBox taddr 
      Height          =   285
      Index           =   1
      Left            =   930
      MaxLength       =   40
      TabIndex        =   1
      Top             =   540
      Width           =   3915
   End
   Begin VB.TextBox taddr 
      Height          =   285
      Index           =   0
      Left            =   930
      MaxLength       =   40
      TabIndex        =   0
      Top             =   240
      Width           =   3915
   End
   Begin VB.TextBox taddr 
      Height          =   285
      Index           =   2
      Left            =   930
      MaxLength       =   40
      TabIndex        =   2
      Top             =   840
      Width           =   3915
   End
   Begin VB.TextBox tfax 
      Height          =   285
      Left            =   930
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2010
      Width           =   2115
   End
   Begin VB.TextBox tphone 
      Height          =   285
      Left            =   930
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1590
      Width           =   2115
   End
   Begin VB.CommandButton badd 
      Appearance      =   0  'Flat
      Caption         =   "&Add"
      Height          =   525
      Left            =   3600
      TabIndex        =   7
      Top             =   1740
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   5280
      TabIndex        =   8
      Top             =   1740
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Code"
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
      Left            =   5100
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Address"
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
      Left            =   210
      TabIndex        =   10
      Top             =   270
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Phone"
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
      Left            =   330
      TabIndex        =   11
      Top             =   1620
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FAX"
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
      Left            =   510
      TabIndex        =   9
      Top             =   2040
      Width           =   300
   End
End
Attribute VB_Name = "frmExtAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bAdd_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim ans As Integer

31750     On Error GoTo bAdd_Click_Error

31760     If Trim(taddr(0)) = "" Then
31770         ans = iMsg("First line of Address must be filled.", 16, "Save Error")
31780         Exit Sub
31790     End If

31800     If Trim(tCode) = "" Then
31810         ans = iMsg("Code must be entered.", 16, "Save Error")
31820         Exit Sub
31830     End If

31840     sql = "Select * from eaddress where " & _
              "code = '" & tCode & "'"
31850     Set tb = New Recordset
31860     RecOpenServer 0, tb, sql
31870     If tb.EOF Then
31880         tb.AddNew
31890     Else
31900         ans = iMsg("This code already used. Edit this entry?", 32 + 4, "NetAcquire")
31910         If ans <> vbYes Then
31920             Exit Sub
31930         End If
31940     End If

31950     tb!Code = tCode
31960     tb!Addr0 = taddr(0)
31970     tb!Addr1 = taddr(1)
31980     tb!addr2 = taddr(2)
31990     tb!addr3 = taddr(3)
32000     tb!Phone = tphone
32010     tb!FAX = tfax

32020     tb.Update

32030     tCode = ""
32040     taddr(0) = ""
32050     taddr(1) = ""
32060     taddr(2) = ""
32070     taddr(3) = ""
32080     tphone = ""
32090     tfax = ""

32100     FillG

32110     Exit Sub

bAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

32120     intEL = Erl
32130     strES = Err.Description
32140     LogError "frmExtAddress", "bAdd_Click", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

32150     Unload Me

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

32160     On Error GoTo FillG_Error

32170     sql = "Select * from eaddress order by code"
32180     Set tb = New Recordset
32190     RecOpenServer 0, tb, sql

32200     g.Rows = 2

32210     Do While Not tb.EOF
32220         s = tb!Code & vbTab
32230         s = s & tb!Addr0 & vbTab
32240         s = s & tb!Addr1 & vbTab
32250         s = s & tb!addr2 & vbTab
32260         s = s & tb!addr3 & vbTab
32270         s = s & tb!Phone & vbTab
32280         s = s & tb!FAX & ""
32290         g.AddItem s
32300         tb.MoveNext
32310     Loop

32320     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

32330     intEL = Erl
32340     strES = Err.Description
32350     LogError "frmExtAddress", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim n As Integer

32360     g.row = 0
32370     For n = 0 To 6
32380         g.Col = n
32390         g.ColWidth(n) = Choose(n + 1, 800, 2685, 2300, 2000, 2000, 1000, 1000)
32400         g = Choose(n + 1, "Code", "Send to", "Address", "", "", "Phone", "FAX")
32410     Next

32420     FillG

End Sub

Private Sub g_Click()

32430     If g.row < 2 Then Exit Sub

32440     g.Col = 0: tCode = g
32450     g.Col = 1: taddr(0) = g
32460     g.Col = 2: taddr(1) = g
32470     g.Col = 3: taddr(2) = g
32480     g.Col = 4: taddr(3) = g
32490     g.Col = 5: tphone = g
32500     g.Col = 6: tfax = g

End Sub

