VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaily 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "NetAcquire - Daily Report"
   ClientHeight    =   7650
   ClientLeft      =   810
   ClientTop       =   1395
   ClientWidth     =   10485
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
   ScaleHeight     =   7650
   ScaleWidth      =   10485
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   270
      TabIndex        =   3
      Top             =   390
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   218955777
      CurrentDate     =   37096
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6615
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   6
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
      FormatString    =   $"frmDaily.frx":0000
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
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
      Height          =   525
      Left            =   4890
      TabIndex        =   1
      Top             =   300
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
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
      Height          =   525
      Left            =   8700
      TabIndex        =   0
      Top             =   300
      Width           =   1245
   End
End
Attribute VB_Name = "frmDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

30380     Unload Me

End Sub

Private Sub bPrint_Click()

          Dim Y As Integer
          Dim fs As Integer

30390     fs = Printer.FontSize

30400     Printer.FontSize = 12
30410     Printer.Print
30420     Printer.Print

30430     Printer.Print "Daily Report for "; Format$(dt, "dd/mmm/yyyy")
30440     Printer.Print

30450     Printer.FontSize = 8

30460     For Y = 0 To g.Rows - 1
30470         g.row = Y
30480         g.Col = 0
30490         Printer.Print g;
30500         g.Col = 1 'name
30510         Printer.Print Tab(7); Left$(g, 17);
30520         g.Col = 2 'chart
30530         Printer.Print Tab(25); g;
30540         g.Col = 3 'gp
30550         Printer.Print Tab(31); Left$(g, 17);
30560         g.Col = 4 'ward
30570         Printer.Print Tab(49); Left$(g, 17);
30580         g.Col = 5 'clinician
30590         Printer.Print Tab(66); g
30600     Next

30610     Printer.FontSize = fs

30620     Printer.EndDoc

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim s As String
          Dim sql As String

30630     On Error GoTo FillG_Error

30640     g.Rows = 2
30650     g.AddItem ""
30660     g.RemoveItem 1

30670     sql = "select * from demographics where " & _
              "rundate = '" & Format$(dt, "dd/mmm/yyyy") & "' " & _
              "order by sampleid"

30680     Set tb = New Recordset
30690     RecOpenClient 0, tb, sql
30700     Do While Not tb.EOF
30710         s = tb!SampleID & vbTab & _
                  tb!PatName & vbTab & _
                  tb!Chart & vbTab & _
                  tb!GP & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & ""
30720         g.AddItem s
30730         tb.MoveNext
30740     Loop

30750     If g.Rows > 2 Then g.RemoveItem 1

30760     Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

30770     intEL = Erl
30780     strES = Err.Description
30790     LogError "fdaily", "FillG", intEL, strES, sql

End Sub

Private Sub dt_CloseUp()

30800     FillG

End Sub


Private Sub Form_Activate()

30810     FillG

End Sub

Private Sub Form_Load()

30820     dt = Format$(Now, "dd/mm/yyyy")

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean

30830     If g.MouseRow = 0 Then
30840         If SortOrder Then
30850             g.Sort = flexSortGenericAscending
30860         Else
30870             g.Sort = flexSortGenericDescending
30880         End If
30890         SortOrder = Not SortOrder
30900     End If

End Sub


