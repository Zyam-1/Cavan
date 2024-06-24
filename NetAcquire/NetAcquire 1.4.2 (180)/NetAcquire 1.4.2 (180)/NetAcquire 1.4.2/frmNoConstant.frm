VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNoConstant 
   Caption         =   "NetAcquire"
   ClientHeight    =   3570
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   11505
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3105
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   5477
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmNoConstant.frx":0000
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   825
      Left            =   10500
      Picture         =   "frmNoConstant.frx":00B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   10530
      Picture         =   "frmNoConstant.frx":071C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2490
      Width           =   795
   End
End
Attribute VB_Name = "frmNoConstant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

Dim fso As New FileSystemObject
Dim sysFold As Folder
Dim tStream As TextStream
Dim ip As String
Dim AllIPs() As String
Dim CurrentIP() As String
Dim Active As Boolean
Dim n As Integer
Dim s As String

On Error GoTo ehFG

g.Rows = 2
g.AddItem ""
g.RemoveItem 1

For n = 0 To UBound(NoConstant)
  s = NoConstant(n).Active & vbTab & _
      NoConstant(n).HospName & vbTab & _
      NoConstant(n).DSN & vbTab & _
      NoConstant(n).DSNBB & vbTab & _
      NoConstant(n).DSNTest & vbTab & _
      NoConstant(n).DSNTestBB & vbTab & _
      NoConstant(n).USR & vbTab & _
      NoConstant(n).USRBB & vbTab & _
      NoConstant(n).USRTest & vbTab & _
      NoConstant(n).USRTestBB & vbTab & _
      NoConstant(n).PWD & vbTab & _
      NoConstant(n).PWDBB & vbTab & _
      NoConstant(n).PWDTest & vbTab & _
      NoConstant(n).PWDTestBB & vbTab & _
      NoConstant(n).HospitalGroup
  g.AddItem s
Next

If g.Rows > 2 Then
  g.RemoveItem 1
  If UBound(NoConstant) > 0 Then
    g.AddItem ""
  End If
End If

Exit Sub

ehFG:

Dim er As Long
Dim es As String

er = Err.Number
es = Err.Description

If er = 9 Then 'Subscript out of range
  Exit Sub
End If

End Sub

Private Sub cmdSave_Click()

Dim fso As New FileSystemObject
Dim sysFold As Folder
Dim tStream As TextStream
Dim s As String
Dim X As Integer
Dim n As Integer

Dim DSN As String
Dim USR As String
Dim PWD As String

'On Error GoTo ehFG

If g.Rows = 2 Then
  g.TextMatrix(1, 0) = "Yes"
End If

Set sysFold = fso.GetSpecialFolder(SystemFolder)
fso.DeleteFile sysFold & "\ZWinNet.bin"
Set tStream = fso.OpenTextFile(sysFold & "\ZWinNet.BIN", ForAppending, True)

s = ""
For n = 1 To g.Rows - 1
  If Trim$(g.TextMatrix(n, 1)) <> "" Then
    For X = 0 To g.Cols - 1
      s = s & g.TextMatrix(n, X) & "|"
    Next
    s = Left$(s, Len(s) - 1) 'remove final "|"
    s = s & vbCr
  End If
Next
s = Left$(s, Len(s) - 1) 'remove final vbCr

s = Obfuscate(s)

tStream.Write s
tStream.Close

GetConnectInfo
FillG

End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub


Private Sub Form_Load()

FillG

End Sub

Private Sub g_Click()

Dim Entries As Integer
Dim n As Integer
Dim s As String

If g.Rows = 2 Then 'The grid is empty
                   'this can happen if sysFold\ZWinNet.BIN does not exist or is blank
  Entries = 0
Else
  Entries = g.Rows - 2
End If

Select Case g.Col
  Case 0:
    If Entries > 1 Then
      If g.TextMatrix(g.Row, 0) = "Yes" Then
        Exit Sub
      ElseIf g.Row < g.Rows - 1 Then
        For n = 1 To g.Rows - 2
          g.TextMatrix(n, 0) = "No"
        Next
        g.TextMatrix(g.Row, 0) = "Yes"
      End If
    End If
  Case 1:
    If Entries = 0 Then
      s = InputBox("Hospital Name")
      If InStr(s, "|") <> 0 Then
        iMsg "Bar character <|> is not allowed!", vbExclamation
      Else
        g.TextMatrix(g.Row, 1) = s
      End If
    Else
      s = InputBox(g.TextMatrix(0, 1))
      If InStr(s, "|") <> 0 Then
        iMsg "Bar character <|> is not allowed!", vbExclamation
      Else
        g.TextMatrix(g.Row, 1) = s
      End If
    End If
  Case Else:
    s = InputBox(g.TextMatrix(0, g.Col))
    If InStr(s, "|") <> 0 Then
      iMsg "Bar character <|> is not allowed!", vbExclamation
    Else
      g.TextMatrix(g.Row, g.Col) = s
    End If
End Select

End Sub


