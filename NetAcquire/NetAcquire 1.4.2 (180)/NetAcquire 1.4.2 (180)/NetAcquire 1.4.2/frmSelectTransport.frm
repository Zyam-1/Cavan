VERSION 5.00
Begin VB.Form frmTrackSelectTransport 
   Caption         =   "NetAcquire - Select Preferred Transport"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtImportPath 
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
      Left            =   4230
      TabIndex        =   11
      Text            =   "E:\Transport\"
      Top             =   4500
      Width           =   4185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   765
      Left            =   8640
      TabIndex        =   0
      Top             =   4020
      Width           =   1035
   End
   Begin VB.TextBox txtFOBPath 
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
      Left            =   4230
      TabIndex        =   6
      Text            =   "E:\Transport\"
      Top             =   3180
      Width           =   4185
   End
   Begin VB.TextBox txtFTPPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   4
      Text            =   "\\192.168.0.33\Transport\"
      Top             =   735
      Width           =   4155
   End
   Begin VB.TextBox txtCDPath 
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
      Left            =   4230
      TabIndex        =   5
      Text            =   "D:\Transport\"
      Top             =   1950
      Width           =   4185
   End
   Begin VB.CommandButton cmdSelectFOB 
      Caption         =   "Use Flash Drive"
      Height          =   1185
      Left            =   360
      Picture         =   "frmSelectTransport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2745
   End
   Begin VB.CommandButton cmdSelectFTP 
      Caption         =   "Use FTP"
      Height          =   1185
      Left            =   360
      Picture         =   "frmSelectTransport.frx":0905
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   300
      Width           =   2745
   End
   Begin VB.CommandButton cmdSelectCD 
      Caption         =   "Use CD"
      Height          =   1185
      Left            =   360
      Picture         =   "frmSelectTransport.frx":116B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1530
      Width           =   2745
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Import Path"
      Height          =   195
      Left            =   2055
      TabIndex        =   10
      Top             =   4560
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Path"
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   3210
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Path"
      Height          =   195
      Left            =   3240
      TabIndex        =   8
      Top             =   795
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Path"
      Height          =   225
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Top             =   1980
      Width           =   945
   End
End
Attribute VB_Name = "frmTrackSelectTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FillDetails()

Dim sql As String
Dim tb As Recordset

cmdSelectFTP.Caption = "Send by FTP"
cmdSelectCD.Caption = "Write to CD"
cmdSelectFOB.Caption = "Write to Key FOB"

sql = "Select * from Options where " & _
      "Description = 'TransportPreferred'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  Select Case UCase$(tb!Contents & "")
    Case "FTP": cmdSelectFTP.Caption = cmdSelectFTP.Caption & " (Active)"
    Case "CD": cmdSelectCD.Caption = cmdSelectCD.Caption & " (Active)"
    Case "FOB": cmdSelectFOB.Caption = cmdSelectFOB.Caption & " (Active)"
  End Select
End If

sql = "Select * from Options where " & _
      "Description = 'TransportFTP'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtFTPPath = tb!Contents & ""
Else
  txtFTPPath = ""
End If

sql = "Select * from Options where " & _
      "Description = 'TransportCD'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtCDPath = tb!Contents & ""
Else
  txtCDPath = ""
End If

sql = "Select * from Options where " & _
      "Description = 'TransportFOB'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtFOBPath = tb!Contents & ""
Else
  txtFOBPath = ""
End If

sql = "Select * from Options where " & _
      "Description = 'TransportImport'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  txtImportPath = tb!Contents & ""
Else
  txtImportPath = ""
End If

End Sub

Private Sub SelectPreferred(ByVal s As String)

Dim sql As String
Dim tb As Recordset

sql = "Select * from Options where " & _
      "Description = 'TransportPreferred'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportPreferred"
End If
tb!Contents = s
tb.Update

sql = "Select * from Options where " & _
      "Description = 'TransportFOB'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportFOB"
End If
tb!Contents = txtFOBPath
tb.Update

FillDetails

End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub


Private Sub cmdSelectCD_Click()

SelectPreferred "CD"

End Sub

Private Sub cmdSelectFOB_Click()

SelectPreferred "FOB"

End Sub


Private Sub cmdSelectFTP_Click()

SelectPreferred "FTP"

End Sub


Private Sub Form_Load()

FillDetails

End Sub

Private Sub txtCDPath_LostFocus()

Dim sql As String
Dim tb As Recordset

If Right$(txtCDPath, 1) <> "\" Then
  txtCDPath = txtCDPath & "\"
End If

sql = "Select * from Options where " & _
      "Description = 'TransportCD'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportCD"
End If
tb!Contents = txtCDPath
tb.Update

End Sub


Private Sub txtFOBPath_LostFocus()

Dim sql As String
Dim tb As Recordset

If Right$(txtFOBPath, 1) <> "\" Then
  txtFOBPath = txtFOBPath & "\"
End If

sql = "Select * from Options where " & _
      "Description = 'TransportFOB'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportFOB"
End If
tb!Contents = txtFOBPath
tb.Update

End Sub


Private Sub txtFTPPath_LostFocus()

Dim tb As Recordset
Dim sql As String

If Right$(txtFTPPath, 1) <> "\" Then
  txtFTPPath = txtFTPPath & "\"
End If

sql = "Select * from Options where " & _
      "Description = 'TransportFTP'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportFTP"
End If
tb!Contents = txtFTPPath
tb.Update

End Sub


Private Sub txtImportPath_Change()

End Sub


Private Sub txtImportPath_LostFocus()

Dim sql As String
Dim tb As Recordset

If Right$(txtImportPath, 1) <> "\" Then
  txtImportPath = txtImportPath & "\"
End If

sql = "Select * from Options where " & _
      "Description = 'TransportImport'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
  tb!Description = "TransportImport"
End If
tb!Contents = txtImportPath
tb.Update

End Sub


