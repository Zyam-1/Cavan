VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmImportDiamed 
   Caption         =   "NetAcquire - Import Diamed Panel"
   ClientHeight    =   2865
   ClientLeft      =   2760
   ClientTop       =   2715
   ClientWidth     =   5475
   Icon            =   "frmImportDiamed.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   765
      Left            =   240
      Picture         =   "frmImportDiamed.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2730
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tLotNumber 
      Height          =   285
      Left            =   3690
      MaxLength       =   10
      TabIndex        =   3
      Top             =   360
      Width           =   1635
   End
   Begin VB.CommandButton bImport 
      Caption         =   "Import Now"
      Height          =   765
      Left            =   4080
      Picture         =   "frmImportDiamed.frx":0F34
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtIssued 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   210
      TabIndex        =   10
      Top             =   2640
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lFile 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1110
      Width           =   5085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File to Import"
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Expiry Date"
      Height          =   195
      Index           =   1
      Left            =   1620
      TabIndex        =   6
      Top             =   150
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Issued Date"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Panel No"
      Height          =   195
      Index           =   0
      Left            =   3750
      TabIndex        =   4
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "frmImportDiamed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bImport_Click()

      Dim f As Integer
      Dim IP As String
      Dim X As Integer
      Dim Y As Integer
      Dim Heading As String
      Dim p As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim Pattern As String

10    On Error GoTo bImport_Click_Error

20    If Trim$(tLotNumber) = "" Then
30      iMsg "Enter Panel Number", vbCritical
40      If TimedOut Then Unload Me: Exit Sub
50    End If

60    f = FreeFile

70    Open lFile For Input As f
80    Line Input #f, IP
90    Close f
    
100   sql = "Select * from AntibodyPanels where " & _
            "LotNumber = '" & tLotNumber & "'"
110   Set tb = New Recordset
120   RecOpenServerBB 0, tb, sql
130   If tb.EOF Then
140     tb.AddNew
150   End If
160   tb!LotNumber = tLotNumber
170   tb!IssuedDate = Format(dtIssued, "dd/mmm/yyyy")
180   tb!ExpiryDate = Format(dtExpiry, "dd/mmm/yyyy")
190   tb!Supplier = "DiaMed"
200   tb!DateEntered = Format(Now, "dd/mmm/yyyy")
210   tb!EnteredBy = UserName
220   tb.Update

230   For X = 0 To 25
240     Heading = Choose(X + 1, "D", "C", "E", "c", "e", "Cw", _
                                "K", "k", "Kpa", "Kpb", "Jsa", "Jsb", _
                                "Fya", "Fyb", _
                                "Jka", "Jkb", "Lea", "Leb", _
                                "P1", "M", "N", "S", "s", _
                                "Lua", "Lub", "Xga")

  
250     sql = "Select * from AntibodyPatterns where " & _
              "LotNumber = '" & tLotNumber & "' " & _
              "and Position = '" & X & "'"
260     Set tb = New Recordset
270     RecOpenServerBB 0, tb, sql
280     If tb.EOF Then
290       tb.AddNew
300     End If
310     tb!LotNumber = tLotNumber
320     tb!Position = X + 4
330     Pattern = Heading & vbTab
340     For Y = 1 To 11
350       p = (X * 11) + Y
360       Pattern = Pattern & IIf(Mid$(IP, p, 1) = "1", "+", "0") & vbTab
370     Next
380     If Right$(Pattern, 1) = vbTab Then
390       Pattern = Left$(Pattern, Len(Pattern) - 1)
400     End If
410     tb!Pattern = Pattern
420     tb.Update
  
430   Next

440   frmDefineABPanel.cmbLotNumber = tLotNumber

450   Unload Me

460   Exit Sub

bImport_Click_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmImportDiamed", "bImport_Click", intEL, strES, sql


End Sub

Private Sub lFile_Click()

10    CD.Flags = cdlOFNReadOnly
20    CD.InitDir = "c:\"
30    CD.ShowOpen

40    lFile = CD.FileName

End Sub


