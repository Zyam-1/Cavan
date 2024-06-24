VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAntiDUsage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anti-D Usage"
   ClientHeight    =   4035
   ClientLeft      =   1545
   ClientTop       =   1605
   ClientWidth     =   6990
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
   Icon            =   "frmAntiDUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   6990
   Begin VB.CommandButton bstart 
      Appearance      =   0  'Flat
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2010
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
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
      Height          =   495
      Left            =   4770
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
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
      Height          =   495
      Left            =   3390
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   2130
      TabIndex        =   18
      Top             =   390
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   3660
      TabIndex        =   19
      Top             =   390
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   92536833
      CurrentDate     =   36963
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   870
      TabIndex        =   21
      Top             =   3840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Between Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2130
      TabIndex        =   20
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label laadpp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   8
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Average Anti-D per patient:"
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
      Left            =   1395
      TabIndex        =   17
      Top             =   1740
      Width           =   1920
   End
   Begin VB.Label lnoanp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   15
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number of Ante Natal Patients:"
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
      TabIndex        =   14
      Top             =   1020
      Width           =   2190
   End
   Begin VB.Label lpnrgtd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4770
      TabIndex        =   11
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label lpnrtd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4770
      TabIndex        =   10
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label lpnrod 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4770
      TabIndex        =   9
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label lnrgtd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   7
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label lnrtd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   6
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label lnrod 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   5
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label ltdoadu 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   4
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number receiving > two doses:"
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
      Index           =   0
      Left            =   1125
      TabIndex        =   3
      Top             =   2820
      Width           =   2190
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number receiving two doses:"
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
      Index           =   0
      Left            =   1260
      TabIndex        =   2
      Top             =   2460
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number receiving one dose:"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   2100
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total doses of Anti-D used:"
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
      Index           =   0
      Left            =   1395
      TabIndex        =   0
      Top             =   1380
      Width           =   1920
   End
End
Attribute VB_Name = "frmAntiDUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub bprint_Click()

      Dim Px As Printer
      Dim OriginalPrinter As String

10    OriginalPrinter = Printer.DeviceName
20    If Not SetFormPrinter() Then Exit Sub

30    Printer.Print
40    Printer.FontBold = True
50    Printer.Print "Anti-D Immunoglobulin Usage"
60    Printer.FontBold = False
70    Printer.Print
80    Printer.Print "         Number of Ante-Natal Patients: "; lnoanp
90    Printer.Print "           Total doses of Anti-D given: "; ltdoadu
100   Printer.Print " Number of Patients receiving one dose: "; lnrod; "("; lpnrod; ")"
110   Printer.Print "Number of Patients receiving two doses: "; lnrtd; "("; lpnrtd; ")"
120   Printer.Print "  Number receiving more than two doses: "; lnrgtd; "("; lpnrgtd; ")"
130   Printer.Print
140   Printer.Print "Average Anti-D per Patient: "; laadpp
150   Printer.Print ""

160   Printer.EndDoc

170   For Each Px In Printers
180     If Px.DeviceName = OriginalPrinter Then
190       Set Printer = Px
200       Exit For
210     End If
220   Next

End Sub

Private Sub bstart_Click()


      Dim sn As Recordset
      Dim sql As String
      Dim strFromTime As String
      Dim strToTime As String
      Dim s As String
      Dim adused As Integer
      Dim totalused As Integer
      Dim used1 As Integer
      Dim used2 As Integer
      Dim used3 As Integer

10    On Error GoTo bstart_Click_Error

20    strFromTime = Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00"
30    strToTime = Format(dtTo, "dd/mmm/yyyy") & " 23:59:59"

40    sql = "select given from anti_d where " & _
            "datetime between '" & _
            strFromTime & "' and '" & strToTime & "'"

50    Set sn = New Recordset
60    RecOpenServerBB 0, sn, sql
70    If sn.EOF Then
80      s = "No Record of any Ante Natal tests" & Chr(10)
90      s = s & "between those two dates."
100     iMsg s, vbInformation
110     If TimedOut Then Unload Me: Exit Sub
120     Exit Sub
130   End If

140   sn.MoveLast
150   lnoanp = Format(sn.RecordCount)
160   sn.MoveFirst
170   totalused = 0: used1 = 0: used2 = 0: used3 = 0
180   Do While Not sn.EOF
190     If Not IsNull(sn!given) Then
200       adused = countof(sn!given)
210     End If
220     totalused = totalused + adused
230     Select Case adused
          Case 1: used1 = used1 + 1
240       Case 2: used2 = used2 + 1
250       Case Is > 2: used3 = used3 + 1
260     End Select
270     sn.MoveNext
280   Loop

290   ltdoadu.Caption = Format(totalused)
300   lnrod.Caption = Format(used1)
310   lnrtd.Caption = Format(used2)
320   lnrgtd.Caption = Format(used3)

330   laadpp.Caption = Format(Val(ltdoadu) / Val(lnoanp), "#.##")
340   lpnrod.Caption = Format(Val(lnrod) / Val(lnoanp), "##.#%")
350   lpnrtd.Caption = Format(Val(lnrtd) / Val(lnoanp), "##.#%")
360   lpnrgtd.Caption = Format(Val(lnrgtd) / Val(lnoanp), "##.#%")

370   Exit Sub

bstart_Click_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmAntiDUsage", "bstart_Click", intEL, strES, sql

End Sub

Private Function countof(ByVal s As String) As Integer
      'returns number of words in s by
      'counting number of spaces in s + 1

      Dim n As Integer
      Dim C As Integer

10    C = 0
20    s = Trim$(s)
30    If s <> "" Then
40      For n = 1 To Len(s)
50        If Mid$(s, n, 1) = " " Then C = C + 1
60      Next
70    C = C + 1
80    End If

90    countof = C

End Function

Private Sub Form_Load()

10    dtTo = Format(Now, "dd/mmm/yyyy")
20    dtFrom = Format(Now - 7, "dd/mmm/yyyy")

End Sub


