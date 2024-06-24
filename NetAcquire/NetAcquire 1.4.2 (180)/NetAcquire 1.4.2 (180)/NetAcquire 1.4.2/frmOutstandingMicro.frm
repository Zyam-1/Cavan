VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOutstandingMicro 
   Caption         =   "NetAcquire"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   1095
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   9120
      Picture         =   "frmOutstandingMicro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4365
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmOutstandingMicro.frx":066A
   End
End
Attribute VB_Name = "frmOutstandingMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

19860 Unload Me

End Sub

Private Sub Form_Activate()

19870 LoadDetails

End Sub

Private Sub LoadDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim s As String
      Dim blnSite As Boolean

19880 On Error GoTo LoadDetails_Error

19890 grd.Rows = 2
19900 grd.AddItem ""
19910 grd.RemoveItem 1

19920 sql = "Select M.*, D.PatName " & _
            "from MicroRequests as M left join Demographics as D " & _
            "on M.SampleID = D.SampleID " & _
            "Order by M.SampleID asc"
19930 Set tb = New Recordset
19940 RecOpenServer 0, tb, sql
19950 Do While Not tb.EOF

19960   s = tb!SampleID & vbTab & _
            tb!PatName & vbTab
        
19970   blnSite = False
        
19980   For n = 0 To 2
19990     If tb!Faecal And 2 ^ n Then
20000       If Not blnSite Then
20010         s = s & "Faecal "
20020         blnSite = True
20030       End If
20040       s = s & Choose(n + 1, "C & S ", "C. Difficile ", "O/P ")
20050     End If
20060   Next
        
20070   For n = 3 To 5
20080     If tb!Faecal And 2 ^ n Then
20090       If Not blnSite Then
20100         s = s & "Faecal "
20110         blnSite = True
20120       End If
20130       s = s & "Occult Blood "
20140       Exit For
20150     End If
20160   Next
          
20170   If tb!Faecal And 2 ^ 6 Then
20180     s = s & "Rota/Adeno "
20190   End If
        
20200   For n = 7 To 10
20210     If tb!Faecal And 2 ^ n Then
20220       s = s & Choose(n + 1, "Toxin A ", "Coli 0157 ", _
                                  "E/P Coli ", "S/S Screen ")
20230     End If
20240   Next
        
20250   For n = 0 To 5
20260     If tb!Urine And 2 ^ n Then
20270       If Not blnSite Then
20280         s = s & "Urine "
20290         blnSite = True
20300       End If
20310       s = s & Choose(n + 1, "C & S", "Pregnancy ", "Fat Globules ", _
                                  "Bence Jones ", "SG ", "HCG ")
20320     End If
20330   Next

20340   grd.AddItem s
20350   tb.MoveNext
        
20360 Loop

20370 If grd.Rows > 2 Then
20380   grd.RemoveItem 1
20390 End If

20400 Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

20410 intEL = Erl
20420 strES = Err.Description
20430 LogError "frmOutstandingMicro", "LoadDetails", intEL, strES, sql


End Sub


