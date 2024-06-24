VERSION 5.00
Begin VB.Form frmUpdatePrinted 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   4605
   ClientLeft      =   2340
   ClientTop       =   3525
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1155
      Left            =   4110
      Picture         =   "frmUpdatePrinted.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1245
   End
   Begin VB.CommandButton cmdRemoveHaemRequests 
      Caption         =   "Remove All Haematology Requests"
      Height          =   525
      Left            =   120
      TabIndex        =   5
      Top             =   2340
      Width           =   3735
   End
   Begin VB.CommandButton bRemoveCoagRequests 
      Caption         =   "Remove All Coagulation Requests"
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   3870
      Width           =   3735
   End
   Begin VB.CommandButton bRemoveBioRequests 
      Caption         =   "Remove All Biochemistry Requests"
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   3735
   End
   Begin VB.CommandButton bCoag 
      Caption         =   "Set All Coagulation Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton bBio 
      Caption         =   "Set All Biochemistry Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   3735
   End
   Begin VB.CommandButton bHaem 
      Caption         =   "Set All Haematology Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   3735
   End
End
Attribute VB_Name = "frmUpdatePrinted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bBio_Click()

      Dim sql As String

28470 On Error GoTo bBio_Click_Error

28480 sql = "Update BioResults " & _
            "Set Printed = 1 " & _
            "where Printed = 0"
28490 Cnxn(0).Execute sql

28500 Exit Sub

bBio_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28510 intEL = Erl
28520 strES = Err.Description
28530 LogError "fUpdatePrinted", "bBio_Click", intEL, strES, sql


End Sub


Private Sub bCoag_Click()

      Dim sql As String

28540 On Error GoTo bCoag_Click_Error

28550 sql = "Update CoagResults " & _
            "Set Printed = 1 " & _
            "where Printed = 0"
28560 Cnxn(0).Execute sql

28570 Exit Sub

bCoag_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28580 intEL = Erl
28590 strES = Err.Description
28600 LogError "fUpdatePrinted", "bCoag_Click", intEL, strES, sql


End Sub

Private Sub bHaem_Click()

      Dim sql As String

28610 On Error GoTo bHaem_Click_Error

28620 sql = "Update HaemResults " & _
            "Set Printed = 1 " & _
            "where Printed = 0"

28630 Cnxn(0).Execute sql

28640 Exit Sub

bHaem_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28650 intEL = Erl
28660 strES = Err.Description
28670 LogError "fUpdatePrinted", "bHaem_Click", intEL, strES, sql


End Sub


Private Sub Command1_Click()

End Sub


Private Sub bRemoveBioRequests_Click()

      Dim sql As String

28680 On Error GoTo bRemoveBioRequests_Click_Error

28690 sql = "Delete from BioRequests"

28700 Cnxn(0).Execute sql

28710 Exit Sub

bRemoveBioRequests_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28720 intEL = Erl
28730 strES = Err.Description
28740 LogError "fUpdatePrinted", "bRemoveBioRequests_Click", intEL, strES, sql


End Sub


Private Sub bRemoveCoagRequests_Click()

      Dim sql As String

28750 On Error GoTo bRemoveCoagRequests_Click_Error

28760 sql = "Delete from CoagRequests"

28770 Cnxn(0).Execute sql

28780 Exit Sub

bRemoveCoagRequests_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28790 intEL = Erl
28800 strES = Err.Description
28810 LogError "fUpdatePrinted", "bRemoveCoagRequests_Click", intEL, strES, sql


End Sub


Private Sub cmdCancel_Click()

28820 Unload Me

End Sub

Private Sub cmdRemoveHaemRequests_Click()

      Dim sql As String

28830 On Error GoTo cmdRemoveHaemRequests_Click_Error

28840 sql = "Delete from HaemRequests"

28850 Cnxn(0).Execute sql

28860 Exit Sub

cmdRemoveHaemRequests_Click_Error:

      Dim strES As String
      Dim intEL As Integer

28870 intEL = Erl
28880 strES = Err.Description
28890 LogError "fUpdatePrinted", "cmdRemoveHaemRequests_Click", intEL, strES, sql

End Sub


