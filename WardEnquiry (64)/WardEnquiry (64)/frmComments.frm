VERSION 5.00
Begin VB.Form frmComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Comments"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   8010
      Picture         =   "frmComments.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4290
      Width           =   735
   End
   Begin VB.TextBox txtComment 
      Height          =   4695
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmComments.frx":066A
      Top             =   270
      Width           =   7575
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Sub Form_Activate()
10    LoadComments
20    SingleUserUpdateLoggedOn UserName

End Sub

Private Sub LoadComments()

Dim OBS As Observations
Dim OB As Observation
Dim S As String
'Dim EGFRComment() As String
Dim n As Integer
Dim AutoComment As String

On Error GoTo LoadComments_Error

S = ""

'30    If GetEGFRComment(pSampleID, EGFRComment) Then
'40    For n = 0 To UBound(EGFRComment)
'50      S = S & EGFRComment(n) & vbCrLf
'60    Next
'70    End If

Set OBS = New Observations
Set OBS = OBS.Load(pSampleID, "Demographic", "Biochemistry", "Haematology", "Coagulation", "Film", "Immunology", "NVRL", "Biomnis", "MATER", "Beaumont")

If Not OBS Is Nothing Then

    For Each OB In OBS
        Select Case UCase$(OB.Discipline)
        Case "DEMOGRAPHIC": S = S & "Demographic: " & OB.Comment & vbCrLf
        Case "BIOCHEMISTRY": S = S & "Biochemistry: " & OB.Comment & vbCrLf
        Case "HAEMATOLOGY": S = S & "Haematology: " & OB.Comment & vbCrLf
        Case "COAGULATION": S = S & "Coagulation: " & OB.Comment & vbCrLf
        Case "FILM": S = S & "Film: " & OB.Comment & vbCrLf
        Case "IMMUNOLOGY": S = S & "Immunology: " & OB.Comment & vbCrLf
        Case "BIOMNIS": S = S & "Biomnis: " & OB.Comment & vbCrLf
        Case "NVRL": S = S & " NVRL : " & OB.Comment & vbCrLf
        Case "MATLAB": S = S & "MATLAB: " & OB.Comment & vbCrLf
        Case "Beaumont": S = S & "Beaumont: " & OB.Comment & vbCrLf
        End Select
    Next
End If

AutoComment = CheckAutoComments(pSampleID, 2)
If Trim$(AutoComment) <> "" Then
    S = S & "Biochemistry: " & AutoComment & vbCrLf
End If
AutoComment = CheckAutoComments(pSampleID, 3)
If Trim$(AutoComment) <> "" Then
    S = S & "Coagulation: " & AutoComment & vbCrLf
End If

txtComment = S

Exit Sub

LoadComments_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmComments", "LoadComments", intEL, strES

End Sub


Public Property Let SampleID(ByVal sNewValue As String)

10    pSampleID = sNewValue

End Property
