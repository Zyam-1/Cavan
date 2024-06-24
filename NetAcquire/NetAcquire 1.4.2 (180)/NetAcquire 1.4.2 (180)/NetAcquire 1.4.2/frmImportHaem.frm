VERSION 5.00
Begin VB.Form frmImportHaem 
   Caption         =   "Import"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label lblSampleID 
      Height          =   255
      Left            =   1500
      TabIndex        =   0
      Top             =   1140
      Width           =   2055
   End
End
Attribute VB_Name = "frmImportHaem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()

ImportHaem

End Sub

Private Sub ImportHaem()

Dim sql As String
Dim tb As Recordset
Dim Con As New Connection

Con = "DRIVER={SQL Server};Server=(local);Database=Port69Live;uid=sa;pwd=;"
Con.Open

sql = "Select * from HaemResults"
Set tb = New Recordset
With tb
  .CursorLocation = adUseServer
  .CursorType = adOpenDynamic
  .LockType = adLockOptimistic
  .ActiveConnection = Con
  .Source = sql
  .Open
End With
Do While Not tb.EOF
  lblSampleID = tb!SampleID
  lblSampleID.Refresh
  If Val(tb!WBC & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'WBC', " & _
          " '" & tb!WBC & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!RBC & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'RBC', " & _
          " '" & tb!RBC & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Hgb & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Hgb', " & _
          " '" & tb!Hgb & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Hct & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Hct', " & _
          " '" & tb!Hct & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MCV & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MCV', " & _
          " '" & tb!MCV & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MCH & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MCH', " & _
          " '" & tb!MCH & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MCHC & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MCHC', " & _
          " '" & tb!MCHC & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Plt & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Plt', " & _
          " '" & tb!Plt & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!RDWSD & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'RDWSD', " & _
          " '" & tb!RDWSD & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!RDWCV & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'RDWCV', " & _
          " '" & tb!RDWCV & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MPV & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MPV', " & _
          " '" & tb!MPV & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!PDW & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'PDW', " & _
          " '" & tb!PDW & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!LymP & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'LymP', " & _
          " '" & tb!LymP & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MonoP & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MonoP', " & _
          " '" & tb!MonoP & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!NeutP & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'NeutP', " & _
          " '" & tb!NeutP & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!EosP & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'EosP', " & _
          " '" & tb!EosP & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!BasP & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'BasP', " & _
          " '" & tb!BasP & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!LymA & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'LymA', " & _
          " '" & tb!LymA & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!MonoA & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'MonoA', " & _
          " '" & tb!MonoA & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!NeutA & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'NeutA', " & _
          " '" & tb!NeutA & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!EosA & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'EosA', " & _
          " '" & tb!EosA & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!BasA & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'BasA', " & _
          " '" & tb!BasA & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Plcr & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Plcr', " & _
          " '" & tb!Plcr & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Retics & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Retics', " & _
          " '" & tb!Retics & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!Monospot & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'Monospot', " & _
          " '" & tb!Monospot & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  If Val(tb!ESR & "") <> 0 Then
    sql = "INSERT INTO HaemResults " & _
          "( SampleID, Code, Result, Valid, Printed, RunTime, RunDate ) VALUES " & _
          "('" & tb!SampleID & "', " & _
          " 'ESR', " & _
          " '" & tb!ESR & "', " & _
          " '" & tb!Valid & "', " & _
          " '" & tb!Printed & "', " & _
          " '" & Format(tb!RunDateTime, "Long Date") & " " & Format(tb!RunDateTime, "Long Time") & "', " & _
          " '" & Format(tb!Rundate, "Long Date") & "')"
    Cnxn(0).Execute sql
  End If
  tb.MoveNext
Loop

End Sub


