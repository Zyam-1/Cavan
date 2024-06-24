Attribute VB_Name = "modExportExcel"
Option Explicit

Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          ByVal CallingForm As Form, Optional ByVal HeadingMatrix As String = "")

      Dim objXL As Object
      Dim objWB As Object
      Dim objWS As Object
      Dim R As Long
      Dim C As Long
      Dim t As Single
      'Assume the calling form has a MSFlexGrid (grdToExport),
      'CommandButton (cmdXL) and Label (lblExcelInfo) (Visible set to False)
      'In the calling form:
      'Private Sub cmdXL_Click()
      'ExportFlexGrid grdToExport, Me
      'End Sub

10    On Error GoTo ehEFG

20    With CallingForm.lblExcelInfo
30      .Caption = "Exporting..."
40      .Visible = True
50      .Refresh
60    End With

70    Set objXL = CreateObject("Excel.Application")
80    Set objWB = objXL.Workbooks.Add
90    Set objWS = objWB.Worksheets(1)

      Dim intLineCount As Integer
      '****Change: Babar Shahzad 2007-11-19
      'Heading for export to excel can be passed as string which would be
      'a string having TABS as column breaks and CR as row break.

100   intLineCount = 0
110   If HeadingMatrix <> "" Then
120       With objWS
              Dim strTokens() As String
130           strTokens = Split(HeadingMatrix, vbCr)
140           intLineCount = UBound(strTokens)
  
150           For R = LBound(strTokens) To UBound(strTokens) - 1
                  'For C = 0 To objGrid.Cols - 1
                'The "'" is required to format the cells as text in Excel
                'otherwise entries like "4/2" are interpreted as a date
160             .range(.cells(R + 1, 1), .cells(R + 1, objGrid.Cols)).MergeCells = True
170             .range(.cells(R + 1, 1), .cells(R + 1, objGrid.Cols)).HorizontalAlignment = 3
180             .range(.cells(R + 1, 1), .cells(R + 1, objGrid.Cols)).Font.Bold = True
190             objWS.cells(R + 1, 1) = "'" & strTokens(R)
    
200           Next
210       End With
   
220   End If

230   With objWS
240     For R = 0 To objGrid.Rows - 1
250       For C = 0 To objGrid.Cols - 1
            'The "'" is required to format the cells as text in Excel
            'otherwise entries like "4/2" are interpreted as a date
260         .cells(R + 1 + intLineCount, C + 1) = "'" & objGrid.TextMatrix(R, C)
270       Next
280     Next

290     .cells.Columns.AutoFit
300   End With

310   objXL.Visible = True

320   Set objWS = Nothing
330   Set objWB = Nothing
340   Set objXL = Nothing

350   CallingForm.lblExcelInfo.Visible = False

360   Exit Sub

ehEFG:
      Dim er As Long
      Dim es As String

370   er = Err.Number
380   es = Err.Description

390   iMsg es

400   With CallingForm.lblExcelInfo
410     .Caption = "Error " & Format(er)
420     .Refresh
430     t = Timer
440     Do While Timer - t < 1: Loop
450     .Visible = False
460   End With

470   Exit Sub

End Sub

