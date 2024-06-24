Attribute VB_Name = "modExportExcel"
Option Explicit

Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          ByVal CallingForm As Form)

      Dim objXL As Object
      Dim objWB As Object
      Dim objWS As Object
      Dim R As Long
      Dim c As Long
      Dim T As Single
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

100   With objWS
110     For R = 0 To objGrid.Rows - 1
120       For c = 0 To objGrid.Cols - 1
130         If c < 256 Then '0-255 max columns in Excel
              'The "'" is required to format the cells as text in Excel
              'otherwise entries like "4/2" are interpreted as a date
140           .Cells(R + 1, c + 1) = "'" & objGrid.TextMatrix(R, c)
150         End If
160       Next
170     Next

180     .Cells.Columns.AutoFit
190   End With

      'This line removes the green error triangle
      'Does not work with XL2000!!! - Requires at least XL2002
      'objXL.ErrorCheckingOptions.NumberAsText = False

200   objXL.Visible = True

210   Set objWS = Nothing
220   Set objWB = Nothing
230   Set objXL = Nothing

240   CallingForm.lblExcelInfo.Visible = False

250   Exit Sub

ehEFG:
      Dim er As Long
      Dim es As String

260   er = Err.Number
270   es = Err.Description

280   iMsg es

290   With CallingForm.lblExcelInfo
300     .Caption = "Error " & Format(er)
310     .Refresh
320     T = Timer
330     Do While Timer - T < 1: Loop
340     .Visible = False
350   End With

360   Exit Sub

End Sub


