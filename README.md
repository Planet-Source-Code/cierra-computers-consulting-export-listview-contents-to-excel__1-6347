<div align="center">

## Export ListView Contents to Excel


</div>

### Description

This will export the contents of a listview into a new Excel Workbook.
 
### More Info
 
frmName.ListView


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cierra Computers & Consulting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cierra-computers-consulting.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cierra-computers-consulting-export-listview-contents-to-excel__1-6347/archive/master.zip)

### API Declarations

Reference the MS Excel Object


### Source Code

```

Public Sub ExportListViewtoExcel(lvwList As Control)
   Dim vntHeader As Variant
   Dim vntData As Variant
   Dim x As Long
   Dim y As Long
   Dim intCol As Integer
   Dim lngRow As Long
   'Get Counts
   intCol = CInt(lvwList.ColumnHeaders.Count - 1)
   lngRow = CLng(lvwList.ListItems.Count - 1)
   ReDim vntHeader(0)
   ReDim vntData(intCol, lngRow)
   'Create Header Array
   For x = 0 To intCol
     ReDim Preserve vntHeader(x)
     vntHeader(x) = lvwList.ColumnHeaders(x + 1).Text
   Next
   'Create Data Array
   For x = 0 To lngRow
    vntData(0, x) = lvwList.ListItems.Item(x + 1).Text
    For y = 1 To intCol
      vntData(y, x) = lvwList.ListItems.Item(x + 1).SubItems(y)
    Next
   Next
   'Create Excel Object
   OpenExcel vntData, vntHeader
End Sub
Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
  Dim lngRow As Long
  Dim intCol As Integer
  Dim varData As Variant
  Dim intStart As Integer
  'Select all Cells and and set the number format to string
  ws.Cells.Select
  ws.Cells.NumberFormat = "@"
  ws.Cells(1, 1).Select
  lngRow = UBound(vntData, 2) + 2
  intCol = UBound(vntData, 1) + 1
  intStart = 2  'Start from line 2
   'Freeze Row 2
   ws.Rows(2).Select
   ws.Activate
   ActiveWindow.FreezePanes = True
   'Add Headers
   For x = 1 To intCol
      varData = vntHeader(x - 1)
      ws.Cells(1, x) = CStr(varData)
      ws.Cells(1, x).Font.Bold = True
   Next
  'Add Data
  For y = 1 To intCol
     For x = intStart To lngRow
        varData = vntData(y - 1, x - 2)
        If IsNull(varData) Then 'Make sure no null values, Excel will choke
             'Add 1 to Move down a column
          ws.Cells(x + 1, y) = ""
        Else
          ws.Cells(x + 1, y) = CStr(varData) 'Convert to String to preserve formatting
        End If
     Next
  Next
  'Resize Columns to Fit
   ws.Columns.AutoFit
End Sub
Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
On Error GoTo Err_OpenExcel
Dim objExcel As Excel.Application
Dim objWrkSht As Worksheet
Dim x As Integer
'Create Excel Object
Set objExcel = CreateObject("Excel.Application")
'Add the Workbook
objExcel.Workbooks.Add
Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
objExcel.Visible = True
'Fill the Workbook with data
ExportRecords vntData, vntHeader, objWrkSht
objExcel.Interactive = True
' Clean up:
Set objExlSht = Nothing
Set objExcel = Nothing
Err_OpenExcel:
   Select Case Err
     Case 0
     Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
     Case Else
        MsgBox Err & ": " & Error, vbCritical, "OpenExcel Error"
   End Select
End Sub
```

