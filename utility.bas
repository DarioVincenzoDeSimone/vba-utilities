Attribute VB_Name = "Modulo1"
'Utility

Public Function select_personal(fieldToSearch As String, sheet As String, firstFieldCondition As String, firstCondition As String, secondFieldCondition As String, secondCondition As String) As String
Dim fieldToSearch_col As Long
Dim firstCondition_col, firstCondition_row As Long
Dim secondCondition_col, secondCondition_row As Long
Dim select_personal_row As Long
Dim lastRow As Long

lastRow = getlastrow(sheet)

Set fieldToSearch_find = ThisWorkbook.Sheets(sheet).Rows(1).Find(fieldToSearch, LookAt:=xlWhole)
Set firstFieldCondition_find = ThisWorkbook.Sheets(sheet).Rows(1).Find(firstFieldCondition, LookAt:=xlWhole)
Set secondFieldCondition_find = ThisWorkbook.Sheets(sheet).Rows(1).Find(secondFieldCondition, LookAt:=xlWhole)

fieldToSearch_col = fieldToSearch_find.column
firstCondition_col = firstFieldCondition_find.column
secondCondition_col = secondFieldCondition_find.column

For i = 2 To lastRow
    If ThisWorkbook.Sheets(sheet).Cells(i, firstCondition_col).Text = firstCondition Then
        If ThisWorkbook.Sheets(sheet).Cells(i, secondCondition_col).Text = secondCondition Then
            select_personal = ThisWorkbook.Sheets(sheet).Cells(i, fieldToSearch_col)
            i = lastRow
        End If
    End If
   
Next i

End Function
Public Function selectCount(countfield As String, whereCondition As String, sheet As String) As Integer
Dim lastRow As Long
Dim fieldCol As Long
lastRow = getlastrow(sheet)
fieldCol = getCol(sheet, countfield, 1)

selectCount = 0
For x = 2 To lastRow
    If ThisWorkbook.Sheets(sheet).Cells(x, fieldCol).Text = whereCondition Then
        selectCount = selectCount + 1
    End If
Next x

End Function


Public Function getlastrow(sheet As String) As Long
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets(sheet)
'getlastrow = sht.Cells.SpecialCells(xlCellTypeLastCell).row 'OLD, not correct
 If Application.WorksheetFunction.CountA(sht.Cells) <> 0 Then
    getlastrow = sht.Cells.Find(What:="*", _
                     LookIn:=xlFormulas, _
                     SearchOrder:=xlByRows, _
                     SearchDirection:=xlPrevious, LookAt:=xlWhole).row
  Else
     getlastrow = 1
  End If
End Function

Public Function getlastcol(sheet As String) As Long
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets(sheet)
'getlastrow = sht.Cells.SpecialCells(xlCellTypeLastCell).row 'OLD, not correct
 If Application.WorksheetFunction.CountA(sht.Cells) <> 0 Then
    getlastcol = sht.Cells.Find(What:="*", _
                     LookIn:=xlFormulas, _
                     SearchOrder:=xlByColumns, _
                     SearchDirection:=xlPrevious, LookAt:=xlWhole).column
  Else
     getlastcol = 1
  End If
End Function

Public Function getCol(sheet As String, colName As String, headerRow As Integer) As Long
Set findField = ThisWorkbook.Sheets(sheet).Rows(headerRow).Find(colName, LookAt:=xlWhole)
getCol = findField.column
End Function
Public Function getRow(sheet As String, valueName As String, colNum As Integer) As Long
Set findField = ThisWorkbook.Sheets(sheet).Columns(colNum).Find(valueName, LookAt:=xlWhole)
getRow = findField.row
End Function

Public Function getList(sheetName As String, field As String, valueField As String, outputField As String) As String()

Dim fieldFindCol As Integer, outputFieldFindCol As Integer
Dim numberComponent As Integer
Dim nrow As Integer, i As Integer
Dim arr() As String
i = 0

'Set sheet = ThisWorkbook.Worksheets(sheetName)
Set fieldFind = ThisWorkbook.Worksheets(sheetName).Rows(1).Find(field, LookAt:=xlWhole)
fieldFindCol = fieldFind.column
Set outputFieldFind = ThisWorkbook.Worksheets(sheetName).Rows(1).Find(outputField, LookAt:=xlWhole)
outputFieldFindCol = outputFieldFind.column

numberComponent = Application.WorksheetFunction.CountIf(Worksheets(sheetName).Columns(fieldFindCol), valueField)
ReDim getList(numberComponent - 1)
ReDim arr(numberComponent - 1)
nrow = getlastrow(sheetName)

For row = 2 To nrow
    If ThisWorkbook.Sheets(sheetName).Cells(row, fieldFindCol) = valueField Then
        arr(i) = ThisWorkbook.Sheets(sheetName).Cells(row, outputFieldFindCol)
        i = i + 1
    End If

Next row
getList = arr
End Function


Public Sub CreateSheet(name As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = name
End Sub
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function colExists(sheet As String, colName As String, headerRow As Integer) As Boolean
    colExists = False
    Set findField = ThisWorkbook.Sheets(sheet).Rows(headerRow).Find(colName, LookAt:=xlWhole)
    If findField Is Nothing Then
        colExists = False
    Else
    colExists = True
    End If
End Function



