Attribute VB_Name = "Read_in"
Option Explicit

Sub ExcelProcesses(enable_disable As Boolean)
    ' Toggle Excel application settings to improve VBA code performance
    Application.ScreenUpdating = enable_disable
    Application.DisplayStatusBar = enable_disable
    Application.EnableEvents = enable_disable
    Application.Calculation = IIf(enable_disable, xlCalculationAutomatic, xlCalculationManual)
End Sub

Sub ClearInputSheet(Optional ws As Worksheet)
    'Clears the Workbook given, or by default the ActiveSheet, marco called by 'Clear Data' independently and called from ReadInData() Sub
    'Clears by finding where the data starts below the Counter-Party Cell (in this case uses Named Ranges). We do not want to clear the column headers
    Dim startRow As Integer
    Dim startColumn As Integer

    If ws Is Nothing Then Set ws = Sheet1
    startRow = ws.Range("Counter_Party").row
    startColumn = ws.Range("Counter_Party").Column
    ws.Range(ws.Cells(startRow + 1, startColumn), ws.Cells(100000, 500)).Delete
    ws.Shapes("Button 12").Visible = False ' export errors
    ws.Shapes("Button 16").Visible = False ' export to CPM
    ws.Shapes("Button 18").Visible = False ' Callibration Export
    ws.Shapes("Button 19").Visible = False ' Fix Duplicates
    ws.Shapes("Button 21").Visible = False ' Create modified workbook
    
    ws.Range("Contract_Number", ws.Range("Contract_Number").End(xlDown)).Sort Key1:=ws.Range("Contract_Number"), Order1:=xlAscending, Header:=xlYes
End Sub

Function WorksheetExists(shtName As String, wb As Workbook) As Boolean
    'Check if a WorkSheet exists by setting it to a variable, if an error is thrown, catch it and just return a MsgBox
    'Returns True if it exist, False if it doesnt exist
    Dim sht As Worksheet

    'If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    If Err.Number <> 0 Then
        MsgBox "Sheet '" & shtName & "' does not exist"
        Sheet2.Range("Error_Log").Offset(k + 1, 0) = "No sheet '" & shtName & "' in file: " & wb.name
        wb.Close SaveChanges:=False
        Application.GoTo Sheet2.Range("Error_Log"), True
    End If
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function GetHeaderCell(ws As Worksheet, columnName As String) As Variant
    'Returns the row and column index of columnName as a Array
    Dim c As Range

    With ws.UsedRange
        Set c = .Find(columnName, LookIn:=xlValues)
        If c Is Nothing Then
            MsgBox "Cannot find column header '" & columnName & "' in " & ws.name & " Sheet. Has it been changed?"
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = "Cannot find column header '" & columnName & "' in " & ws.name & " Sheet. Has it been changed?"
            Application.GoTo Sheet2.Range("Error_Log"), True
            Exit Function
        End If
    End With
    GetHeaderCell = Array(c.row, c.Column)
End Function

Sub ReadInData()
    ' Reads in data from the latest updates portfolio workbook found in directory written in the excel sheet highlighted in yellow
    ' It gets data from Flex/Fix sheets from the workbook and copies and pastes the data into the check_data "input" worksheet.
    ' Before it pastes the data, it makes sure it will be pasting it to the same Columns, so the format is the same
    ' Called by using "Read In Data" Button from check_data workbook.

    Dim objWorkbook As Workbook 'Stores the updated portfolio data in our 'objective' workbook
    Dim lastRowFlex As Long 'Number of rows in flex sheet from obj workbook
    Dim lastRowFix As Long 'Number of rows in fix sheet from obj workbook
    Dim directory As String 'Directory of updated portfolio data
    Dim CD_ws As Worksheet ' Check_Data (CD) 'Input' worksheet
    Dim Flex_ws As Worksheet 'Flex worksheet from updated portfolio data
    Dim Fix_ws As Worksheet 'Fix worksheet from updated portfolio data
    Dim CD_format As New Collection ' stores all column headers in order from CD Input sheet
    Dim CD_CP(2) As Integer 'CD Counter-Party (CP) cell row (0) and column (1)
    Dim CD_CID(2) As Integer 'CD Customer ID (CID) cell row (0) and column (1)
    Dim columnHeader As Range 'Where all the column headers are
    Dim Flex_CP As Variant 'Counter-Party (CP) Cell in flex sheet
    Dim Flex_CID As Variant 'Customer ID (CID) Cell in flex sheet
    Dim fixCutoffColumn As Integer 'Column index in CD 'Input' sheet before the monthly open/closed start (this is not present in fix data from updated portfolio data)
    Dim fixCutoffName As String 'Name of the column preceding the open/closed columns
    Dim FixSheetCutoff As Variant 'Cell in fix sheet which matches the fixCutoffName above
    Dim FixSheetRestartColumn As Integer 'This is just +1 of the column index of variable above, as in fix data sheet there are no closed/open volumes
    Dim FixRestartColumn As Integer 'Column index in CD sheet where closed/open volume stops: 360 + fixCutoffColumn
    Dim Fix_CP As Variant 'Counter-Party (CP) Cell in fix sheet
    Dim Fix_CID As Variant 'Customer ID (CID) Cell in fix sheet
    Dim flexRange As Range 'The range we are copying from flex data sheet
    Dim FixRange1 As Range 'The first range we are copying from fix data sheet up until the FixSheetCutoff
    Dim FixRange2 As Range 'The second range we are copying from fix data sheet after the FixSheetRestartColumn

    ' Clear old Error Logs
    k = 0
    Sheet2.Range("Error_Log").Offset(1, 0).Resize(100000, 1).ClearContents

    ' Directory of excel sheet with flex and fix sheets
    directory = Range("Read_In_Directory").Value
    directory = Replace(directory, Chr(34), vbNullString) ' removes speech marks if it has any from path

    ExcelProcesses False

    ' Open the workbook (ReadOnly is important... also error checking directory)
    On Error Resume Next
    Set objWorkbook = Workbooks.Open(directory, ReadOnly:=True)
    If Err.Number <> 0 Then
        MsgBox "Could not find the file " & directory
        Exit Sub
    End If
    On Error GoTo 0

    ' Check and set Worksheets where data will be copied (Read in flex/fix) or pasted (CD) to
    Set CD_ws = ThisWorkbook.Sheets("Input")
    If Not WorksheetExists("Flex", objWorkbook) Then Exit Sub: objWorkbook.Close SaveChanges:=False
    Set Flex_ws = objWorkbook.Worksheets("Flex")
    If Not WorksheetExists("Fix", objWorkbook) Then Exit Sub: objWorkbook.Close SaveChanges:=False
    Set Fix_ws = objWorkbook.Worksheets("Fix")

    ' Clear old portfolio data in CD
    ClearInputSheet CD_ws

    ' Get address of Counter-Party (CP) cell (we assume this is the first one of interest)
    CD_CP(0) = CD_ws.Range("Counter_Party").row
    CD_CP(1) = CD_ws.Range("Counter_Party").Column

    ' Get address of Customer ID (CID) cell (we assume this is the last one of interest)
    CD_CID(0) = CD_ws.Range("Cust_ID").row
    CD_CID(1) = CD_ws.Range("Cust_ID").Column

    ' Adds string of each column header to the collection to cross-check with Read In data
    For Each columnHeader In CD_ws.Range(CD_ws.Cells(CD_CP(0), CD_CP(1)), CD_ws.Cells(CD_CID(0), CD_CID(1)))
        CD_format.Add columnHeader.Value
    Next

    ' Flex Sheet
    ' Get address of Counter-Party Cell
    Flex_CP = GetHeaderCell(Flex_ws, "Counter-Party")
    If IsEmpty(Flex_CP) Then Exit Sub: objWorkbook.Close SaveChanges:=False

    ' Get address of Customer ID cell
    Flex_CID = GetHeaderCell(Flex_ws, "Customer Number")
    If IsEmpty(Flex_CID) Then Exit Sub: objWorkbook.Close SaveChanges:=False

    ' Check Format/Order of Column Headers (Flex)
    If Not CheckFormat(Flex_ws, Flex_CP, Flex_CID, CD_CP, CD_format) Then objWorkbook.Close SaveChanges:=False: Exit Sub

    ' Find cutoff column for Fix data
    fixCutoffColumn = CD_ws.Range("M1_").Column - 1
    fixCutoffName = CD_ws.Cells(CD_CP(0), fixCutoffColumn).Value

    ' Find this column in Fix sheet
    FixSheetCutoff = GetHeaderCell(Fix_ws, fixCutoffName)
    If IsEmpty(FixSheetCutoff) Then Exit Sub: objWorkbook.Close SaveChanges:=False

    FixSheetRestartColumn = FixSheetCutoff(1) + 1
    FixRestartColumn = fixCutoffColumn + 360 + 1

    ' Fix Sheet
    ' Get address of Counter-Party Cell
    Fix_CP = GetHeaderCell(Fix_ws, "Counter-Party")
    If IsEmpty(Fix_CP) Then Exit Sub: objWorkbook.Close SaveChanges:=False

    ' Get address of Customer ID cell
    Fix_CID = GetHeaderCell(Fix_ws, "Customer number")
    If IsEmpty(Fix_CID) Then Exit Sub: objWorkbook.Close SaveChanges:=False

    ' Check Format/Order of Fix column headers (Fix)
    If Not CheckFormat(Fix_ws, Fix_CP, Fix_CID, CD_CP, CD_format, fixCutoffColumn) Then objWorkbook.Close SaveChanges:=False: Exit Sub

    ' Copy and paste data into data checker
    ' Flex Data
    lastRowFlex = Flex_ws.Cells(Flex_ws.Rows.Count, Flex_CP(1)).End(xlUp).row
    Set flexRange = Flex_ws.Range(Flex_ws.Cells(Flex_CP(0) + 1, Flex_CP(1)), Flex_ws.Cells(lastRowFlex, Flex_CID(1)))
    flexRange.Copy CD_ws.Cells(CD_CP(0) + 1, CD_CP(1))

    ' Fix Data
    lastRowFix = Fix_ws.Cells(Fix_ws.Rows.Count, Fix_CP(1)).End(xlUp).row
    Set FixRange1 = Fix_ws.Range(Fix_ws.Cells(Fix_CP(0) + 1, Fix_CP(1)), Fix_ws.Cells(lastRowFix, FixSheetCutoff(1)))
    FixRange1.Copy CD_ws.Cells(3 + lastRowFlex, CD_CP(1))

    Set FixRange2 = Fix_ws.Range(Fix_ws.Cells(Fix_CP(0) + 1, FixSheetRestartColumn), Fix_ws.Cells(lastRowFix, Fix_CID(1)))
    FixRange2.Copy CD_ws.Cells(3 + lastRowFlex, FixRestartColumn)

    ' Clear the clipboard
    Application.CutCopyMode = False

    ' Close the workbooks
    objWorkbook.Close SaveChanges:=False
    
    ' Set Multiple ID Column to N (-3 for headers)
    ' CD_ws.Range(CD_ws.Range("Multiple?").Offset(1, 0), CD_ws.Range("Multiple?").Offset(CD_ws.Range("Cust_ID").End(xlDown).row - 3, 0)).Value = "N"
    CD_ws.Range("Modified?").Value = "N"

    ' Center view back top left so buttons are visible
    Application.GoTo CD_ws.Range("A3"), True

    ExcelProcesses True

    MsgBox "Data Successfully Read In"
End Sub

Function CheckFormat(ws As Worksheet, startCell As Variant, endCell As Variant, CD_startCell As Variant, CD_format As Collection, Optional cutoffColumn As Integer = -1) As Boolean
    'Checks the columns are in the same order (or just the same in general) before reading in the data from the flex/fix sheets to the CD Input Sheet
    Dim formatRange As Range
    Dim i As Long, CD_i As Long
    
    Set formatRange = ws.Range(ws.Cells(startCell(0), startCell(1)), ws.Cells(endCell(0), endCell(1)))
    For i = 1 To formatRange.Cells.Count
        If cutoffColumn > -1 And i > cutoffColumn - (CD_startCell(1) - 1) Then
            CD_i = i + 360
        Else
            CD_i = i
        End If
        If formatRange.Cells(i).Value <> CD_format(CD_i) Then
            MsgBox "Format: columns '" & formatRange.Cells(i).Value & "' and '" & CD_format(CD_i) & "' do not align. Please check format of columns."
            Sheet2.Range("Error_Log").Offset(k + 1, 0).Value = "Sheet: " & ws.name & ", Format: columns '" & formatRange.Cells(i).Value & "' and '" & CD_format(CD_i) & "' do not align. Please check format of columns."
            CheckFormat = False
            Exit Function
        End If
    Next
    CheckFormat = True
End Function
