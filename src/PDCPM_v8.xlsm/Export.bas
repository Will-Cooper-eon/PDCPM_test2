Attribute VB_Name = "Export"
Option Explicit

Sub Export()
    ' Main macro called from the export button
    
    If safeToExport = "yes" Then
        Dim CD_ws As Worksheet 'Check Data (CD) worksheet
        Set CD_ws = ThisWorkbook.Sheets("Input")
    
        ' If "Modified?" = "Y"
        If CD_ws.Range("Modified?") = "Y" Then
            'Dim modified_path As String
            'modified_path = CreateModifiedFile
            'Call CreateModifiedFile
            'MsgBox "Modified"
            Call ExportToCPM(CreateModifiedFile)
        Else ' If not
            Call ExportToCPM
        End If
    Else
        MsgBox "Please check the format of the data using the 'Check Format' button before exporting."
    End If
    
End Sub

Function CreateModifiedFile() As String
    ' Creates a  portfolio data file after contracts from a customer with multiple IDs have been modified
    
    Dim CD_ws As Worksheet 'Check Data (CD) worksheet
    Set CD_ws = ThisWorkbook.Sheets("Input")
    
    Dim r As Long ' row iterator
    Dim Contract_Change_Row As Long ' last fix contract row index
    
    ' Sort via Flex_Fix column (Fix will be above flex because alphabet)
    CD_ws.Range(CD_ws.Range("Contract_Number"), CD_ws.Range("Cust_ID").End(xlDown)).Sort Key1:=CD_ws.Range("Flex_Fix"), Order1:=xlAscending, Header:=xlYes, MatchCase:=True

    ' Find the row where fix contracts end, and flex contracts start.
    For r = CD_ws.Range("Counter_Party").row + 2 To CD_ws.Range("Counter_Party").End(xlDown).row ' (+2 to skip header and first row)
        If CD_ws.Cells(r, CD_ws.Range("Flex_Fix").Column) <> CD_ws.Cells(r - 1, CD_ws.Range("Flex_Fix").Column) Then
            Contract_Change_Row = r - 1
        End If
    Next r
    
    ' Create new work book
    Dim Modified_Data_wb As Workbook
    Set Modified_Data_wb = Workbooks.Add
    With Modified_Data_wb
    .title = "" & ExtractMonthYear(CD_ws.Range("Read_In_Directory")) & "Modified"
    '.SaveAs fileName:="" & Application.ThisWorkbook.Path & "\" & ExtractMonthYear(CD_ws.Range("Read_In_Directory")) & " Modified.xls"
    End With
    
    ' Add Fix and Flex sheets
    Modified_Data_wb.Worksheets("Sheet1").name = "Fix"
    Modified_Data_wb.Sheets.Add.name = "Flex"
    
    ' Copy and paste data before (fix) and after (flex) the row cutoff found above into their respective sheets
    Dim Flex_Range As Range
    Set Flex_Range = CD_ws.Range(CD_ws.Cells(Contract_Change_Row + 1, CD_ws.Range("Counter_Party").Column), CD_ws.Cells(CD_ws.Range("Counter_Party").End(xlDown).row, CD_ws.Range("Cust_ID").Column))
    Flex_Range.Copy Destination:=Modified_Data_wb.Sheets("Flex").Range("A2")
    
    Dim Fix_Range As Range
    Set Fix_Range = CD_ws.Range(CD_ws.Cells(CD_ws.Range("Counter_Party").row + 1, CD_ws.Range("Counter_Party").Column), CD_ws.Cells(Contract_Change_Row, CD_ws.Range("M1_").Column - 1))
    Fix_Range.Copy Destination:=Modified_Data_wb.Sheets("Fix").Range("A2")
    Set Fix_Range = CD_ws.Range(CD_ws.Cells(CD_ws.Range("Counter_Party").row + 1, CD_ws.Range("M1_").Column + 360), CD_ws.Cells(Contract_Change_Row, CD_ws.Range("Cust_ID").Column))
    Fix_Range.Copy Destination:=Modified_Data_wb.Sheets("Fix").Cells(2, (CD_ws.Range("M1_").Column - CD_ws.Range("Counter_Party").Column) + 1)
    
    Dim Header_Range As Range
    Set Header_Range = CD_ws.Range(CD_ws.Range("Counter_Party"), CD_ws.Range("Cust_ID"))
    Header_Range.Copy Destination:=Modified_Data_wb.Sheets("Fix").Range("A1")
    Header_Range.Copy Destination:=Modified_Data_wb.Sheets("Flex").Range("A1")
    
    Call SetUsedColumnsWidthAndHeight(Modified_Data_wb.Sheets("Fix"), 20, 15)
    Call SetUsedColumnsWidthAndHeight(Modified_Data_wb.Sheets("Flex"), 20, 15)
    
    ' Save and leave open for user to double check
    Modified_Data_wb.SaveAs fileName:="" & Application.ThisWorkbook.Path & "\" & ExtractMonthYear(CD_ws.Range("Read_In_Directory")) & " Modified.xls"
    
    CreateModifiedFile = "" & Application.ThisWorkbook.Path & "\" & ExtractMonthYear(CD_ws.Range("Read_In_Directory")) & " Modified.xls"
    
End Function

Sub CreateModifiedFileMacro()
    Call CreateModifiedFile
End Sub

Sub ExportToCPM(Optional modified_portfolio_wb As String)
    'This Sub, as the name suggests, exports all the checked data directly into the CPM model
    'This removes another step, which was previously prone to human error by manually copy and pasting
    'In some places we assume where we export the data to on the sheet is the same across contract type, but not the same across each contract type (fix/flex)
   
    'This makes sure the data has had its format checking, by using the 'Check Format' Button which calls the check_data() Sub in module 1 and SortAndCheckIds() in Module 2
    'Only if it returns zero errors will this statement be satisfied
    If safeToExport = "yes" Then ' does this actually do anything? as the button to run this macro is only visible when there are no errors
    
        Dim ReadIn_directory As String 'Directory for updated portfolio data that we are reading data in from to copy
        Dim Export_directory As String 'Direcrory for CPM excel model to paste the read in data
        Dim CPMWorkbook As Workbook 'Uses string directory above to open the CPM workbook
        Dim ReadInWorkbook As Workbook 'Uses string directory above to open the updated portfolio data workbook
        Dim Fix_ws As Worksheet 'Read in workbook has a fix data sheet
        Dim Flex_ws As Worksheet 'Read in workbook has a flex data sheet
        Dim Fix_Entire_data As Worksheet 'Sheet in the CPM that stores all fixed data
        Dim Flex_Entire_data As Worksheet 'Sheet in the CPM that stores all flex data
        Dim Fix_less10_ws As Worksheet 'Sheet in the CPM that stores all fixed data used to prices <10GWh clusters
        Dim Flex_less10_ws As Worksheet 'Sheet in the CPM that stores all flex data used to prices <10GWh clusters
        Dim Fix_Above10_ws As Worksheet 'Sheet in the CPM that stores all fixed data above 10GWh to be bespoke priced
        Dim Flex_Above10_ws As Worksheet 'Sheet in the CPM that stores all flex data above 10GWh to be bespoke priced
        Dim Fix_Export As Variant 'Cell in the CPM sheet where fixed read in data will be pasted (exported)
        Dim Flex_Export As Variant 'Cell in the CPM sheet where flex read in data will be pasted (exported)
        Dim Flex_CP As Variant 'Counter-Party (CP) Cell in flex sheet
        Dim Flex_CID As Variant 'Customer ID (CID) Cell in flex sheet
        Dim Fix_CP As Variant 'Counter-Party (CP) Cell in fix sheet
        Dim Fix_CID As Variant 'Customer ID (CID) Cell in fix sheet
        Dim Fix_EM As Variant 'Expected Margin (EM) Cell in fix sheet
        Dim Flex_EM As Variant 'Expected Margin (EM) Cell in flex sheet
        Dim Flex_Year1 As Variant 'Year 1 Cell in flex sheet
        Dim Fix_Year1 As Variant 'Year 1 Cell in fix sheet
        Dim lastRow_flex As Integer 'Number of rows in flex sheet from read in, updated portfolio data, workbook
        Dim Flex_Range As Range 'Range where flex data is in the flex sheet
        Dim CD_ws As Worksheet ' Check_Data (CD) 'Input' worksheet
        Dim CD_CP(2) As Integer 'CD Counter-Party (CP) cell row (0) and column (1)
        Dim CD_CID(2) As Integer 'CD Customer ID (CID) cell row (0) and column (1)
        Dim lastRow_fix As Integer 'Number of rows in fix sheet from read in, updated portfolio data, workbook
        Dim Fix_Range_1 As Range 'The first range we are copying from fix data sheet up until the fix_cutoff_column
        Dim i As Integer 'Iterator
        Dim volume As Double 'Collective volume of counterparty
        Dim fix_i As Integer 'Iterator for fix data
        Dim less10_i As Integer 'Iterator for <10GWh data
        Dim contract_count As Integer 'How many contracts a singular counter-party has
        Dim flex_i As Integer 'Iterator for flex data
        Dim x As Long 'Iterator for the contracts a singular counter-party has
        Dim lastRow_CD As Integer 'Number of rows in CD 'Input' sheet
        Dim flex_fix_column As Integer 'Column index of column in CD which states if it's a fix or flex contract
        Dim fix_cutoff_column As Integer 'Column index where the monthly closed/open volume starts
        Dim fix_lastRow_above10 As Long 'Number of rows in fixed data >10GWh sheet
        Dim flex_lastRow_above10 As Long 'Number of rows in flex data >10GWh sheet
        
        ExcelProcesses (False)
        
        'Clear old Error Logs
        k = 0
        Sheet2.Range("Error_Log").Offset(1, 0).Resize(100000, 1).ClearContents
        
        'Get sheets for Check_Data (CD) workbook (ie the ActiveWorkbook) and used Named Ranges to get Rows/Columns of data
        If Read_in.WorksheetExists("Input", ThisWorkbook) = False Then Exit Sub
        Set CD_ws = ThisWorkbook.Sheets("Input")
        'Get address of Counter-Party (CP) cell (we assume this is first column of interest)
        CD_CP(0) = CD_ws.Range("Counter_Party").row
        CD_CP(1) = CD_ws.Range("Counter_Party").Column
        'get address of Customer ID (CID) cell ( we assume this is last column of interest)
        CD_CID(0) = CD_ws.Range("Cust_ID").row
        CD_CID(1) = CD_ws.Range("Cust_ID").Column
        
        'Get directories from the excel sheet, highlighted in yellow and green
        If CD_ws.Range("Modified?") = "Y" Then ' ThisWorkbook.Sheets("Input").Range("Modified?") = "Y" Then
            ReadIn_directory = modified_portfolio_wb
            ReadIn_directory = Replace(ReadIn_directory, Chr(34), vbNullString) ' removes speech marks if it has any from path
        Else
            ReadIn_directory = CD_ws.Range("Read_In_Directory").Value
            ReadIn_directory = Replace(ReadIn_directory, Chr(34), vbNullString) ' removes speech marks if it has any from path
        End If
        Export_directory = CD_ws.Range("Export_Directory").Value
        Export_directory = Replace(Export_directory, Chr(34), vbNullString)
    
        'Open up CPM Model as a Workbook
        On Error Resume Next
        Set CPMWorkbook = Workbooks.Open(Export_directory)
        If Err.Number <> 0 Then
            MsgBox "Could not find the file " & Export_directory
            Exit Sub
        End If
        On Error GoTo 0
        
        'Open ReadIn workbook (ie the latest portfolio available)
        On Error Resume Next
        Set ReadInWorkbook = Workbooks.Open(ReadIn_directory, ReadOnly:=True)
        If Err.Number <> 0 Then
            MsgBox "Could not find the file " & ReadIn_directory
            Exit Sub
        End If
        On Error GoTo 0

        'Open (almost) all required WorkSheets from both directories above and set them to variables
        'Read in sheets
        If Read_in.WorksheetExists("Flex", ReadInWorkbook) = False Then Exit Sub
        Set Flex_ws = ReadInWorkbook.Worksheets("Flex")
        If Read_in.WorksheetExists("Fix", ReadInWorkbook) = False Then Exit Sub
        Set Fix_ws = ReadInWorkbook.Worksheets("Fix")
        
        'Export Sheets
        If Read_in.WorksheetExists("Entire_Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
        Set Fix_Entire_data = CPMWorkbook.Worksheets("Entire_Portfolio_Data_Fixed")
        If Read_in.WorksheetExists("Entire_Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
        Set Flex_Entire_data = CPMWorkbook.Worksheets("Entire_Portfolio_Data_Flex")
        
        If Read_in.WorksheetExists("Less10_Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
        Set Fix_less10_ws = CPMWorkbook.Worksheets("Less10_Portfolio_Data_Fixed")
        If Read_in.WorksheetExists("Less10_Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
        Set Flex_less10_ws = CPMWorkbook.Worksheets("Less10_Portfolio_Data_Flex")
        
        If Read_in.WorksheetExists("Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
        Set Fix_Above10_ws = CPMWorkbook.Worksheets("Portfolio_Data_Fixed")
        If Read_in.WorksheetExists("Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
        Set Flex_Above10_ws = CPMWorkbook.Worksheets("Portfolio_Data_Flex")
            
        'Get where to copy and paste data (ie row under Counter-Party), this is done by .find to search for names of columns
        'where to paste in CPM sheets
        Fix_Export = Read_in.GetHeaderCell(Fix_Entire_data, "Counter-Party")
        If IsEmpty(Fix_Export) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Flex_Export = Read_in.GetHeaderCell(Flex_Entire_data, "Counter-Party")
        If IsEmpty(Flex_Export) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Flex_CP = Read_in.GetHeaderCell(Flex_ws, "Counter-Party")
        If IsEmpty(Flex_CP) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Flex_CID = Read_in.GetHeaderCell(Flex_ws, "Customer Number")
        If IsEmpty(Flex_CID) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Fix_CP = Read_in.GetHeaderCell(Fix_ws, "Counter-Party")
        If IsEmpty(Fix_CP) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Fix_CID = Read_in.GetHeaderCell(Fix_ws, "Customer Number")
        If IsEmpty(Fix_CID) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Fix_EM = Read_in.GetHeaderCell(Fix_Entire_data, "Expected Margin")
        If IsEmpty(Fix_EM) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Flex_EM = Read_in.GetHeaderCell(Flex_Entire_data, "Expected Margin")
        If IsEmpty(Flex_EM) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Flex_Year1 = Read_in.GetHeaderCell(Flex_Entire_data, "Year 1")
        If IsEmpty(Flex_Year1) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        Fix_Year1 = Read_in.GetHeaderCell(Flex_Entire_data, "Year 1")
        If IsEmpty(Fix_Year1) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
        
        'Clear data before copy and paste
        'we do not clear expected margin (EM) column as it has embedded formulas
        Flex_Entire_data.Range(Flex_Entire_data.Cells(Flex_Export(0) + 1, Flex_Export(1)), Flex_Entire_data.Cells(100000, Flex_EM(1) - 1)).ClearContents
        Fix_Entire_data.Range(Fix_Entire_data.Cells(Fix_Export(0) + 1, Fix_Export(1)), Fix_Entire_data.Cells(100000, Fix_EM(1) - 1)).ClearContents
        Fix_less10_ws.Range(Fix_less10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)), Fix_less10_ws.Cells(100000, Fix_EM(1) - 1)).ClearContents
        Flex_less10_ws.Range(Flex_less10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)), Flex_less10_ws.Cells(100000, Flex_EM(1) - 1)).ClearContents
        
        'Do the copy and pasting for the entire flex/fix data sets and adjust contract number column using .autofill
        'Flex Data
        lastRow_flex = Flex_ws.Cells(Rows.Count, 1).End(xlUp).row
        Set Flex_Range = Flex_ws.Range(Flex_ws.Cells(Flex_CP(0) + 1, Flex_CP(1)), Flex_ws.Cells(lastRow_flex, Flex_CID(1)))
        Flex_Range.Copy
        Flex_Entire_data.Cells(Flex_Export(0) + 1, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Flex_less10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Adjust contract number column (assume contract number is column A and will atleast have atleast 3 numbers in it(which should always be true)
        Flex_Entire_data.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
        Flex_Entire_data.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_Entire_data.Range("A" & Flex_Export(0) + 1 & ":A" & (lastRow_flex + (Flex_Export(0) - Flex_CP(0))))
        Flex_less10_ws.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
        Flex_less10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_less10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & (lastRow_flex + (Flex_Export(0) - Flex_CP(0))))
        
        'Fix Data
        lastRow_fix = Fix_ws.Cells(Rows.Count, 1).End(xlUp).row
        Set Fix_Range_1 = Fix_ws.Range(Fix_ws.Cells(Fix_CP(0) + 1, Fix_CP(1)), Fix_ws.Cells(lastRow_fix, Fix_CID(1)))
        Fix_Range_1.Copy
        Fix_Entire_data.Cells(Fix_Export(0) + 1, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Fix_less10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Adjust contract number column
        Fix_Entire_data.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
        Fix_Entire_data.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_Entire_data.Range("A" & Fix_Export(0) + 1 & ":A" & (lastRow_fix + (Fix_Export(0) - Fix_CP(0))))
        Fix_less10_ws.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
        Fix_less10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_less10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & (lastRow_fix + (Fix_Export(0) - Fix_CP(0))))

        'Get counter-parties with total volume > 10GWh, paste rows into respective flex/fix sheets so they can be bespoke priced.
        'clear sheets first
        Fix_Above10_ws.Range(Fix_Above10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)), Fix_Above10_ws.Cells(100000, Fix_EM(1) - 1)).ClearContents
        Flex_Above10_ws.Range(Flex_Above10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)), Flex_Above10_ws.Cells(100000, Flex_EM(1) - 1)).ClearContents
        
        lastRow_CD = CD_ws.Cells(Rows.Count, 2).End(xlUp).row
        flex_fix_column = CD_ws.Range("Flex_Fix").Column
        fix_cutoff_column = CD_ws.Range("M1_").Column - 1
        
        'Assume starting row/ column of all fix/flex sheets are the same (which should always be true)
        fix_i = Fix_Export(0) + 1
        less10_i = Fix_Export(0) + 1
        flex_i = Flex_Export(0) + 1
        volume = 0
        
        'Loop through all (data) rows in Check_data file (as it has sorted all flex/fix contracts together)
        'Was originally using Read In files to export for this but with calculating totoal volume across flex/fix it made sense to use the already sorted and merged check_data
        For i = CD_CP(0) + 1 To lastRow_CD
            'If contract is not being renewed, dont include it when calculating total volume of customer
            If CD_ws.Cells(i, CD_CID(1) - 1) = "Y" Then volume = volume + CD_ws.Cells(i, CD_ws.Range("Year_1").Column).Value2
            'If counter-party name is the same as row below then add one to variable counter number of contracts for single counter-party
            If UCase(CD_ws.Cells(i, CD_CP(1))) = UCase(CD_ws.Cells(i + 1, CD_CP(1))) Then
                contract_count = contract_count + 1
            Else
                'When the next row is not the same counter-party check the total volume of counterparty >10GWh
                If volume >= 10# Then
                'Loop through all contracts the singular counter-party has using 'contract_count' variable
                    For x = i - contract_count To i
                        'Now add contract to its respective flex/fix sheet
                        If CD_ws.Cells(x, flex_fix_column) = "Flex" Then
                            CD_ws.Range(CD_ws.Cells(x, CD_CP(1)), CD_ws.Cells(x, CD_CID(1))).Copy 'Flex_Above10_ws.Cells(flex_i, Flex_Export(1))
                            Flex_Above10_ws.Cells(flex_i, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                            flex_i = flex_i + 1
                        ElseIf CD_ws.Cells(x, flex_fix_column) = "Fix" Then
                            'Because we are copying from check_data, into Read in files we need to consider the gap for the closed/open volumes
                            CD_ws.Range(CD_ws.Cells(x, CD_CP(1)), CD_ws.Cells(x, fix_cutoff_column)).Copy
                            Fix_Above10_ws.Cells(fix_i, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                            CD_ws.Range(CD_ws.Cells(x, fix_cutoff_column + 360 + 1), CD_ws.Cells(i, CD_CID(1))).Copy
                            Fix_Above10_ws.Cells(fix_i, (Fix_Export(1) - CD_CP(1)) + fix_cutoff_column + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                            fix_i = fix_i + 1
                        End If
                    Next
                End If
                'reset volume and contract counter when new counter-party
                volume = 0
                contract_count = 0
            End If
        Next
        
        'adjust contract number (assume contract number is column A and will atleast have atleast 3 numbers in it(which should always be true)
        fix_lastRow_above10 = Fix_Above10_ws.Cells(Rows.Count, 2).End(xlUp).row
        Fix_Above10_ws.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
        Fix_Above10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_Above10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & fix_lastRow_above10)
        flex_lastRow_above10 = Flex_Above10_ws.Cells(Rows.Count, 2).End(xlUp).row
        Flex_Above10_ws.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
        Flex_Above10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_Above10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & flex_lastRow_above10)

        ReadInWorkbook.Close SaveChanges:=False
        
        'I think leaving the CPM Workbook open would be nice to see the exported data straight away
        'CPMWorkbook.Close Savechanges:=False

        ExcelProcesses (True)

        MsgBox "Export succcesful!"
    Else
        MsgBox "Please check the format of the data using the 'Check Format' button before exporting."
    End If
    

End Sub

Sub Callibration_Export()
    'Basically just export entire portfolio to every sheet in CPM, no above 10 GWh
   If safeToExport = "yes" Then
        Dim answer As Integer
        answer = MsgBox("This exported data is only for callibration, do you want to continue?", vbQuestion + vbYesNo)
        If answer = vbYes Then
        
            MsgBox "Use the 'Export to CPM' button after callibrations are done to load the correct data into the CPM before saving it."
        
            Dim ReadIn_directory As String 'Directory for updated portfolio data that we are reading data in from to copy
            Dim Export_directory As String 'Direcrory for CPM excel model to paste the read in data
            Dim CPMWorkbook As Workbook 'Uses string directory above to open the CPM workbook
            Dim ReadInWorkbook As Workbook 'Uses string directory above to open the updated portfolio data workbook
            Dim Fix_ws As Worksheet 'Read in workbook has a fix data sheet
            Dim Flex_ws As Worksheet 'Read in workbook has a flex data sheet
            Dim Fix_Entire_data As Worksheet 'Sheet in the CPM that stores all fixed data
            Dim Flex_Entire_data As Worksheet 'Sheet in the CPM that stores all flex data
            Dim Fix_less10_ws As Worksheet 'Sheet in the CPM that stores all fixed data used to prices <10GWh clusters
            Dim Flex_less10_ws As Worksheet 'Sheet in the CPM that stores all flex data used to prices <10GWh clusters
            Dim Fix_Above10_ws As Worksheet 'Sheet in the CPM that stores all fixed data above 10GWh to be bespoke priced
            Dim Flex_Above10_ws As Worksheet 'Sheet in the CPM that stores all flex data above 10GWh to be bespoke priced
            Dim Fix_Export As Variant 'Cell in the CPM sheet where fixed read in data will be pasted (exported)
            Dim Flex_Export As Variant 'Cell in the CPM sheet where flex read in data will be pasted (exported)
            Dim Flex_CP As Variant 'Counter-Party (CP) Cell in flex sheet
            Dim Flex_CID As Variant 'Customer ID (CID) Cell in flex sheet
            Dim Fix_CP As Variant 'Counter-Party (CP) Cell in fix sheet
            Dim Fix_CID As Variant 'Customer ID (CID) Cell in fix sheet
            Dim Fix_EM As Variant 'Expected Margin (EM) Cell in fix sheet
            Dim Flex_EM As Variant 'Expected Margin (EM) Cell in flex sheet
            Dim Flex_Year1 As Variant 'Year 1 Cell in flex sheet
            Dim Fix_Year1 As Variant 'Year 1 Cell in fix sheet
            Dim lastRow_flex As Integer 'Number of rows in flex sheet from read in, updated portfolio data, workbook
            Dim Flex_Range As Range 'Range where flex data is in the flex sheet
            Dim lastRow_fix As Integer 'Number of rows in fix sheet from read in, updated portfolio data, workbook
            Dim Fix_Range_1 As Range 'The first range we are copying from fix data sheet up until the fix_cutoff_column

            ExcelProcesses (False)
            
            'Clear old Error Logs
            k = 0
            Sheet2.Range("Error_Log").Offset(1, 0).Resize(100000, 1).ClearContents
            
            'Get directories from the excel sheet, highlighted in yellow and green
            ReadIn_directory = Range("Read_In_Directory").Value
            ReadIn_directory = Replace(ReadIn_directory, Chr(34), vbNullString) ' removes speech marks if it has any from path
            Export_directory = Range("Export_Directory").Value
            Export_directory = Replace(Export_directory, Chr(34), vbNullString)
        
            'Open up CPM Model as a Workbook
            On Error Resume Next
            Set CPMWorkbook = Workbooks.Open(Export_directory)
            If Err.Number <> 0 Then
                MsgBox "Could not find the file " & Export_directory
                Exit Sub
            End If
            On Error GoTo 0
            
            'Open ReadIn workbook (ie the latest portfolio available)
            On Error Resume Next
            Set ReadInWorkbook = Workbooks.Open(ReadIn_directory, ReadOnly:=True)
            If Err.Number <> 0 Then
                MsgBox "Could not find the file " & ReadIn_directory
                Exit Sub
            End If
            On Error GoTo 0
    
            'Open (almost) all required WorkSheets from both directories above and set them to variables
            'Read in sheets
            If Read_in.WorksheetExists("Flex", ReadInWorkbook) = False Then Exit Sub
            Set Flex_ws = ReadInWorkbook.Worksheets("Flex")
            If Read_in.WorksheetExists("Fix", ReadInWorkbook) = False Then Exit Sub
            Set Fix_ws = ReadInWorkbook.Worksheets("Fix")
            
            'Export Sheets
            If Read_in.WorksheetExists("Entire_Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
            Set Fix_Entire_data = CPMWorkbook.Worksheets("Entire_Portfolio_Data_Fixed")
            If Read_in.WorksheetExists("Entire_Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
            Set Flex_Entire_data = CPMWorkbook.Worksheets("Entire_Portfolio_Data_Flex")
            
            If Read_in.WorksheetExists("Less10_Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
            Set Fix_less10_ws = CPMWorkbook.Worksheets("Less10_Portfolio_Data_Fixed")
            If Read_in.WorksheetExists("Less10_Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
            Set Flex_less10_ws = CPMWorkbook.Worksheets("Less10_Portfolio_Data_Flex")
            
            If Read_in.WorksheetExists("Portfolio_Data_Fixed", CPMWorkbook) = False Then Exit Sub
            Set Fix_Above10_ws = CPMWorkbook.Worksheets("Portfolio_Data_Fixed")
            If Read_in.WorksheetExists("Portfolio_Data_Flex", CPMWorkbook) = False Then Exit Sub
            Set Flex_Above10_ws = CPMWorkbook.Worksheets("Portfolio_Data_Flex")

            'Get where to copy and paste data (ie row under Counter-Party), this is done by .find to search for names of columns
            'where to paste in CPM sheets
            Fix_Export = Read_in.GetHeaderCell(Fix_Entire_data, "Counter-Party")
            If IsEmpty(Fix_Export) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Flex_Export = Read_in.GetHeaderCell(Flex_Entire_data, "Counter-Party")
            If IsEmpty(Flex_Export) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Flex_CP = Read_in.GetHeaderCell(Flex_ws, "Counter-Party")
            If IsEmpty(Flex_CP) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Flex_CID = Read_in.GetHeaderCell(Flex_ws, "Customer Number")
            If IsEmpty(Flex_CID) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Fix_CP = Read_in.GetHeaderCell(Fix_ws, "Counter-Party")
            If IsEmpty(Fix_CP) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Fix_CID = Read_in.GetHeaderCell(Fix_ws, "Customer Number")
            If IsEmpty(Fix_CID) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Fix_EM = Read_in.GetHeaderCell(Fix_Entire_data, "Expected Margin")
            If IsEmpty(Fix_EM) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Flex_EM = Read_in.GetHeaderCell(Flex_Entire_data, "Expected Margin")
            If IsEmpty(Flex_EM) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Flex_Year1 = Read_in.GetHeaderCell(Flex_Entire_data, "Year 1")
            If IsEmpty(Flex_Year1) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            Fix_Year1 = Read_in.GetHeaderCell(Flex_Entire_data, "Year 1")
            If IsEmpty(Fix_Year1) Then Exit Sub: ReadInWorkbook.Close SaveChanges:=False
            
            'Clear data before copy and paste
            'we do not clear expected margin (EM) column as it has embedded formulas
            Flex_Entire_data.Range(Flex_Entire_data.Cells(Flex_Export(0) + 1, Flex_Export(1)), Flex_Entire_data.Cells(100000, Flex_EM(1) - 1)).ClearContents
            Fix_Entire_data.Range(Fix_Entire_data.Cells(Fix_Export(0) + 1, Fix_Export(1)), Fix_Entire_data.Cells(100000, Fix_EM(1) - 1)).ClearContents
            Fix_less10_ws.Range(Fix_less10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)), Fix_less10_ws.Cells(100000, Fix_EM(1) - 1)).ClearContents
            Flex_less10_ws.Range(Flex_less10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)), Flex_less10_ws.Cells(100000, Flex_EM(1) - 1)).ClearContents
            Fix_Above10_ws.Range(Fix_Above10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)), Fix_Above10_ws.Cells(100000, Fix_EM(1) - 1)).ClearContents
            Flex_Above10_ws.Range(Flex_Above10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)), Flex_Above10_ws.Cells(100000, Flex_EM(1) - 1)).ClearContents
          
            'Do the copy and pasting for the entire flex/fix data sets and adjust contract number column using .autofill
            'Flex Data
            lastRow_flex = Flex_ws.Cells(Rows.Count, 1).End(xlUp).row
            Set Flex_Range = Flex_ws.Range(Flex_ws.Cells(Flex_CP(0) + 1, Flex_CP(1)), Flex_ws.Cells(lastRow_flex, Flex_CID(1)))
            Flex_Range.Copy
            Flex_Entire_data.Cells(Flex_Export(0) + 1, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            Flex_less10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            Flex_Above10_ws.Cells(Flex_Export(0) + 1, Flex_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            'Adjust contract number column (assume contract number is column A and will atleast have atleast 3 numbers in it(which should always be true)
            Flex_Entire_data.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
            Flex_Entire_data.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_Entire_data.Range("A" & Flex_Export(0) + 1 & ":A" & (lastRow_flex + (Flex_Export(0) - Flex_CP(0))))
            Flex_less10_ws.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
            Flex_less10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_less10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & (lastRow_flex + (Flex_Export(0) - Flex_CP(0))))
            Flex_Above10_ws.Range("A" & Flex_Export(0) + 3 & ":A" & 100000).Clear
            Flex_Above10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & Flex_Export(0) + 2).AutoFill Destination:=Flex_Above10_ws.Range("A" & Flex_Export(0) + 1 & ":A" & (lastRow_flex + (Flex_Export(0) - Flex_CP(0))))
            
            'Fix Data
            lastRow_fix = Fix_ws.Cells(Rows.Count, 1).End(xlUp).row
            Set Fix_Range_1 = Fix_ws.Range(Fix_ws.Cells(Fix_CP(0) + 1, Fix_CP(1)), Fix_ws.Cells(lastRow_fix, Fix_CID(1)))
            Fix_Range_1.Copy
            Fix_Entire_data.Cells(Fix_Export(0) + 1, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            Fix_less10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            Fix_Above10_ws.Cells(Fix_Export(0) + 1, Fix_Export(1)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            'Adjust contract number column
            Fix_Entire_data.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
            Fix_Entire_data.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_Entire_data.Range("A" & Fix_Export(0) + 1 & ":A" & (lastRow_fix + (Fix_Export(0) - Fix_CP(0))))
            Fix_less10_ws.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
            Fix_less10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_less10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & (lastRow_fix + (Fix_Export(0) - Fix_CP(0))))
            Fix_Above10_ws.Range("A" & Fix_Export(0) + 3 & ":A" & 100000).Clear
            Fix_Above10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & Fix_Export(0) + 2).AutoFill Destination:=Fix_Above10_ws.Range("A" & Fix_Export(0) + 1 & ":A" & (lastRow_fix + (Fix_Export(0) - Fix_CP(0))))

            ReadInWorkbook.Close SaveChanges:=False
        
            ExcelProcesses (True)
            
            MsgBox "Export of Callibration succesful"
        
        End If
        
    Else
        MsgBox "Please check the format of the data using the 'Check Format' button before exporting."
    End If

End Sub


Sub Export_Errors()
    'If errors are flagged after checking format of data this allows user to export the errors into a new workbook so they can send it to whoever can sort the problem
    'This will be utilsied using a button that pops up if check_data sub returns atleast one error
    
    Dim err_wb As Workbook
    Dim ReadIn_directory As String
    Dim directory As String
    Dim lastPathIndex As Integer
    Dim lastDotIndex As Integer
    Dim exists As Boolean
    
    ' Get file name from read in directory
    ReadIn_directory = ThisWorkbook.Sheets("Input").Range("Read_In_Directory")
    ' Get just the file name and extension
    lastPathIndex = InStrRev(ReadIn_directory, Application.PathSeparator)
    If lastPathIndex >= 1 Then ReadIn_directory = Right(ReadIn_directory, Len(ReadIn_directory) - lastPathIndex) ': directory = Left(ReadIn_directory, Len(ReadIn_directory) - lastPathIndex)
    ' Now get the file name without the extension
    lastDotIndex = InStrRev(ReadIn_directory, ".")
    If lastDotIndex >= 1 Then ReadIn_directory = Left(ReadIn_directory, lastDotIndex - 1)

    ReadIn_directory = Application.ActiveWorkbook.Path & Application.PathSeparator & ReadIn_directory & " Errors.xlsx"

    ' Create new workbook and add 2 sheets to it
    ' Check if a wb already exists with the same name
    On Error Resume Next
    Set err_wb = Workbooks.Open(ReadIn_directory)
    If Err.Number <> 0 Then
        ' If it doesnt exist (and throws an error) then create new workbook
        Set err_wb = Workbooks.Add
        err_wb.Worksheets.Add Count:=1
        err_wb.Worksheets(1).name = "Data"
        err_wb.Worksheets(2).name = "Error Log"
        exists = False
    Else
        ' Clear error workbook (obvously does nothing if we just created it)
        err_wb.Sheets("Data").UsedRange.Clear
        err_wb.Sheets("Error Log").UsedRange.Clear
        exists = True
        End If
    On Error GoTo 0
    
    ' Copy and paste
    ThisWorkbook.Sheets("Input").UsedRange.Offset(2, 0).Copy err_wb.Sheets("Data").Cells(2, 1)
    ThisWorkbook.Sheets("Error_Log").UsedRange.Copy err_wb.Sheets("Error Log").Cells(1, 1)
    
    ' Adjusts column widths
    err_wb.Sheets("Data").Cells.EntireColumn.ColumnWidth = 12
    err_wb.Sheets("Data").Columns(2).ColumnWidth = 30
    err_wb.Sheets("Error Log").Cells.EntireColumn.AutoFit

    
    ' Saves wb
    If exists = False Then
        err_wb.SaveAs fileName:=ReadIn_directory
    ElseIf exists = True Then
        err_wb.Save
    End If
    
End Sub

Function ExtractMonthYear(filePath As String) As String
    Dim fileName As String
    Dim parts() As String
    Dim nameParts() As String
    
    ' Extract the file name from the full path
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    ' Remove the extension
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Split by "-" to remove extra text
    parts = Split(fileName, "-")

    ' Trim spaces and return the relevant part
    nameParts = Split(Trim(parts(0)), " ")
    ExtractMonthYear = nameParts(0) & " " & nameParts(1)
End Function

Sub SetUsedColumnsWidthAndHeight(ws As Worksheet, colWidth As Double, colHeight As Double)
    Dim lastCol As Integer
    Dim colLetter As String
    
    ' Find the last used column
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                
    ' Convert column number to letter
    colLetter = Split(ws.Cells(1, lastCol).Address, "$")(1)

    ' Apply column width to all used columns
    ws.Columns("A:" & colLetter).ColumnWidth = colWidth
    ws.Rows(1).RowHeight = colHeight
End Sub
