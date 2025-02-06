Attribute VB_Name = "Sort_FurtherChecks"
Option Explicit

Sub Sort(from_cell As String, to_cell As String, ws As Worksheet)

    ws.Range(from_cell, ws.Range(to_cell).End(xlDown)).Sort Key1:=ws.Range(from_cell), Order1:=xlAscending, Header:=xlYes, MatchCase:=True
    ws.Range("Contract_Number", ws.Range("Contract_Number").End(xlDown)).Sort Key1:=ws.Range("Contract_Number"), Order1:=xlAscending, Header:=xlYes

End Sub

Sub Sort_Errors()
    Dim CD_ws As Worksheet 'Check Data (CD) worksheet
    Set CD_ws = ThisWorkbook.Sheets("Input")
    
    CD_ws.Sort.SortFields.Clear
    CD_ws.Sort.SortFields.Add Key:=CD_ws.Range("Counter_Party"), _
        SortOn:=xlSortOnCellColor, Order:=xlDescending, DataOption:=xlSortNormal
    
    With CD_ws.Sort
    .SetRange Range("Contract_Number", Range("Cust_ID").End(xlDown))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
    
End Sub

Sub UnSort()

    Dim CD_ws As Worksheet 'Check Data (CD) worksheet
    Set CD_ws = ThisWorkbook.Sheets("Input")
    
    CD_ws.Sort.SortFields.Clear
    CD_ws.Sort.SortFields.Add Key:=CD_ws.Range("Contract_Number"), _
        Order:=xlAscending, DataOption:=xlSortNormal
    
    With CD_ws.Sort
    .SetRange Range("Contract_Number", Range("Cust_ID").End(xlDown))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With

End Sub

Sub Highlight_Entire_Row(from_column As Integer, to_column As Integer, row As Long, ws As Worksheet, include_row_above As Boolean, Optional Color As Long = 0)
    Dim c As Integer
    If Color = 0 Then Color = RGB(255, 192, 203)
    For c = from_column To to_column
        If ws.Cells(row, c).Interior.ColorIndex <> 3 Then ws.Cells(row, c).Interior.Color = Color
        If include_row_above = True Then
            If ws.Cells(row - 1, c).Interior.ColorIndex = xlColorIndexNone Then ws.Cells(row - 1, c).Interior.Color = Color
        End If
    Next
End Sub

Sub Fix_Duplicates()

    Dim CD_ws As Worksheet 'Check Data (CD) worksheet
    Set CD_ws = ThisWorkbook.Sheets("Input")

    'Resort by name
    Call Sort("Counter_Party", "Cust_ID", CD_ws)
    
    Dim Total_Duplicate_Index As Object
    Dim outerList As Object
    Dim innerList As Object
    Dim outerIndex As Long, innerIndex As Long
    Dim lastCol As Long
    Dim ws As Worksheet
    Dim cellValue As Variant
    
    ' Initialize the outer ArrayList
    Set Total_Duplicate_Index = CreateObject("System.Collections.ArrayList")
    
    ' Reference the target worksheet
    Set ws = ThisWorkbook.Sheets("Duplicate")
    
    If ws.Range("A1") = "" Then MsgBox "No contracts found to fix. Re-check the format if needed.": Exit Sub
    
    ' Loop through each row in the used range
    For outerIndex = 1 To ws.UsedRange.Rows.Count
        ' Determine the last populated column for the current row
        lastCol = ws.Cells(outerIndex, ws.Columns.Count).End(xlToLeft).Column
        
        ' Initialize the inner ArrayList for the current row
        Set innerList = CreateObject("System.Collections.ArrayList")
        
        ' Loop through the columns up to the last populated column
        For innerIndex = 1 To lastCol
            cellValue = ws.Cells(outerIndex, innerIndex).Value
            innerList.Add cellValue
        Next innerIndex
        
        ' Add the inner ArrayList to the outer ArrayList
        Total_Duplicate_Index.Add innerList
    Next outerIndex
    
    Set ws = ThisWorkbook.Sheets("Input")
    
    Call Duplicate_IDs(Total_Duplicate_Index, ws)
    
'    Dim r As Long
'    For outerIndex = 0 To Total_Duplicate_Index.Count - 1
'        For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1
'            r = Total_Duplicate_Index(outerIndex)(innerIndex)
'            ws.Cells(r, ws.Range("Multiple?").Column).Value = "Y"
'        Next innerIndex
'    Next outerIndex
    
    ' Call check_data.check_data
    
    Dim r As Long
    For outerIndex = 0 To Total_Duplicate_Index.Count - 1
        For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1
            r = Total_Duplicate_Index(outerIndex)(innerIndex)
            ' ws.Cells(r, ws.Range("Multiple?").Column).Value = "Y"
            Call Highlight_Entire_Row(ws.Range("Counter_Party").Column, ws.Range("Cust_ID").Column, r, CD_ws, False, RGB(218, 247, 166))
        Next innerIndex
    Next outerIndex
    
    ws.Range("Modified?").Value = "Y"
    
    MsgBox "Contracts with multiple IDs now aligned. Check format to double check."

End Sub

Sub Duplicate_IDs(Total_Duplicate_Index As Object, ws As Worksheet)

    Dim PD_Dict As Object
    Set PD_Dict = Create_PD_Dict()

    Dim Ratings As Object
    Dim PDs As Object
    Dim SDs As Object
    Dim SEinsurances As Object
    Dim CEEinsurances As Object
    Dim Customer_IDs As Object
    Dim Rating As Variant
    Dim Lowest_PD As Variant
    Dim Lowest_PD_Index As Long
    Dim outerIndex As Long, innerIndex As Long
    
    For outerIndex = 0 To Total_Duplicate_Index.Count - 1 ' loop through customers
        Set Ratings = CreateObject("System.Collections.ArrayList")
        Set PDs = CreateObject("System.Collections.ArrayList")
        Set SDs = CreateObject("System.Collections.ArrayList")
        Set SEinsurances = CreateObject("System.Collections.ArrayList")
        Set CEEinsurances = CreateObject("System.Collections.ArrayList")
        Set Customer_IDs = CreateObject("System.Collections.ArrayList")
        For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1 ' loop through contracts indexes for that customer
            ' Want to get lowest PD, Highest SD and Highest Insurance
            Ratings.Add ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Credit_Rating").Column).Value2
            SDs.Add ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Security_Deposit").Column).Value2
            SEinsurances.Add ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("SE_Insured_Amount").Column).Value2
            CEEinsurances.Add ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("CEE_Insured_Amount").Column).Value2
            Customer_IDs.Add ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Cust_ID").Column).Value
        Next innerIndex

        Lowest_PD = PD_Dict(Ratings(0))
        Lowest_PD_Index = 0
        For Each Rating In Ratings
            PDs.Add PD_Dict(Rating)
            If PD_Dict(Rating) < Lowest_PD Then
                Lowest_PD = PD_Dict(Rating)
                Lowest_PD_Index = Ratings.IndexOf(Rating, 0)
            End If
        Next Rating
        SDs.Sort
        SDs.Reverse
        SEinsurances.Sort
        SEinsurances.Reverse
        CEEinsurances.Sort
        CEEinsurances.Reverse
        
        For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1 ' loop through contracts indexes for that customer
            ' Set all parameters to values found above
            ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Credit_Rating").Column).Value = Ratings.Item(Lowest_PD_Index) 'PDs.Sort.Item(PDs.Count)
            ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Security_Deposit").Column).Value = SDs.Item(0) 'SDs.Sort.Item(SDs.Count)
            ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("SE_Insured_Amount").Column).Value = SEinsurances.Item(0)
            ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("CEE_Insured_Amount").Column).Value = SEinsurances.Item(0)
            ws.Cells(Total_Duplicate_Index(outerIndex)(innerIndex), ws.Range("Cust_ID").Column).Value = "'" & Customer_IDs.Item(0) & "" ' assign first ID it comes across
            Call Highlight_Entire_Row(ws.Range("Counter_Party").Column, ws.Range("Cust_ID").Column, (Total_Duplicate_Index(outerIndex)(innerIndex)), ws, False, RGB(218, 247, 166))
        Next innerIndex
    Next outerIndex
    

End Sub

Sub Further_checks() 'Optional ByVal DuplicatesFixed As Boolean = False)
    'This Function checks that infomation is conistent across contracts from the same Counterparty.
    
    Dim CD_ws As Worksheet 'Check Data (CD) worksheet
    Set CD_ws = ThisWorkbook.Sheets("Input")
    
    'Sorts all columns by sorting Counter-Party name alphabetically, data needs to be sorted for checks to work
    Call Sort("Counter_Party", "Cust_ID", CD_ws)
    
    'Dim k As Long
    Dim lastRow As Long
    Dim i As Long
    Dim x As Long
    
    ThisWorkbook.Sheets("Duplicate").Range("A1:ZZ1000").ClearContents
    
    'using Named ranges here now
    'Counterparty cell is the first Column of interest, Customer ID is last Column of interest
    Dim CP_Row As Integer 'all headers on same row so just use this
    Dim CP_Column As Integer
    Dim CID_Column As Integer
    
    CP_Row = CD_ws.Range("Counter_Party").row
    CP_Column = CD_ws.Range("Counter_Party").Column
    CID_Column = CD_ws.Range("Cust_ID").Column
    lastRow = Cells(Rows.Count, CP_Column).End(xlUp).row
    
    'Removes any 'error shading' from previous check formats
    CD_ws.Range(CD_ws.Cells(CP_Row + 1, CP_Column), CD_ws.Cells(lastRow, CID_Column)).Interior.ColorIndex = xlNone
    
    Dim Start_Mit_Column As Integer
    Dim End_Mit_Column As Integer
    Dim Unsecure_SE_Days As Range
    Dim Commodity_Column As Integer
    Dim Import_Export_Column As Integer
    Dim Fix_Flex_Column As Integer
    Dim LGD_Columns As Range
    Dim Start_LGD_Column As Integer
    Dim End_LGD_Column As Integer
    Dim Uplift_Column As Integer
    Dim Aged_Debt_Column As Integer
    Dim SD_Column As Integer
    Dim SE_Amount_Column As Integer
    Dim CEE_Amount_Column As Integer
    Dim Credit_Column As Integer
    
    Start_Mit_Column = CD_ws.Range("Mitigations").Cells(1, 1).Column
    End_Mit_Column = Range("Mitigations").Columns.Count + Start_Mit_Column - 1
    Set Unsecure_SE_Days = Range("Unsecured_Settlements_Exposure__Days")
    Commodity_Column = CD_ws.Range("Commodity").Column
    Import_Export_Column = CD_ws.Range("Import_Export").Column
    Fix_Flex_Column = CD_ws.Range("Flex_Fix").Column
    Start_LGD_Column = CD_ws.Range("Loss_Given_Default").Cells(1, 1).Column
    End_LGD_Column = CD_ws.Range("Loss_Given_Default").Columns.Count + Start_LGD_Column - 1
    Set LGD_Columns = CD_ws.Range("Loss_Given_Default")
    Uplift_Column = CD_ws.Range("Uplift").Column
    Aged_Debt_Column = CD_ws.Range("Aged_Debt").Column
    SD_Column = CD_ws.Range("Security_Deposit").Column
    SE_Amount_Column = CD_ws.Range("SE_Insured_Amount").Column
    CEE_Amount_Column = CD_ws.Range("CEE_Insured_Amount").Column
    Credit_Column = CD_ws.Range("Credit_Rating").Column
    
    Dim Error_On_Line As Boolean
    Dim Row_Above As Boolean
    
    Dim Duplicate_Name As Object
    Dim Total_Duplicate_Index As Object
    Dim Duplicate_Index As Object
    Set Duplicate_Name = CreateObject("System.Collections.ArrayList")
    Set Total_Duplicate_Index = CreateObject("System.Collections.ArrayList")
    Set Duplicate_Index = CreateObject("System.Collections.ArrayList")
    Dim tempIndex As Object
    Dim Duplicate_Bool As Boolean
    Dim n As Long

    'Loop from row CP_Row+1 , CP_Row is the header row
    For i = CP_Row + 1 To lastRow
    
        Error_On_Line = False
        Row_Above = False
    
        'Cant crosscheck row above with the header row
        If i <> CP_Row + 1 Then
            'Check Counter-Party name with row above, and if they match
            If UCase(Cells(i, CP_Column).Value) = UCase(Cells(i - 1, CP_Column).Value) Then
                'If i = CP_Row + 2 Then Duplicate_Index.Add i - 1
                Duplicate_Index.Add i
                'Check if Customer ID matches:
                If Cells(i, CID_Column).Value <> Cells(i - 1, CID_Column).Value Then 'And CD_ws.Cells(i, CD_ws.Range("Multiple?").Column).Value <> "Y" Then
                    Duplicate_Bool = True
                    Duplicate_Name.Add Cells(i, CP_Column).Value
'                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
'                        "ID not matching across Contracts from " & Cells(i, CP_Column).Value & ". Contracts: " & i - CP_Row - 1 & " and " & i - CP_Row & "."
                    'Error_On_Line = True
                    'Row_Above = True
'                    'Range(Cells(i, CP_Column), Cells(i - 1, CID_Column)).Interior.Color = RGB(255, 192, 203)
                    Range(Cells(i, CID_Column), Cells(i - 1, CID_Column)).Interior.ColorIndex = 3
'                    k = k + 1
                End If ' Else
                'Check if Mitigation parameters match:
                For x = Start_Mit_Column To End_Mit_Column
                    If Cells(i, x).Value <> Cells(i - 1, x).Value Then
                        Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                            Cells(CP_Row, x) & " not matching across Contracts from " & Cells(i, CP_Column).Value & ". Contracts: " & i - CP_Row - 1 & " and " & i - CP_Row & "."
                        Error_On_Line = True
                        Row_Above = True
                        'Range(Cells(i, CP_Column), Cells(i - 1, CID_Column)).Interior.Color = RGB(255, 192, 203)
                        Range(Cells(i, x), Cells(i - 1, x)).Interior.ColorIndex = 3
                        k = k + 1
                    End If
                Next x
                'Check credit rating
                If Cells(i, Credit_Column).Value <> Cells(i - 1, Credit_Column).Value Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                        Cells(CP_Row, Credit_Column) & " not matching across Contracts from " & Cells(i, CP_Column).Value & ". Contracts: " & i - CP_Row - 1 & " and " & i - CP_Row & "."
                    Error_On_Line = True
                    Row_Above = True
                    'Range(Cells(i, CP_Column), Cells(i - 1, CID_Column)).Interior.Color = RGB(255, 192, 203)
                    Range(Cells(i, Credit_Column), Cells(i - 1, Credit_Column)).Interior.ColorIndex = 3
                    k = k + 1
                End If
                ' End If
            Else
            
                If Duplicate_Bool = True Then
                    Duplicate_Index.Add Duplicate_Index(0) - 1
                    Duplicate_Index.Sort
                    Set tempIndex = CreateObject("System.Collections.ArrayList")
                    For n = 0 To Duplicate_Index.Count - 1
                        tempIndex.Add Duplicate_Index(n)
                    Next n
                    Total_Duplicate_Index.Add tempIndex
                    Duplicate_Bool = False
                    
                End If
                
                Duplicate_Index.Clear
            
            End If
            ' Incase the last row has the different customer numbers
            If i = lastRow Then
                If Duplicate_Bool = True Then
                    Duplicate_Index.Add Duplicate_Index(0) - 1
                    Duplicate_Index.Sort
                    Set tempIndex = CreateObject("System.Collections.ArrayList")
                    For n = 0 To Duplicate_Index.Count - 1
                        tempIndex.Add Duplicate_Index(n)
                    Next n
                    Total_Duplicate_Index.Add tempIndex
                    Duplicate_Bool = False
                    
                End If
                
                Duplicate_Index.Clear
            End If
                
        End If
        
        
        
        
        ' Check LGD = 0 (loop through 3 columns)
        For x = Start_LGD_Column To End_LGD_Column
            If Cells(i, x).Value <> 0 Then
                Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, x) & ", must be 0"
                Error_On_Line = True
                'Range(Cells(i, CP_Column), Cells(i - 1, CID_Column)).Interior.Color = RGB(255, 192, 203)
                Range(Cells(i, x), Cells(i, x)).Interior.ColorIndex = 3
                k = k + 1
            End If
        Next x
        
        'Checks unsecured SE days
        If CD_ws.Cells(i, Unsecure_SE_Days.Cells(1, 1).Column).Value < 28 Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Unsecure_SE_Days.name.name & ", must be at least 28 days"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Unsecure_SE_Days.Cells(1, 1).Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check Aged Debt = 0
        If CD_ws.Cells(i, Aged_Debt_Column).Value <> 0# Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, Aged_Debt_Column) & ", must be £0"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Aged_Debt_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check Uplift > 1
        If CD_ws.Cells(i, Uplift_Column).Value < 1# Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, Uplift_Column) & ", must be >= 1"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Uplift_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check Flex_Fix = Flex, Fix, Deemed or Default
        If CD_ws.Cells(i, Fix_Flex_Column).Value <> "Fix" And CD_ws.Cells(i, Fix_Flex_Column).Value <> "Flex" And CD_ws.Cells(i, Fix_Flex_Column).Value <> "Deemed" And CD_ws.Cells(i, Fix_Flex_Column).Value <> "Default" Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, Fix_Flex_Column) & ", must be either 'Fix', 'Flex', 'Deemed' or 'Default'"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Fix_Flex_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check Commodity = P or G
        If CD_ws.Cells(i, Commodity_Column).Value <> "P" And CD_ws.Cells(i, Commodity_Column).Value <> "G" Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, Commodity_Column) & ", must be either 'P' or 'G'"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Commodity_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check SD >= 0
        If CD_ws.Cells(i, SD_Column).Value < 0# Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, SD_Column) & ", must be greater than £0"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, SD_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check SE insured amount >= 0
        If CD_ws.Cells(i, SE_Amount_Column).Value < 0# Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, SE_Amount_Column) & ", must be greater than £0"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, SE_Amount_Column).Interior.ColorIndex = 3
            k = k + 1
        End If
        
        ' Check SE insured amount >= 0
        If CD_ws.Cells(i, CEE_Amount_Column).Value < 0# Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, CEE_Amount_Column) & ", must be greater than £0"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, CEE_Amount_Column).Interior.ColorIndex = 3
            k = k + 1
        End If

        ' Check import/export
        If CD_ws.Cells(i, Import_Export_Column).Value <> "import" And CD_ws.Cells(i, Import_Export_Column).Value <> "export" Then
            Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "Contract: " & i - CP_Row & ", Column: " & Cells(CP_Row, Import_Export_Column) & ", must be either 'import' or 'export'"
            Error_On_Line = True
            'Range(Cells(i, CP_Column), Cells(i, CID_Column)).Interior.Color = RGB(255, 192, 203)
            Cells(i, Import_Export_Column).Interior.ColorIndex = 3
            k = k + 1
        End If

        'highlight entire row
        If Error_On_Line = True Then Call Highlight_Entire_Row(CP_Column, CID_Column, i, CD_ws, Row_Above)
        
    Next i
    
    If Total_Duplicate_Index.Count > 0 Then
        ' MsgBox "There are " & Total_Duplicate_Index.Count & " customers with varied customer numbers. They are highlighted in orange. Click the xx button to align them."
        Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                    "There are " & Total_Duplicate_Index.Count & " customers with varied customer numbers. They are highlighted in orange. Click the 'Fix Multiple ID Contracts' button to align them."
        k = k + 1
        
        Dim r As Long
        Dim outerIndex As Long, innerIndex As Long
        For outerIndex = 0 To Total_Duplicate_Index.Count - 1
            For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1
                r = Total_Duplicate_Index(outerIndex)(innerIndex)
                Call Highlight_Entire_Row(CP_Column, CID_Column, r, CD_ws, False, RGB(250, 200, 152))
                ThisWorkbook.Sheets("Duplicate").Cells(outerIndex + 1, innerIndex + 1).Value = Total_Duplicate_Index(outerIndex)(innerIndex)
            Next innerIndex
        Next outerIndex
    End If
    
'    Dim outerIndex As Long, innerIndex As Long
'    For outerIndex = 0 To Total_Duplicate_Index.Count - 1
'        MsgBox "List " & outerIndex + 1 & ":"
'        For innerIndex = 0 To Total_Duplicate_Index(outerIndex).Count - 1
'            MsgBox Total_Duplicate_Index(outerIndex)(innerIndex)
'        Next innerIndex
'    Next outerIndex

    'Call Duplicate_IDs(Total_Duplicate_Index, CD_ws)
    
    
    'Check same customer name with same ID
    'Sort by ID (including contract Number), check names for same ID
    CD_ws.Range("Contract_Number", CD_ws.Range("Cust_ID").End(xlDown)).Sort Key1:=CD_ws.Range("Cust_ID"), Order1:=xlAscending, Header:=xlYes

    For i = CP_Row + 1 To lastRow

        Error_On_Line = False
        Row_Above = False

        'Cant crosscheck row above with the header row
        If i <> CP_Row + 1 Then
            'Check Counter-Party name with row above, and if they match
            If Cells(i, CID_Column).Value = Cells(i - 1, CID_Column).Value Then
                'Check if Customer ID matches:
                If UCase(Cells(i, CP_Column).Value) <> UCase(Cells(i - 1, CP_Column).Value) Then
                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
                        "Name not matching across Contracts with ID: " & Cells(i, CID_Column).Value & ". Contracts: " & Cells(i - 1, CP_Column - 1) & " and " & Cells(i, CP_Column - 1) & "."
                    Error_On_Line = True
                    Row_Above = True
                    Range(Cells(i, CP_Column), Cells(i - 1, CP_Column)).Interior.ColorIndex = 3
                    k = k + 1
                End If
'                If Cells(i, Credit_Column).Value <> Cells(i - 1, Credit_Column).Value Then
'                    Sheet2.Range("Error_Log").Offset(k + 1, 0) = _
'                        "Credit Rating not matching across Contracts with ID: " & Cells(i, CID_Column).Value & ". Contracts: " & Cells(i - 1, CP_Column - 1) & " and " & Cells(i, CP_Column - 1) & "."
'                    Error_On_Line = True
'                    Row_Above = True
'                    Range(Cells(i, Credit_Column), Cells(i - 1, Credit_Column)).Interior.ColorIndex = 3
'                    k = k + 1
'                End If
            End If
        End If
        
        'highlight entire row
        If Error_On_Line = True Then Call Highlight_Entire_Row(CP_Column, CID_Column, i, CD_ws, Row_Above)
        
    Next i

    'Resort by name
    Call Sort("Counter_Party", "Cust_ID", CD_ws)
    
End Sub

Function Create_PD_Dict() As Object
    Dim PD_Dict
    Set PD_Dict = CreateObject("Scripting.Dictionary")
    PD_Dict.Add "AAA", 0.0001
    PD_Dict.Add "AA+", 0.0002
    PD_Dict.Add "AA", 0.0003
    PD_Dict.Add "AA-", 0.0004
    PD_Dict.Add "A+", 0.0006
    PD_Dict.Add "A", 0.0007
    PD_Dict.Add "A-", 0.0008
    PD_Dict.Add "BBB+", 0.001
    PD_Dict.Add "BBB", 0.0014
    PD_Dict.Add "BBB-", 0.0022
    PD_Dict.Add "BB+", 0.0041
    PD_Dict.Add "BB", 0.0066
    PD_Dict.Add "BB-", 0.0126
    PD_Dict.Add "B+", 0.0189
    PD_Dict.Add "B", 0.0292
    PD_Dict.Add "B-", 0.0454
    PD_Dict.Add "CCC", 0.3315
    PD_Dict.Add "_<400", 0.833
    PD_Dict.Add "400to425", 0.7646
    PD_Dict.Add "426to450", 0.696
    PD_Dict.Add "451to475", 0.6257
    PD_Dict.Add "476to500", 0.5379
    PD_Dict.Add "501to525", 0.4489
    PD_Dict.Add "526to550", 0.3714
    PD_Dict.Add "551to575", 0.2958
    PD_Dict.Add "576to600", 0.2299
    PD_Dict.Add "601to625", 0.1767
    PD_Dict.Add "626to650", 0.1325
    PD_Dict.Add "651to675", 0.0972
    PD_Dict.Add "676to700", 0.0726
    PD_Dict.Add "701to725", 0.0505
    PD_Dict.Add "726to750", 0.0376
    PD_Dict.Add "751to775", 0.0265
    PD_Dict.Add "776to800", 0.0189
    PD_Dict.Add "801to825", 0.0133
    PD_Dict.Add "826to850", 0.0095
    PD_Dict.Add "851to875", 0.0068
    PD_Dict.Add "_876+", 0.0049
    
    Set Create_PD_Dict = PD_Dict
    
End Function
