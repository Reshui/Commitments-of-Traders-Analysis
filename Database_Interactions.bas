Attribute VB_Name = "Database_Interactions"
Public DataBase_Not_Found As Boolean

Private AfterEventHolder As ClassQTE

Option Explicit

 Sub database_details(combined_wb_bool As Boolean, Report_Type As String, Optional ByRef cn As Object, Optional ByRef table_name As String, Optional ByRef db_path As String)
'===================================================================================================================
'Determines if database exists. IF it does the appropriate variables or properties are assigned values if needed.
'===================================================================================================================
    Dim Report_Name As String, user_database_path As String ', T As Variant
    
    If Report_Type = "T" Then
        Report_Name = "TFF"
    Else
        Report_Name = Evaluate("VLOOKUP(""" & Report_Type & """,Report_Abbreviation,2,FALSE)")
    End If
    
    If UUID Then
        db_path = Environ$("USERPROFILE") & "\Documents\" & Report_Name & ".accdb"
        DataBase_Not_Found = False
    Else

        user_database_path = Variable_Sheet.Range(Report_Type & "_Database_Path").value
            
        If user_database_path = vbNullString Or Not FileOrFolderExists(user_database_path) Then
        
            DataBase_Not_Found = True
            
            If Data_Retrieval.Running_Weekly_Retrieval Then
                Exit Sub
            Else
                MsgBox Report_Name & " database not found."
                Re_Enable
                End
            End If
            
        Else
            db_path = user_database_path
            DataBase_Not_Found = False
        End If

    End If
    
    If Not IsMissing(table_name) Then table_name = Report_Name & IIf(combined_wb_bool = True, "_Combined", "_Futures_Only")
    
    If Not cn Is Nothing Then cn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & db_path & ";"
    
    'Set T = cn.Properties
    
End Sub
Private Function FilterColumnsAndDelimit(fieldsInDatabase As Variant, Report_Type As String, includePriceColumn As Boolean) As String
'===================================================================================================================
'Loops table found on Variables Worksheet that contains True/False values for wanted columns
'An array of wanted columns with some re-ordering is returned
'===================================================================================================================
    Dim wantedColumns() As Variant
    
    wantedColumns = Filter_Market_Columns(False, True, convert_skip_col_to_general:=False, Report_Type:=Report_Type, Create_Filter:=True, InputA:=fieldsInDatabase)
    
    ReDim Preserve wantedColumns(LBound(wantedColumns) To UBound(wantedColumns) + IIf(includePriceColumn = True, 1, 0))
    
    If includePriceColumn Then wantedColumns(UBound(wantedColumns)) = "Price"

    FilterColumnsAndDelimit = WorksheetFunction.TextJoin(",", True, wantedColumns)
    
End Function

Function FieldsFromRecordSet(record As Object, use_brackets As Boolean) As Variant
'===================================================================================================================
'record is a RecordSET object containing a single row of data from which field names are retrieved,formatted and output as an array
'===================================================================================================================
    Dim x As Integer, Z As Byte, fieldNamesInRecord() As Variant, currentFieldName As String

    ReDim fieldNamesInRecord(1 To record.Fields.Count - 1)
    
    For x = 0 To record.Fields.Count - 1
        
        currentFieldName = record(x).name
        
        If Not currentFieldName = "ID" Then
            Z = Z + 1
            If use_brackets Then
                fieldNamesInRecord(Z) = "[" & currentFieldName & "]"
            Else
                fieldNamesInRecord(Z) = currentFieldName
            End If
            
        End If

    Next x
    
    FieldsFromRecordSet = fieldNamesInRecord

End Function

Function QueryDatabaseForContract(Report_Type As String, combined_wb_bool As Boolean, contract_code As String) As Variant
'===================================================================================================================
'Retrieves filtered data from database and returns as an array
'===================================================================================================================
    Dim record As Object, cn As Object, tableNameWithinDatabase As String

    Dim SQL As String, delimitedWantedColumns As String, allFieldNames() As Variant
    
'    Dim retrievalTimer As TimedTask
'    Set retrievalTimer = New TimedTask: retrievalTimer.Start "Contract Retrieval(" & contract_code & ") ~ " & Time

    On Error GoTo Close_Connection

    Set cn = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, Report_Type, cn, tableNameWithinDatabase

    With cn
        '.CursorLocation = adUseServer
        .Open
        Set record = .Execute(tableNameWithinDatabase, , adCmdTable)
    End With
    
    allFieldNames = FieldsFromRecordSet(record, use_brackets:=True)
    
    record.Close
    
    delimitedWantedColumns = FilterColumnsAndDelimit(allFieldNames, Report_Type:=Report_Type, includePriceColumn:=True)

    SQL = "SELECT " & delimitedWantedColumns & " FROM " & tableNameWithinDatabase & " WHERE [CFTC_Contract_Market_Code]='" & contract_code & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
    
    With record
    
        .Open SQL, cn
         QueryDatabaseForContract = TransposeData(.GetRows)

    End With

    'If Not retrievalTimer Is Nothing Then retrievalTimer.DPrint
    
Close_Connection:

    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Function

Public Sub Update_DataBase(Data_Array As Variant, combined_wb_bool As Boolean, Report_Type As String, debugOnly As Boolean)
'===================================================================================================================
'Uodates a given data table one row at a time
'===================================================================================================================
    Dim table_name As String, field_names() As Variant, x As Long, _
    Row_Data() As Variant, Y As Byte, legacy_combined_table_name As String, _
    Legacy_Combined_Data As Boolean, oldest_added_date As Date
    
    
    'dim  Records_to_Update As Long,
    Dim contract_code As String, SQL As String, legacy_database_path As String
    
    Dim record As Object, cn As Object ',row_date As Date, number_of_records_command As Object, number_of_records_returned As Object
    
    On Error GoTo Close_Connection
    
    Const yyyy_mm_dd_column As Byte = 3, legacy_abbreviation As String = "L"
    'Const contract_code_column As Byte = 4

    ReDim Row_Data(LBound(Data_Array, 2) To UBound(Data_Array, 2))
    
    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    'Set number_of_records_command = CreateObject("ADODB.Command")
    'Set number_of_records_returned = CreateObject("ADODB.RecordSet")

    If Report_Type = legacy_abbreviation And combined_wb_bool = True Then Legacy_Combined_Data = True

    Call database_details(combined_wb_bool, Report_Type, cn, table_name)   'Generates a connection string and assigns a table to modify

    With cn
        '.CursorLocation = adUseServer                                   'Batch update won't work otherwise
        .Open
        Set record = .Execute(CommandText:=table_name, Options:=adCmdTable) 'This record will be used to retrieve field names
    End With
    
    field_names = FieldsFromRecordSet(record, use_brackets:=False)     'Field names from database returned as an array
    record.Close

'    With number_of_records_command
'        'Command will be used to ensure that there aren't duplicate entries in the database
'        .ActiveConnection = cn
'        .CommandText = "SELECT Count([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & " WHERE [Report_Date_as_YYYY-MM-DD] = ? AND [CFTC_Contract_Market_Code] = ?;"
'        .CommandType = adCmdText
'        .Prepared = True
'
'        With .Parameters
'            .Append number_of_records_command.CreateParameter("YYYY-MM-DD", adDate, adParamInput)
'            .Append number_of_records_command.CreateParameter("Contract_Code", adVarWChar, adParamInput, 6)
'        End With
'
'    End With
    
    'Call delete_cftc_data_from_database(CDate(data_array(1, yyyy_mm_dd_column)), Report_Type, combined_wb_bool)

    With record
        'This Recordset will be used to add new data to the database table via the batchupdate method
        .Open table_name, cn, adOpenForwardOnly, adLockOptimistic
        
        oldest_added_date = Data_Array(LBound(Data_Array, 1), yyyy_mm_dd_column)

        For x = LBound(Data_Array, 1) To UBound(Data_Array, 1)
            
            If Not (debugOnly Or Legacy_Combined_Data) Then
                If Data_Array(x, yyyy_mm_dd_column) < oldest_added_date Then oldest_added_date = Data_Array(x, yyyy_mm_dd_column)
            End If
            
            'If data_array(X, yyyy_mm_dd_column) > DateSerial(2000, 1, 1) Then GoTo next_row
            
'            contract_code = data_array(X, contract_code_column)
'            row_date = data_array(X, yyyy_mm_dd_column)
'
'            If Legacy_Combined_Data Then
'                If row_date < oldest_added_date Then oldest_added_date = row_date
'            End If
            
            'number_of_records_command.Parameters("Contract_Code").value = contract_code
            'number_of_records_command.Parameters("YYYY-MM-DD").value = row_date

            'Set number_of_records_returned = number_of_records_command.Execute

            'If number_of_records_returned(0) = 0 Then 'If new row can be uniquely identified with a date and contract code
                
                'number_of_records_returned.Close
                
                'Records_to_Update = Records_to_Update + 1
                
                For Y = LBound(Data_Array, 2) To UBound(Data_Array, 2)
                    'loop a row from the input variable array and assign values
                    'Last array value is designated for price data and is conditionally retrieved outside of this for loop
                    If Not (IsError(Data_Array(x, Y)) Or IsEmpty(Data_Array(x, Y))) Then
                        
                        If IsNumeric(Data_Array(x, Y)) Then
                            Row_Data(Y) = Data_Array(x, Y)
                        ElseIf Data_Array(x, Y) = "." Or Trim(Data_Array(x, Y)) = vbNullString Then
                            Row_Data(Y) = Null
                        Else
                            Row_Data(Y) = Data_Array(x, Y)
                        End If

                    Else
                        Row_Data(Y) = Null
                    End If

                Next Y
                
                If Not debugOnly Then
                    .AddNew field_names, Row_Data
                    .Update
                End If
                
            'Else
            '    number_of_records_returned.Close
            'End If
next_row:
        Next x
        
        'If Records_to_Update > 0 And Not debugOnly Then .UpdateBatch
        
    End With

    If Not (debugOnly Or Legacy_Combined_Data) Then 'retrieve price data from the legacy combined table
        'Legacy COmbined Data should be the first data retrieved
        Call database_details(True, legacy_abbreviation, table_name:=legacy_combined_table_name, db_path:=legacy_database_path)
    
        'T alias is for table that is being updated
        SQL = "Update " & table_name & " as T INNER JOIN [" & legacy_database_path & "]." & legacy_combined_table_name & " as Source_TBL ON Source_TBL.[Report_Date_as_YYYY-MM-DD]=T.[Report_Date_as_YYYY-MM-DD] AND Source_TBL.[CFTC_Contract_Market_Code]=T.[CFTC_Contract_Market_Code]" & _
            " SET T.[Price] = Source_TBL.[Price] WHERE T.[Report_Date_as_YYYY-MM-DD]>=CDate('" & Format(oldest_added_date, "YYYY-MM-DD") & "');"
        
        cn.Execute CommandText:=SQL, Options:=adCmdText + adExecuteNoRecords

    End If
    
    If Not debugOnly Then
        If Range(Report_Type & "_Combined").value = combined_wb_bool Then
            'This will signal to worksheet activate events to update the currently visible data
            COT_ABR_Match(Report_Type).Cells(1, 8).value = True
        End If
    End If
    
Close_Connection:
    
    If Err.Number <> 0 Then
        
        MsgBox "An error occurred while attempting to update table [ " & table_name & " ] in database " & cn.Properties("Data Source") & _
        vbNewLine & vbNewLine & _
        "Error description: " & Err.description

    End If

'    Set number_of_records_command = Nothing
'
'    If Not number_of_records_returned Is Nothing Then   'RecordSet object
'        If number_of_records_returned.State = adStateOpen Then number_of_records_returned.Close
'        Set number_of_records_returned = Nothing
'    End If
    
    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Public Function TransposeData(ByRef Data As Variant) As Variant
'===================================================================================================================
'Since recordset.getrows returns each array row as a database column, data will need to be parsed into rows for display
'===================================================================================================================
    Dim x As Long, Y As Byte, Output() As Variant

    ReDim Output(1 To UBound(Data, 2) + 1, 1 To UBound(Data, 1) + 1)
    
    For Y = LBound(Data, 1) To UBound(Data, 1)
        
        For x = LBound(Data, 2) To UBound(Data, 2)
            Output(x + 1, Y + 1) = IIf(IsNull(Data(Y, x)) And Not Y = UBound(Data, 1), 0, Data(Y, x))
        Next x
        
    Next Y
    
    TransposeData = Output

End Function

Sub delete_cftc_data_from_database(smallest_date As Date, Report_Type As String, Combined_Version As Boolean)


    Dim SQL As String, table_name As String, cn As Object, combined_wb_bool As Boolean
    
    Set cn = CreateObject("ADODB.Connection")

    database_details Combined_Version, Report_Type, cn, table_name
    
    On Error GoTo No_Table
    SQL = "DELETE FROM " & table_name & " WHERE [Report_Date_as_YYYY-MM-DD] >= Cdate('" & Format(smallest_date, "YYYY-MM-DD") & "');"

    With cn
        .Open
        .Execute SQL, , adExecuteNoRecords
        .Close
    End With

    Set cn = Nothing

Exit Sub
    
No_Table:
    
    MsgBox "TableL " & table_name & " not found within database."
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    
End Sub

Public Function Latest_Date(Report_Type As String, combined_wb_bool As Boolean, ICE_Query As Boolean) As Date
'===================================================================================================================
'Returns the date for the most recent data within a database
'===================================================================================================================
    Dim table_name As String, SQL As String, cn As Object, record As Object, var_str As String
    
    Const filter As String = "('Cocoa','B','RC','G','Wheat','W');"
    
    On Error GoTo Connection_Unavailable

    Set cn = CreateObject("ADODB.Connection")
    
    database_details combined_wb_bool, Report_Type, cn, table_name

    If DataBase_Not_Found Then
        Set cn = Nothing
        Latest_Date = 0
        Exit Function
    End If
    
    If Not ICE_Query Then var_str = "NOT "
    
    SQL = "SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & _
    " WHERE " & var_str & "[CFTC_Contract_Market_Code] IN " & filter

    With cn
        '.CursorLocation = adUseServer
        .Open
        Set record = .Execute(SQL, , adCmdText)
    End With
    
    If Not IsNull(record(0)) Then
        Latest_Date = record(0)
    Else
        Latest_Date = 0
    End If

Connection_Unavailable:

    If Err.Number <> 0 Then Latest_Date = 0

    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Function
Sub UpdateDatabasePrices(Data As Variant, Report_Type As String, combined_wb_bool As Boolean, price_column As Byte)
'===================================================================================================================
'Updates database with price data from a given array. Array should come from a worksheet
'===================================================================================================================
    Dim SQL As String, table_name As String, x As Integer, cn As Object, price_update_command As Object, CC_Column As Byte
    
    Const date_column As Byte = 1
    
    CC_Column = price_column - 1

    Set cn = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, Report_Type, cn, table_name

    SQL = "UPDATE " & table_name & _
        " SET [Price] = ? " & _
        " WHERE [CFTC_Contract_Market_Code] = ? AND [Report_Date_as_YYYY-MM-DD] = ?;"
    
    cn.Open
    
    Set price_update_command = CreateObject("ADODB.Command")

    With price_update_command
    
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = SQL
        .Prepared = True
        
        With .Parameters
            .Append price_update_command.CreateParameter("Price", adDouble, adParamInput, 20)
            .Append price_update_command.CreateParameter("Contract Code", adChar, adParamInput, 6)
            .Append price_update_command.CreateParameter("Date", adDBDate, adParamInput, 8)
        End With
        
    End With

    For x = LBound(Data, 1) To UBound(Data, 1)

        On Error GoTo Exit_Code
        
        With price_update_command

            With .Parameters
            
                If Not IsEmpty(Data(x, price_column)) Then
                    .Item("Price").value = Data(x, price_column)
                Else
                    .Item("Price").value = Null
                End If
                
                .Item("Contract Code").value = Data(x, CC_Column)
                .Item("Date").value = Data(x, date_column)
                
            End With
            
            .Execute
            
        End With
        
    Next x
    
Exit_Code:

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    Set price_update_command = Nothing

End Sub
Public Sub Retrieve_Price_From_Source_Upload_To_DB()
Attribute Retrieve_Price_From_Source_Upload_To_DB.VB_Description = "Takes the contract code from a currently active data sheet and retrieves price data and uploads it to each database where needed."
Attribute Retrieve_Price_From_Source_Upload_To_DB.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
'Retrieves dates from a given data table, retrieves accompanying dates and then uploads to database
'===================================================================================================================
    Dim Worksheet_Data() As Variant, WS As Variant, price_column As Byte, _
    Report_Type As String, Price_Symbols As Collection, contract_code As String, _
    Source_Ws As Worksheet, D As Byte, current_Filters() As Variant, LO As ListObject, Price_Data_Found As Boolean
    
    Const legacy_initial As String = "L"
    
    For Each WS In Array(LC, DC, TC)
        
        If ThisWorkbook.ActiveSheet Is WS Then
            Report_Type = Array("L", "D", "T")(D)
            Set Source_Ws = WS
            Exit For
        End If
        D = D + 1
    Next WS
    
    If Source_Ws Is Nothing Then Exit Sub
    
    Set LO = Source_Ws.ListObjects(Report_Type & "_Data")
    
    price_column = Evaluate("=VLOOKUP(""" & Report_Type & """,Report_Abbreviation,5,FALSE)") + 1
    
    With LO.DataBodyRange
        Worksheet_Data = .Resize(.Rows.Count, price_column).value
    End With
    
    contract_code = Worksheet_Data(1, price_column - 1)
    
    Set Price_Symbols = ContractDetails
    
    If HasKey(Price_Symbols, contract_code) Then
    
        Retrieve_Tuesdays_CLose Worksheet_Data, price_column, Price_Symbols(contract_code), overwrite_all_prices:=True, dates_in_column_1:=True, Data_Found:=Price_Data_Found
        
        If Price_Data_Found Then
            
            Price_Data_Found = False
            
            'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
            UpdateDatabasePrices Worksheet_Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column
            
            'Overwrites all other database tables with price data from Legacy_Combined
            
            overwrite_with_legacy_combined_prices contract_code
            
            ChangeFilters LO, current_Filters
                
            LO.DataBodyRange.Columns(price_column).value = WorksheetFunction.Index(Worksheet_Data, 0, price_column)
            
            RestoreFilters LO, current_Filters
        Else
            MsgBox "Unable to retrieve data."
        End If
        
    Else
        MsgBox "A symbol is unavailable for: [ " & contract_code & " ] on worksheet " & Symbols.name & "."
    End If
    
End Sub

Sub overwrite_with_legacy_combined_prices(Optional specific_contract As String = ";", Optional minimum_date As Variant)
'===========================================================================================================
' Overwrites a given table found within a database with price data from the legacy combined table in the legacy database
'===========================================================================================================
Dim SQL As String, table_name As String, cn As Object, legacy_database_path As String
  
Dim Report_Type As Variant, Combined_Version As Variant, contract_filter As String
    
    Const legacy_initial As String = "L"
    
    On Error GoTo Close_Connections

    database_details True, legacy_initial, db_path:=legacy_database_path
    
    contract_filter = " WHERE F.[Price] <> NULL"
    
    If Not specific_contract = ";" Then
        contract_filter = contract_filter & " AND F.[CFTC_Contract_Market_Code] = '" & specific_contract & "'"
    End If
    
    If Not IsMissing(minimum_date) Then
        If IsDate(minimum_date) Then
            contract_filter = contract_filter & " AND T.[Report_Date_as_YYYY-MM-DD] >= Cdate('" & Format(minimum_date, "YYYY-MM-DD") & "')"
        End If
    End If
    
    contract_filter = contract_filter & ";"
    
    For Each Report_Type In Array(legacy_initial, "D", "T")
        
        For Each Combined_Version In Array(True, False)
            
            If Combined_Version = True Then
                'Related Report tables currently share the same database so only 1 connecton is needed between the 2
                Set cn = CreateObject("ADODB.Connection")
                Call database_details(CBool(Combined_Version), CStr(Report_Type), cn)
                cn.Open
                
            End If
            
            If Not (Report_Type = legacy_initial And Combined_Version = True) Then
                
                database_details CBool(Combined_Version), CStr(Report_Type), table_name:=table_name
            
                SQL = "UPDATE " & table_name & _
                    " as T INNER JOIN [" & legacy_database_path & "].Legacy_Combined as F ON (F.[Report_Date_as_YYYY-MM-DD] = T.[Report_Date_as_YYYY-MM-DD] AND T.[CFTC_Contract_Market_Code] = F.[CFTC_Contract_Market_Code])" & _
                    " SET T.[Price] = F.[Price]" & contract_filter
                
                cn.Execute SQL, , adExecuteNoRecords

            End If
            
        Next Combined_Version
        
        cn.Close
        Set cn = Nothing
        
    Next Report_Type

Close_Connections:
    
    If Not cn Is Nothing Then
        With cn
            If .State = adStateOpen Then .Close
        End With
        Set cn = Nothing
    End If

End Sub


Sub Replace_All_Prices()
Attribute Replace_All_Prices.VB_Description = "Retrieves price data for all available contracts where a price symbol is available and uploads it to each database."
Attribute Replace_All_Prices.VB_ProcData.VB_Invoke_Func = " \n14"
'=======================================================================================================
'For every contract code for which a price symbol is available, query new prices and upload to every database
'=======================================================================================================
    Dim Symbol_Info As Collection, CO As Variant, SQL As String, cn As Object, New_Data_Available As Boolean, _
    table_name As String, record As Object, Data() As Variant

    Const legacy_initial As String = "L"
    Const combined_Bool As Boolean = True
    Const price_column As Byte = 3
    
    If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then
        Exit Sub
    End If
    
    On Error GoTo Close_Connection
    
    Set Symbol_Info = ContractDetails

    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    database_details combined_Bool, legacy_initial, cn, table_name
    
    cn.Open
    
    For Each CO In Symbol_Info
        
        If CO.PriceSymbol <> vbNullString Then
    
            SQL = "SELECT [Report_Date_as_YYYY-MM-DD],[CFTC_Contract_Market_Code],[Price] FROM " & table_name & " WHERE [CFTC_Contract_Market_Code] = '" & CO.contractCode & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
            
            With record
            
                .Open SQL, cn
                
                If Not .EOF And Not .BOF Then
                
                    Data = TransposeData(.GetRows)
                    .Close
                    
                    Call Retrieve_Tuesdays_CLose(Data, price_column, Symbol_Info(CO.contractCode), overwrite_all_prices:=True, dates_in_column_1:=True, Data_Found:=New_Data_Available)
                    
                    If New_Data_Available Then
                        
                        New_Data_Available = False
                        
                        Call UpdateDatabasePrices(Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column)
                        
                        overwrite_with_legacy_combined_prices CO.contractCode
                    
                    End If
                    
                Else
                    .Close
                End If
                
            End With

            'Overwrites all other database tables with price data from Legacy_Combine
        End If
        
    Next CO

Close_Connection:

    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub
Public Function change_table_data(LO As ListObject, retrieve_combined_data As Boolean, Report_Type As String, contract_code As String, triggered_by_linked_charts As Boolean) As Variant
'===================================================================================================================
'Retrieves data and updates a given listobject
'===================================================================================================================
    Dim Data() As Variant, Script_2_Run As String, Last_Calculated_Column As Integer, _
    First_Calculated_Column As Byte, table_filters() As Variant, Data_Variables As Range
    
    Dim DebugTasks As New TimerC, queryForCode As String
    
    Const calculateFieldTask As String = "Calculations", outputToSheetTask As String = "Output to worksheet."
    
    Const resetUsedRangeTask As String = "Reset used range." ',applyFiltersTask As String = "Re-apply worksheet filters."
    
    queryForCode = "Query database for (" & contract_code & ")"
    
    DebugTasks.description = "Retrieve data from database and place on worksheet."
    
    Set Data_Variables = COT_ABR_Match(Report_Type)
    
    Data = Data_Variables.value
    
    First_Calculated_Column = 3 + Data(1, 5) 'Raw data coluumn count + (price) + (Empty) + (start)
    Last_Calculated_Column = Data(1, 3)
    
    With DebugTasks
    
        .StartTask queryForCode
        
         Data = QueryDatabaseForContract(Report_Type, retrieve_combined_data, contract_code)
         
        .EndTask queryForCode
    
        ReDim Preserve Data(1 To UBound(Data, 1), 1 To Last_Calculated_Column)
            
        .StartTask calculateFieldTask
        
        Select Case Report_Type
            Case "L":
                Data = Legacy_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26)
            Case "D":
                Data = Disaggregated_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26)
            Case "T":
                Data = TFF_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26, 52)
            
        End Select
        
        .EndTask calculateFieldTask
        
    End With
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    ChangeFilters LO, table_filters
    
    DebugTasks.StartTask outputToSheetTask
    
    With LO
    
        With .DataBodyRange
            On Error Resume Next
            .SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0
            .Cells(1, 1).Resize(UBound(Data, 1), UBound(Data, 2)).Value2 = Data
        End With
        
        'DebugTasks.StartTask resetUsedRangeTask
        
        .Resize .Range.Resize(UBound(Data, 1) + 1, .Range.Columns.Count)
        
        Reset_Worksheet_UsedRange .Range
        
        'DebugTasks.EndTask resetUsedRangeTask
        
    End With
    
    DebugTasks.EndTask outputToSheetTask
    
    With LO.Sort
        If .SortFields.Count > 0 Then .Apply
    End With
    
    If Not triggered_by_linked_charts Then
        
        With DebugTasks
        
            '.StartTask applyFiltersTask
        
             RestoreFilters LO, table_filters
             Application.ScreenUpdating = True
            
            '.EndTask applyFiltersTask
            
        End With
            
    End If
    
    Data_Variables.Cells(1, 8).Resize(1, 2).value = Array(False, contract_code)
    
    Debug.Print DebugTasks.ToString
    
    Application.Calculation = xlCalculationAutomatic
    
End Function
Public Sub Manage_Table_Visual(Report_Type As String, Calling_Worksheet As Worksheet)
'==================================================================================================
'This sub is used to update the GUI after contracts have been updated upon activation of the calling worksheet
'==================================================================================================
    Dim Current_Details() As Variant
    
    Current_Details = COT_ABR_Match(Report_Type).value
    
    If Current_Details(1, 8) = True Then
        Call change_table_data(Calling_Worksheet.ListObjects(Report_Type & "_Data"), CBool(Current_Details(1, 7)), Report_Type, CStr(Current_Details(1, 9)), False)
    End If
    
End Sub
Sub Latest_Contracts()
Attribute Latest_Contracts.VB_Description = "Queries available databases for the latest contracts in a specified timeframe."
Attribute Latest_Contracts.VB_ProcData.VB_Invoke_Func = " \n14"

Dim L_Table As String, L_Path As String, D_Path As String, D_Table As String, queryAvailable As Boolean
     
    Dim SQL_2 As String, date_cutoff As String, connectionString As String, qt As QueryTable

    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]"
    
    Const queryName As String = "Update Latest Contracts"
        
    On Error GoTo 0
    
    date_cutoff = "CDATE('" & Format(DateSerial(Year(Now) - 2, 1, 1), "yyyy-mm-dd") & "')"
    
    ' Get all contract [names,codes,dates] From legacy and Disaggregated. Inner join it with a max date query and return names,codes,availability where max date
    
    On Error GoTo Close_Connection

        database_details True, "L", , L_Table, L_Path
        database_details True, "D", , D_Table, D_Path

'FQ.code in ('Wheat','B','RC','W','G','Cocoa')
'IIF(

'ICE Brent Crude Futures and Options - ICE Futures Europe

SQL_2 = "Select contractNames.contractCode,contractNames.cName,iif(ISNULL(FQ.code)=-1,'L,T', iif(LCASE(Trim(FQ.name)) LIKE  'ice%ice%','D','L,D')) From" & _
    " (((" & _
         " SELECT {nameField} as cName,{dateField} as currentDate,{codeField} as contractCode" & _
          " FROM [{L_Path}].{L_Table}" & _
           " WHERE {dateField} >= {date_cutoff}" & _
            " Union" & _
            "(SELECT D.{nameField} as cName,D.{dateField} as currentDate,D.{codeField} as contractCode" & _
               " FROM {D_Path}.{D_Table} as D" & _
           " LEFT JOIN {L_Path}.{Legacy_Combined} as L" & _
           " ON L.{codeField}= D.{codeField} and D.{dateField}=L.{dateField}" & _
           " WHERE L.{codeField} Is Null" & _
           " AND D.{dateField} >= {CDATE('2020-01-01')})" & _
        " ) as contractNames" & " INNER Join" & _
         " (SELECT MAX({dateField}) as maxdate,{codeField} as contractCode FROM [{L_Path}].{L_Table}" & _
          " GROUP BY {codeField}" & " HAVING MAX({dateField})>={date_cutoff}" & " Union" & _
             " SELECT MAX({dateField}) as maxdate,{codeField} as contractCode FROM [{D_Path}].{D_Table}" & _
              " GROUP BY {codeField} HAVING MAX({dateField})>={date_cutoff})as latestContracts" & _
        " ON latestContracts.contractCode=contractNames.contractCode AND latestContracts.maxdate = contractNames.currentDate" & _
    " )" & _
    " LEFT JOIN (Select {codeField} as code, {dateField} as contractDate,{nameField} as name From [{D_Path}].{D_Table}) as FQ" & _
    " ON FQ.code  = latestContracts.contractCode and FQ.contractDate=latestContracts.maxdate)" & _
    " Order by contractNames.cName ASC;"
    
    Call Interpolator(SQL_2, nameField, dateField, codeField, L_Path, L_Table, dateField, date_cutoff, _
 nameField, dateField, codeField, D_Path, D_Table, L_Path, L_Table, codeField, codeField, dateField, dateField, codeField, dateField, date_cutoff, _
             dateField, codeField, L_Path, L_Table, codeField, dateField, _
            date_cutoff, dateField, codeField, D_Path, D_Table, codeField, _
            dateField, date_cutoff, codeField, dateField, nameField, D_Path, D_Table)
            
    
    connectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & L_Path & ";"
    
    On Error GoTo Close_Connection
     
    With Available_Contracts
     
        For Each qt In .QueryTables
            If qt.name Like queryName & "*" Then
                queryAvailable = True
                Exit For
            End If
        Next qt
        
        If Not queryAvailable Then
            Set qt = .QueryTables.Add(connectionString, .Range("G1"))
        End If
        
    End With
    
    With qt
    
        .CommandText = SQL_2
        .BackgroundQuery = True
        .Connection = connectionString
        .CommandType = xlCmdSql
        .MaintainConnection = False
        .name = queryName
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SaveData = False
        .fieldNames = False
        '.PreserveFormatting
        
        Set AfterEventHolder = New ClassQTE
        
        AfterEventHolder.HookUpLatestContracts qt
        
        .Refresh True
        
    End With
    
Close_Connection:
    'Debug.Print Err.Number
    
End Sub
Sub Latest_Contracts_After_Refresh(RefreshedQueryTable As QueryTable, Success As Boolean)
        
    Dim results() As Variant
    
    Set AfterEventHolder = Nothing
    
    If Success Then
        
        With RefreshedQueryTable.ResultRange
            results = .value
            .ClearContents
        End With
        
        With Available_Contracts.ListObjects("Contract_Availability")
    
            With .DataBodyRange
                .SpecialCells(xlCellTypeConstants).ClearContents
                .Cells(1, 1).Resize(UBound(results, 1), UBound(results, 2)).value = results
            End With
            
            .Resize .Range.Cells(1, 1).Resize(UBound(results, 1) + 1, .ListColumns.Count)
    
        End With

    End If
    
    '=IF(COUNTIF(Symbols_TBL[Contract Code -CFTC],[@[Contract Code]]),true,False)
    
End Sub
Sub Interpolator(inputStr As String, ParamArray values() As Variant)

    Dim RightBrace As Long, leftSplit() As String, Z As Long, D As Long, noEscapeCharacter As Boolean

    leftSplit = Split(inputStr, "{")

    Const escapeCharacter As String = "\"

    For Z = LBound(leftSplit) To UBound(leftSplit)

        If Z > LBound(leftSplit) Then
        
            If Right$(leftSplit(Z), 1) = "\" Then
                noEscapeCharacter = False
            Else
                noEscapeCharacter = True
            End If
            
            If noEscapeCharacter Then
                
                RightBrace = InStr(1, leftSplit(Z), "}")
                
                leftSplit(Z) = values(D) & Right$(leftSplit(Z), Len(leftSplit(Z)) - RightBrace)
                D = D + 1
            End If

        End If

    Next Z

    inputStr = Join(leftSplit, vbNullString)

End Sub

Function GetAllContractDataInFavorites(Report_Type As String, getCombinedData As Boolean, minWeeks As Long) As Collection

    Dim SQL As String, tableName As String, cn As Object, record As Object, SQL2 As String, _
    favoritedContractCodes As String, queryResult() As Variant, fieldNames As String, contractClctn As Collection, AllContracts As New Collection
    
    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]"
          
    queryResult = WorksheetFunction.Transpose(Variable_Sheet.ListObjects("Current_Favorites").DataBodyRange.Columns(1).value)
    
    favoritedContractCodes = Join(QuotedForm(queryResult, "'"), ",")
          
    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    Call database_details(getCombinedData, Report_Type, cn, tableName)   'Generates a connection string and assigns a table to modify

    With cn
        .Open
        Set record = .Execute(CommandText:=tableName, Options:=adCmdTable) 'This record will be used to retrieve field names
    End With
    
    fieldNames = FilterColumnsAndDelimit(FieldsFromRecordSet(record, use_brackets:=True), Report_Type, includePriceColumn:=False)   'Field names from database returned as an array
    record.Close
'
    SQL2 = "SELECT " & codeField & " FROM " & tableName & " WHERE " & dateField & " = CDATE('" & Format(Variable_Sheet.Range("Most_Recently_Queried_Date").value, "yyyy-mm-dd") & "') AND " & codeField & " in (" & favoritedContractCodes & ");"
    
    'datefilter = "AND " & dateField & " >= CDATE('" & Format(DateAdd("ww", -minWeeks, Variable_Sheet.Range("Most_Recently_Queried_Date").value), "yyyy-mm-dd") & _
    "')
    
    SQL = "SELECT " & fieldNames & " FROM " & tableName & _
    " WHERE " & codeField & " in (" & SQL2 & ") Order BY " & codeField & " ASC," & dateField & " ASC;"
    
    Erase queryResult
    
    With record
        .Open SQL, cn, adOpenStatic, adLockReadOnly, adCmdText
        queryResult = TransposeData(.GetRows)
        .Close
    End With
    
    cn.Close
    
    Dim codeColumn As Byte, nameColumn As Byte, rowIndex As Long, columnIndex As Byte, _
    queryRow() As Variant, CC As Variant, Output As New Collection
    
    codeColumn = UBound(queryResult, 2)
    nameColumn = 2
    
    ReDim queryRow(1 To codeColumn)
    
    With AllContracts
        'Group contracts into separate collections for further processing
        For rowIndex = LBound(queryResult, 1) To UBound(queryResult, 1)
        
            For columnIndex = 1 To codeColumn
                queryRow(columnIndex) = queryResult(rowIndex, columnIndex)
            Next columnIndex
        
            On Error GoTo Create_Contract_Collection
            Set contractClctn = .Item(queryRow(codeColumn))
            
            On Error GoTo 0
            'Use dates as a key
            contractClctn.Add queryRow, CStr(queryRow(1))
            
        Next rowIndex
        
        Erase queryResult

    End With
    
    With Output
        For rowIndex = 1 To AllContracts.Count
            .Add Multi_Week_Addition(AllContracts(rowIndex), Append_Type.Multiple_1d), AllContracts(rowIndex)(1)(codeColumn)
        Next rowIndex
    End With
    
    Set GetAllContractDataInFavorites = Output
    
    Exit Function
    
Create_Contract_Collection:

    Set contractClctn = New Collection
    AllContracts.Add contractClctn, queryRow(codeColumn)
    
    Resume Next
    
End Function

Private Sub Generate_Database_Dashboard()

    Dim contractClctn As Collection, tempData As Variant, Output() As Variant, _
    outputRow As Integer, tempRow As Integer, tempCol As Byte, commercialNetColumn As Byte, _
    stochasticCalculations() As Variant, dateRange As Integer, Z As Byte, targetColumn As Integer, _
    nonCommercialNetColumn As Byte, queryFutOnly As Boolean
    
    Const threeYearsInWeeks As Integer = 156, sixMonthsInWeeks As Byte = 26, oneYearInWeeks As Byte = 52, _
    previousWeeksToCalculate As Byte = 1
    
    On Error GoTo No_Data
    
    If DashV2.Shapes("FUT Only").OLEFormat.Object.value = 1 Then
        queryFutOnly = True
    End If
    
    Set contractClctn = GetAllContractDataInFavorites("L", Not queryFutOnly, threeYearsInWeeks + previousWeeksToCalculate + 2)
    
    With contractClctn
        If .Count = 0 Then Exit Sub
        ReDim Output(1 To .Count, 1 To DashV2.ListObjects("Dashboard_Results").ListColumns.Count)
    End With
    
    On Error GoTo 0
    
    For Each tempData In contractClctn
        
        outputRow = outputRow + 1
        'Contract name without exchange name
        Output(outputRow, 1) = Left$(tempData(UBound(tempData, 1), 2), InStrRev(tempData(UBound(tempData, 1), 2), "-") - 2)
        
        commercialNetColumn = UBound(tempData, 2) + 1
        nonCommercialNetColumn = commercialNetColumn + 2
        
        ReDim Preserve tempData(1 To UBound(tempData, 1), 1 To UBound(tempData, 2) + 4)
        
        'Commercial Net Position calculation
        For tempRow = LBound(tempData, 1) To UBound(tempData, 1)
        
            tempData(tempRow, commercialNetColumn) = tempData(tempRow, 7) - tempData(tempRow, 8)
            tempData(tempRow, nonCommercialNetColumn) = tempData(tempRow, 4) - tempData(tempRow, 5)
            'Net Change
'            If tempRow > LBound(tempData, 1) Then
'                tempData(tempRow, commercialNetColumn + 1) = tempData(tempRow, commercialNetColumn) - tempData(tempRow - 1, commercialNetColumn)
'            End If
            
        Next tempRow
        
        'Commercial Long,Short, Net,Total Oi   Index against all dates available
        For Z = 0 To 6
            targetColumn = Array(3, 7, 8, commercialNetColumn, 4, 5, nonCommercialNetColumn)(Z)
            Output(outputRow, 2 + Z) = Stochastic_Calculations(targetColumn, UBound(tempData, 1), tempData, previousWeeksToCalculate, True)(1)
        Next Z
        
        'Variable Index calculations
        For Z = 0 To 2
            dateRange = Array(threeYearsInWeeks, oneYearInWeeks, sixMonthsInWeeks)(Z)
            
            'Add a loop to test if all data within the range are within the date range
            '
            '
            '================================================================================
            If UBound(tempData, 1) >= dateRange Then
                Output(outputRow, 9 + Z) = Stochastic_Calculations(CInt(commercialNetColumn), dateRange, tempData, previousWeeksToCalculate, True)(1)
            End If
        Next Z
        
        contractClctn.Remove tempData(1, commercialNetColumn - 1)
        
    Next tempData
    
    On Error GoTo 0
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    With DashV2
        
        With .ListObjects("Dashboard_Results")
            
            With .DataBodyRange
            
                .ClearContents
                
                With .Resize(UBound(Output, 1), UBound(Output, 2))
                    .value = Output
                    .Sort key1:=.Columns(1), Orientation:=xlSortColumns, ORder1:=xlAscending, header:=xlNo, MatchCase:=False
                End With
                
            End With
            
            If UBound(Output, 1) <> .ListRows.Count Then
                .Resize .Range.Resize(UBound(Output, 1) + 1, .ListColumns.Count)
            End If
            
        End With
        
        .Range("A1").value = Variable_Sheet.Range("Most_Recently_Queried_Date").value
        
    End With
    
    Re_Enable
    
    Exit Sub
    
No_Data:
    MsgBox "An error occurred. " & Err.description
End Sub


Public Function Assign_Charts_WS(Report_Type As String) As Worksheet
    
    Dim WSA() As Variant, T As Byte
    
    WSA = Array(L_Charts, D_Charts, T_Charts)
    
    T = Application.Match(Report_Type, Array("L", "D", "T"), 0) - 1
    
    Set Assign_Charts_WS = WSA(T)

End Function

Public Function Assign_Linked_Data_Sheet(Report_Type As String) As Worksheet

    Dim WSA() As Variant, T As Byte
    
    WSA = Array(LC, DC, TC)
    
    T = Application.Match(Report_Type, Array("L", "D", "T"), 0) - 1
    
    Set Assign_Linked_Data_Sheet = WSA(T)
    
End Function
    
Public Sub Save_For_Github()

    If UUID Then
        Range("Github_Version").value = True
        Custom_SaveAS
    End If

End Sub
Public Sub Handle_Contract_GUI()
    
    Dim WS As Worksheet, Userform_Active As Boolean
    
    Set WS = ThisWorkbook.ActiveSheet
    
    Userform_Active = IsLoadedUserform("Contract_Selection")
    
    Select Case True
        
        Case WS Is LC, WS Is TC, WS Is DC
        
        'Case ff
        
        Case Else
        
            If Userform_Active Then
                Unload Contract_Selection
            End If
    
    End Select
    
End Sub
Private Sub Load_Database_path_Selector_Userform()
     Database_Path_Selector.Show
End Sub
Private Sub Adjust_Contract_Selection_Shapes()

    Dim gg As Range, WS As Variant
    
    For Each WS In Array(LC, DC, TC)
        Set gg = WS.Range("A1")
    
        With WS.Shapes("Launch Selection")
    
            .Top = gg.Top
            .Left = gg.Left
            .Width = gg.Width
            .Height = gg.Height
    
        End With
    
    Next WS

End Sub
Sub OverwritePricesAfterDate()

'======================================================================================================
'Will generate an array to represent all data within the legacy combined database since a certain date N.
'Price data will be retrieved for that array and used to update the database.
'======================================================================================================
Dim Symbols As Collection, SQL As String, cn As Object, tableName As String, queryResult() As Variant, CC As Integer

    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]"
    
    Dim codeColumn As Byte, rowIndex As Long, columnIndex As Byte, contractClctn As Collection, _
    queryRow() As Variant, AllContracts As New Collection, minDate As String, succesfulPriceRetrieval As Boolean
    
    minDate = InputBox("Input date in form YYYY-MM-DD")

    Set cn = CreateObject("ADODB.COnnection")
    
    database_details True, "L", cn, tableName
    
    SQL = "SELECT " & Join(Array(dateField, codeField, "Price"), ",") & " FROM " & tableName & " WHERE " & dateField & " >=Cdate('" & minDate & "');"
    
    codeColumn = 2

    With cn
        .Open
         queryResult = TransposeData(.Execute(SQL, , adCmdText).GetRows)
        .Close
    End With
    
    Set Symbols = ContractDetails
    
    ReDim queryRow(1 To UBound(queryResult, 2))
    
    With AllContracts
        'Group contracts into separate collections for further processing
        For rowIndex = LBound(queryResult, 1) To UBound(queryResult, 1)
        
            For columnIndex = 1 To UBound(queryResult, 2)
                queryRow(columnIndex) = queryResult(rowIndex, columnIndex)
            Next columnIndex
        
            On Error GoTo Create_Contract_Collection
            Set contractClctn = .Item(queryRow(codeColumn))
            
            On Error GoTo 0
            'Use dates as a key
            contractClctn.Add queryRow, CStr(queryRow(1))
            
        Next rowIndex
        
        Erase queryResult
        Erase queryRow
        
    End With
    
    With AllContracts
    
        For CC = .Count To 1 Step -1
            
            Set contractClctn = .Item(CC)
            
            queryResult = Multi_Week_Addition(contractClctn, Append_Type.Multiple_1d)
            
            .Remove queryResult(1, codeColumn)
            
            If HasKey(Symbols, CStr(queryResult(1, codeColumn))) Then
            
                Retrieve_Tuesdays_CLose queryResult, 3, Symbols(queryResult(1, codeColumn)), True, True, succesfulPriceRetrieval
                
                If succesfulPriceRetrieval Then .Add queryResult, queryResult(1, codeColumn)
            
            End If
        
        Next CC
    
    End With
    
    queryResult = Multi_Week_Addition(AllContracts, Append_Type.Multiple_2d)
    
    On Error GoTo 0
    
    UpdateDatabasePrices queryResult, "L", True, 3
    
    overwrite_with_legacy_combined_prices minimum_date:=CDate(minDate)
    
    Exit Sub
    
Create_Contract_Collection:

    Set contractClctn = New Collection
    AllContracts.Add contractClctn, queryRow(codeColumn)
    
    Resume Next

End Sub


