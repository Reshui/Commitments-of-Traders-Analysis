Attribute VB_Name = "Database_Interactions"
Public DataBase_Not_Found As Boolean
Option Explicit

Private Sub database_details(combined_wb_bool As Boolean, Report_Type As String, Optional ByRef cn As Object, Optional ByRef table_name As String, Optional ByRef db_path As String)
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

        user_database_path = HUB.Range(Report_Type & "_Database_Path").value
            
        If user_database_path = vbNullString Or Dir(user_database_path) = "" Then
        
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
    
    If Not cn Is Nothing Then cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & db_path & ";"
    
    'Set T = cn.Properties
    
End Sub
Public Sub save_recent_contract_identifiers(Report_Type As String)
'============================================================================================
'When called this sub will update a calculated table so that it contains the names and contract codes
'for contracts with the greatest date within the database.
'============================================================================================
    Dim LO As ListObject
    
    Set LO = Range(Report_Type & "_Contract_TBL").ListObject
    
    Call Latest_Contracts(Report_Type, combined_wb_bool:=True, LO:=LO)
    
End Sub
Private Function generate_filtered_columns_string(columns_in_table As Variant, Report_Type As String) As String
'===================================================================================================================
'Loops table found on Variables Worksheet that contains True/False values for wanted columns
'An array of wanted columns with some re-ordering is returned
'===================================================================================================================
    Dim X As Long, Y As Long, order_clctn As New Collection, column_names_str As String, report_type_full_name As String
    Dim wanted_columns() As Variant, columns_available() As Variant, final_column_names() As String, DD As Variant
    
    If Report_Type = "T" Then
        report_type_full_name = "TFF"
    Else
        report_type_full_name = Evaluate("VLOOKUP(""" & Report_Type & """,Report_Abbreviation,2,FALSE)")
    End If
    
    columns_available = Variable_Sheet.ListObjects(report_type_full_name & "_User_Selected_Columns").DataBodyRange.Value2

    wanted_columns = Filter_Market_Columns(True, False, convert_skip_col_to_general:=False, Report_Type:=Report_Type, Create_Filter:=True)
    
    For X = LBound(wanted_columns) To UBound(wanted_columns)
        'Loop list o wanted contracts from stored table of TRUE False values
        'Add field name from database if value doesn't equal skipcolumn
        If Not wanted_columns(X) = xlSkipColumn Then
            order_clctn.Add columns_in_table(X), columns_available(X, 1)
        End If
        
    Next X

    Y = 0
    
    ReDim final_column_names(1 To order_clctn.Count + 1)

    final_column_names(UBound(final_column_names)) = "Price"

    For Each DD In Array(3, 1, 4)
    
        Y = Y + 1
        
        If Not Y = 3 Then
            final_column_names(Y) = order_clctn(columns_available(DD, 1))
        Else
            final_column_names(UBound(final_column_names) - 1) = order_clctn(columns_available(DD, 1))
            Y = 2
        End If
        
        order_clctn.Remove columns_available(DD, 1)
        
    Next DD

    For X = 1 To order_clctn.Count
        Y = Y + 1
        final_column_names(Y) = order_clctn(X)
    Next X
    
    generate_filtered_columns_string = Join(final_column_names, ",")

End Function

Function generate_field_names(record As Object, use_brackets As Boolean) As Variant
'===================================================================================================================
'record is a RecordSET object containing a single row of data from which field names are retrieved,formatted and output as an array
'===================================================================================================================
    Dim X As Long, fields_in_record() As Variant, Z As Long, field_name As String

    ReDim fields_in_record(1 To record.Fields.Count - 1)

    For X = 0 To record.Fields.Count - 1
    
        field_name = record(X).Name
        
        If Not field_name = "ID" Then
            Z = Z + 1
            If use_brackets Then
                fields_in_record(Z) = "[" & field_name & "]"
            Else
                fields_in_record(Z) = field_name
            End If
            
        End If

    Next X
'
    generate_field_names = fields_in_record

End Function

Function Retrieve_Contract_Data_From_DB(Report_Type As String, combined_wb_bool As Boolean, contract_code As String) As Variant
'===================================================================================================================
'Retrieves filtered data from database and returns as an array
'===================================================================================================================
    Dim record As Object, cn As Object, table_name As String

    Dim SQL As String, db_data As Variant, column_names_sql As String, field_names() As Variant

    On Error GoTo Close_Connection

    Set cn = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, Report_Type, cn, table_name

    With cn
        .Open
        Set record = .Execute(table_name, , adCmdTable)
    End With
    
    field_names = generate_field_names(record, use_brackets:=True)
    
    record.Close
    
    column_names_sql = generate_filtered_columns_string(field_names, Report_Type:=Report_Type)

    SQL = "SELECT " & column_names_sql & " FROM " & table_name & " WHERE [CFTC_Contract_Market_Code]='" & contract_code & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
    
    With record
        .Open SQL, cn
        db_data = .GetRows 'Returns a 0 based 2D array
    End With
     
    Retrieve_Contract_Data_From_DB = db_columns_to_array(db_data)

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

Public Sub Update_DataBase(data_array As Variant, combined_wb_bool As Boolean, Report_Type As String)
'===================================================================================================================
'Uodates a given data table one row at a time
'===================================================================================================================
    Dim table_name As String, cn As Object, number_of_records_command As Object, _
    field_names() As Variant, X As Long, record As Object, number_of_records_returned As Object, _
    row_data() As Variant, Y As Long, legacy_combined_table_name As String, Legacy_Combined_Data As Boolean, Records_to_Update As Long, oldest_added_date As Date
    
    Dim row_date As Date, contract_code As String, SQL As String, legacy_database_path As String
    
    On Error GoTo Close_Connection
    
    Const yyyy_mm_dd_column As Long = 3
    Const contract_code_column As Long = 4
    Const legacy_abbreviation As String = "L"

    ReDim row_data(LBound(data_array, 2) To UBound(data_array, 2))
    
    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    Set number_of_records_command = CreateObject("ADODB.Command")
    Set number_of_records_returned = CreateObject("ADODB.RecordSet")

    If Report_Type = legacy_abbreviation And combined_wb_bool = True Then Legacy_Combined_Data = True

    Call database_details(combined_wb_bool, Report_Type, cn, table_name)   'Generates a connection string and assigns a table to modify

    With cn
        .CursorLocation = adUseClient                                    'Batch update won't work otherwise
        .Open
        Set record = .Execute(CommandText:=table_name, Options:=adCmdTable) 'This record will be used to retrieve field names
    End With
    
    field_names = generate_field_names(record, use_brackets:=False)     'Field names from database returned as an array
    record.Close

    With number_of_records_command
        'Command will be used to ensure that there aren't duplicate entries in the database
        .ActiveConnection = cn
        .CommandText = "SELECT Count([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & " WHERE [Report_Date_as_YYYY-MM-DD] = ? AND [CFTC_Contract_Market_Code] = ?;"
        .CommandType = adCmdText
        .Prepared = True
        
        With .Parameters
            .Append number_of_records_command.CreateParameter("YYYY-MM-DD", adDate, adParamInput)
            .Append number_of_records_command.CreateParameter("Contract_Code", adVarWChar, adParamInput, 6)
        End With

    End With
    
    With record
        'This Recordset will be used to add new data to the database table via the batchupdate method
        .Open table_name, cn, adOpenForwardOnly, adLockBatchOptimistic
        
        oldest_added_date = data_array(LBound(data_array, 1), yyyy_mm_dd_column)

        For X = LBound(data_array, 1) To UBound(data_array, 1)
            
            'If data_array(X, yyyy_mm_dd_column) > DateSerial(2000, 1, 1) Then GoTo next_row
            
            contract_code = data_array(X, contract_code_column)
            row_date = data_array(X, yyyy_mm_dd_column)
            
            If Legacy_Combined_Data Then
                If row_date < oldest_added_date Then oldest_added_date = row_date
            End If
            
            number_of_records_command.Parameters("Contract_Code").value = contract_code
            number_of_records_command.Parameters("YYYY-MM-DD").value = row_date

            Set number_of_records_returned = number_of_records_command.Execute

            If number_of_records_returned(0) = 0 Then
                'If new row can be uniquely identified with a date andcontract code
                number_of_records_returned.Close
                
                Records_to_Update = Records_to_Update + 1
                
                For Y = LBound(data_array, 2) To UBound(data_array, 2)
                    'loop a row from the input variable array and assign values
                    'Last array value is designated for price data and is conditionally retrieved outside of this for loop
                    If Not (IsError(data_array(X, Y)) Or IsEmpty(data_array(X, Y))) Then
                        
                        If IsNumeric(data_array(X, Y)) Then
                            row_data(Y) = data_array(X, Y)
                        ElseIf data_array(X, Y) = "." Or Trim(data_array(X, Y)) = vbNullString Then
                            row_data(Y) = Null
                        Else
                            row_data(Y) = data_array(X, Y)
                        End If

                    Else
                        row_data(Y) = Null
                    End If

                Next Y
                
                .AddNew field_names, row_data
                
            Else
                number_of_records_returned.Close
            End If
next_row:
        Next X
        
        If Records_to_Update > 0 Then .UpdateBatch
        
    End With

    If Not Legacy_Combined_Data And Records_to_Update > 0 Then 'retrieve price data from the legacy combined table
        'Legacy COmbined Data should be the first data retrieved
        Call database_details(True, legacy_abbreviation, table_name:=legacy_combined_table_name, db_path:=legacy_database_path)
    
        'T alias is for table that is being updated
        SQL = "Update " & table_name & " as T INNER JOIN [" & legacy_database_path & "]." & legacy_combined_table_name & " as Source_TBL ON Source_TBL.[Report_Date_as_YYYY-MM-DD]=T.[Report_Date_as_YYYY-MM-DD] AND Source_TBL.[CFTC_Contract_Market_Code]=T.[CFTC_Contract_Market_Code]" & _
            " SET T.[Price] = Source_TBL.[Price] WHERE T.[Report_Date_as_YYYY-MM-DD]>=CDate('" & Format(oldest_added_date, "YYYY-MM-DD") & "');"
        
        cn.Execute CommandText:=SQL, Options:=adCmdText + adExecuteNoRecords

    End If
    
    If Range(Report_Type & "_Combined").value = combined_wb_bool Then
        'This will signal to worksheet activate events to update the currently visible data
        COT_ABR_Match(Report_Type).Offset(, 7).value = True
    End If
    
Close_Connection:
    
    If Err.Number <> 0 Then
        
        MsgBox "An error occurred while attempting to update table [ " & table_name & " ] in database " & cn.Properties("Data Source") & _
        vbNewLine & vbNewLine & _
        "Error description: " & Err.Description

    End If

    Set number_of_records_command = Nothing
    
    If Not number_of_records_returned Is Nothing Then   'RecordSet object
        If number_of_records_returned.State = adStateOpen Then number_of_records_returned.Close
        Set number_of_records_returned = Nothing
    End If
    
    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Public Sub Latest_Contracts(Report_Type As String, combined_wb_bool As Boolean, LO As ListObject)
'===================================================================================================================
'Returns a 2D array (contract_name,contract code) for the latest data in the database
'===================================================================================================================
    Dim table_name As String, SQL As String, cn As Object, record As Object, _
    query_return() As Variant, returned_data As Variant

    On Error GoTo Close_Connection

    Set cn = CreateObject("ADODB.Connection")
    
    database_details combined_wb_bool, Report_Type, cn, table_name

    SQL = "SELECT [Market_and_Exchange_Names],[CFTC_Contract_Market_Code] FROM " & table_name & _
          " as T_Name WHERE [Report_Date_as_YYYY-MM-DD] in (SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & ")" & _
          " ORDER BY [Market_and_Exchange_Names] ASC;"
    
    With cn
        .Open
        Set record = .Execute(SQL, , adCmdText)
    End With

    query_return = record.GetRows 'Returns a 0 based 2D array
    
    record.Close
    returned_data = db_columns_to_array(query_return)
    
    With LO
    
        With .DataBodyRange
            .ClearContents
            .Cells(1, 1).Resize(UBound(returned_data, 1), UBound(returned_data, 2)).value = returned_data
        End With
        .Resize .Range.Cells(1, 1).Resize(UBound(returned_data, 1) + 1, 2)
        
    End With
    
Close_Connection:
    
    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    If Err.Number <> 0 Then
        MsgBox "An error occurred while attempting to query the " & table_name & " table for the latest contract names." & vbNewLine & vbNewLine & _
        "Error Description :" & Err.Description
    End If
    
End Sub

Public Function db_columns_to_array(ByRef Data As Variant) As Variant
'===================================================================================================================
'Since recordset.getrows returns each array row as a database column, data will need to be parsed into rows for display
'===================================================================================================================
    Dim X As Long, Y As Long, output() As Variant

    ReDim output(1 To UBound(Data, 2) + 1, 1 To UBound(Data, 1) + 1)

    For Y = 1 To UBound(output, 2)
    
        For X = 1 To UBound(output, 1)
            
            output(X, Y) = IIf(IsNull(Data(Y - 1, X - 1)) And Not Y = UBound(output, 2), 0, Data(Y - 1, X - 1))
        
        Next X
        
    Next Y
    
    db_columns_to_array = output
    
End Function

Sub delete_cftc_data_from_database(smallest_date As String, Report_Type As String)


    Dim SQL As String, table_name As String, cn As Object, combined_wb_bool As Boolean, X As Integer
    
    
    For X = 1 To 2
        
        If X = 2 Then combined_wb_bool = True
        
        Set cn = CreateObject("ADODB.Connection")
    
        database_details combined_wb_bool, Report_Type, cn, table_name
        
        On Error GoTo No_Table
        SQL = "DELETE FROM " & table_name & " WHERE [Report_Date_as_YYYY-MM-DD] >= Cdate('" & Format(smallest_date, "YYYY-MM-DD") & "');"
    
        With cn
            .Open
            .Execute SQL, , adExecuteNoRecords
            .Close
        End With

        Set cn = Nothing
    
    Next X
    
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
    Dim table_name As String, SQL As String, cn As adodb.Connection, record As Object, var_str As String
    
    Const filter As String = "('Cocoa','B','RC','G','Wheat','W');"
    
    On Error GoTo Connection_Unavailable

    Set cn = CreateObject("ADODB.Connection")
    
    database_details combined_wb_bool, Report_Type, cn, table_name

    If DataBase_Not_Found Then
        Set cn = Nothing
        Latest_Date = 0
        Exit Function
    End If
    
    var_str = IIf(ICE_Query = False, "NOT ", vbNullString)
    
    SQL = "SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & _
    " WHERE " & var_str & "[CFTC_Contract_Market_Code] IN " & filter
    
    With cn
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
Sub update_database_prices(Data As Variant, Report_Type As String, combined_wb_bool As Boolean, price_column As Long)
'===================================================================================================================
'Updates database with price data from a given array. Array should come from a worksheet
'===================================================================================================================
    Dim SQL As String, table_name As String, X As Long, cn As Object, price_update_command As Object, CC_Column As Long
    
    Const date_column As Integer = 1
    
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

    For X = LBound(Data, 1) To UBound(Data, 1)

        If Not IsEmpty(Data(X, price_column)) Then

            On Error GoTo Exit_Code
            
            With price_update_command

                With .Parameters
                    .Item("Price").value = Data(X, price_column)
                    .Item("Contract Code").value = Data(X, CC_Column)
                    .Item("Date").value = Data(X, date_column)
                End With
                
                .Execute
                
            End With

        End If
        
    Next X
    
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
    Dim Worksheet_Data() As Variant, WS As Variant, price_column As Long, _
    Report_Type As String, Price_Symbols As Collection, contract_code As String, _
    Source_Ws As Worksheet, D As Long, current_filters() As Variant, LO As ListObject, Price_Data_Found As Boolean
    
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
    
    Set Price_Symbols = Get_Price_Symbols
    
    If HasKey(Price_Symbols, contract_code) Then
    
        Retrieve_Tuesdays_CLose Worksheet_Data, price_column, Price_Symbols(contract_code), dates_in_column_1:=True, Data_Found:=Price_Data_Found
        
        If Price_Data_Found Then
            
            Price_Data_Found = False
            
            'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
            update_database_prices Worksheet_Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column
            
            'Overwrites all other database tables with price data from Legacy_Combined
            
            overwrite_with_legacy_combined_prices contract_code
            
            ChangeFilters LO, current_filters
                
            LO.DataBodyRange.Columns(price_column).value = WorksheetFunction.Index(Worksheet_Data, 0, price_column)
            
            RestoreFilters LO, current_filters
        
        End If
        
    Else
        MsgBox "A symbol is unavailable for: [ " & contract_code & " ] on worksheet " & Symbols.Name & "."
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
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If

End Sub


Sub Replace_All_Prices()
Attribute Replace_All_Prices.VB_Description = "Retrieves price data for all available contracts where a price symbol is available and uploads it to each database."
Attribute Replace_All_Prices.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Symbol_Info As Collection, Z As Long, SQL As String, cn As Object, New_Data_Available As Boolean, _
    table_name As String, record As Object, Symbols_P() As Variant, Data() As Variant, contract_code As String

    Const legacy_initial As String = "L"
    Const combined_Bool As Boolean = True
    Const price_column = 3
    
    If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then
        Exit Sub
    End If
    
    On Error GoTo Close_Connection
    
    Symbols_P = Symbols.ListObjects("Symbols_TBL").DataBodyRange.Columns(1).value
    
    Set Symbol_Info = Get_Price_Symbols

    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    database_details combined_Bool, legacy_initial, cn, table_name
    
    cn.Open
    
    For Z = 1 To UBound(Symbols_P, 1)
        
        contract_code = Symbols_P(Z, 1)
        
        If HasKey(Symbol_Info, contract_code) Then
        
            SQL = "SELECT [Report_Date_as_YYYY-MM-DD],[CFTC_Contract_Market_Code],[Price] FROM " & table_name & " WHERE [CFTC_Contract_Market_Code] = '" & contract_code & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
            
            With record
            
                .Open SQL, cn
                
                If Not .EOF And Not .BOF Then
                
                    Data = db_columns_to_array(.GetRows)
                    .Close
                    
                    Call Retrieve_Tuesdays_CLose(Data, price_column, Symbol_Info(contract_code), dates_in_column_1:=True, Data_Found:=New_Data_Available)
                    
                    If New_Data_Available Then
                        New_Data_Available = False
                        Call update_database_prices(Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column)
                    End If
                    
                Else
                    .Close
                End If
                
            End With

            'Overwrites all other database tables with price data from Legacy_Combine
        End If
        
    Next Z

Close_Connection:

    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    If Err.Number = 0 Then overwrite_with_legacy_combined_prices
    
End Sub
Public Function change_table_data(LO As ListObject, retrieve_combined_data As Boolean, Report_Type As String, contract_code As String, triggered_by_linked_charts As Boolean) As Variant
'===================================================================================================================
'Retrieves data and updates a given listobject
'===================================================================================================================
    Dim Data() As Variant, Script_2_Run As String, Last_Calculated_Column As Long, _
    First_Calculated_Column As Long, table_filters() As Variant, Data_Variables As Range
    
    Set Data_Variables = COT_ABR_Match(Report_Type)
    
    With Data_Variables
        Data = Variable_Sheet.Range(.Cells(1, 1), .Offset(, .CurrentRegion.Columns.Count - 1)).value
    End With
    
    First_Calculated_Column = 3 + Data(1, 5) 'Raw data coluumn count + (price) + (Empty) + (start)
    Last_Calculated_Column = Data(1, 3)

    Data = Retrieve_Contract_Data_From_DB(Report_Type, retrieve_combined_data, contract_code)
    
    ReDim Preserve Data(1 To UBound(Data, 1), 1 To Last_Calculated_Column)
    
    Select Case Report_Type
        Case "L":
            Data = Legacy_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26)
        Case "D":
            Data = Disaggregated_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26)
        Case "T":
            Data = TFF_Multi_Calculations(Data, UBound(Data, 1), First_Calculated_Column, 156, 26, 52)
        
    End Select
    
    Application.ScreenUpdating = False
    
    ChangeFilters LO, table_filters
    
    With LO
    
        With .DataBodyRange
            On Error Resume Next
            .SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0
            .Cells(1, 1).Resize(UBound(Data, 1), UBound(Data, 2)).Value2 = Data
        End With
        
        .Resize .Range.Resize(UBound(Data, 1) + 1, .Range.Columns.Count)
    
        Reset_Worksheet_UsedRange .Range
        
    End With
    
    With LO.Sort    'Will allow user to maintain their sort order
        If .SortFields.Count > 0 Then .Apply
    End With
    
    If Not triggered_by_linked_charts Then
        RestoreFilters LO, table_filters
        Application.ScreenUpdating = True
    End If
    
    Data_Variables.Offset(, 7).Resize(1, 2).value = Array(False, contract_code)
    
End Function
Public Sub Manage_Table_Visual(Report_Type As String, Calling_Worksheet As Worksheet)
    
    Dim Current_Details() As Variant
    
    With COT_ABR_Match(Report_Type)
        Current_Details = Variable_Sheet.Range(.Cells(1, 1), .Offset(, .CurrentRegion.Columns.Count - 1))
    End With
    
    If Current_Details(1, 8) = True Then 'thisworkbook.Worksheets(calling_worksheet.Name).update_table=True
        Call change_table_data(Calling_Worksheet.ListObjects(Report_Type & "_Data"), CBool(Current_Details(1, 7)), Report_Type, CStr(Current_Details(1, 9)), False)
    End If
    
End Sub
