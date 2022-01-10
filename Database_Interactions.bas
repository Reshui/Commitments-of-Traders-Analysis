Attribute VB_Name = "Database_Interactions"
Public DataBase_Not_Found As Boolean
Option Explicit

Private Sub database_details(combined_wb_bool As Boolean, report_type As String, Optional ByRef cn As Object, Optional ByRef table_name As String, Optional ByRef db_path As String)
'===================================================================================================================
'Determines database path is available, and generates a connection string.
'===================================================================================================================
    Dim Report_Name As String, user_database_path As String ', T As Variant
    
    Report_Name = Evaluate("VLOOKUP(""" & report_type & """,Report_Abbreviation,2,FALSE)")
    
    If UUID Then
        db_path = Environ$("USERPROFILE") & "\Documents\" & Report_Name & ".accdb"
        DataBase_Not_Found = False
    Else

        user_database_path = Assign_Linked_Data_Sheet(report_type).Range("Database_Path").value
            
        If user_database_path = vbNullString Or Dir(user_database_path) = "" Then
        
            DataBase_Not_Found = True
            
            If Data_Retrieval.Running_Weekly_Retrieval Then
                Exit Sub
            Else
                MsgBox IIf(Report_Name = "TFF", "Traders in Financial Futures", Report_Name) & " database not found."
                Re_Enable
                End
            End If
            
        Else
            db_path = user_database_path
            DataBase_Not_Found = False
        End If

    End If
    
    If Not IsMissing(table_name) Then table_name = Evaluate("=VLOOKUP(""" & report_type & """,Report_Abbreviation,2,FALSE)") & IIf(combined_wb_bool = True, "_Combined", "_Futures_Only")
    
    If Not cn Is Nothing Then cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & db_path & ";"
    
    'Set T = cn.Properties
    
End Sub
Public Sub store_contract_ids(report_type As String)
'============================================================================================
'When called this sub will update a calculated table so that it contains the names and contract codes
'for contracts with the greatest date within the database.
'============================================================================================
    Dim LO As ListObject
    
    Set LO = Range(report_type & "_Contract_TBL").ListObject
    
    Latest_Contracts report_type, combined_wb_bool:=True, LO:=LO
    
End Sub
Private Function generate_filtered_columns_string(columns_in_table As Variant, report_type As String) As String
'===================================================================================================================
'Loops table found on Variabkes Worksheet that contains True/False values for wanted columns
'Using that data return a joined list of wanted colunmns with reordering
'===================================================================================================================
    Dim X As Long, Y As Long, order_clctn As New Collection, column_names_str As String, report_type_full_name As String
    Dim wanted_columns() As Variant, columns_available() As Variant, final_column_names() As String, DD As Variant
    
    report_type_full_name = Evaluate("VLOOKUP(""" & report_type & """,Report_Abbreviation,2,FALSE)")
    
    columns_available = Variable_Sheet.ListObjects(report_type_full_name & "_User_Selected_Columns").DataBodyRange.Value2

    wanted_columns = Filter_Market_Columns(True, False, convert_skip_col_to_general:=False, report_type:=report_type, Create_Filter:=True)

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
'record is a record containing a single row of data from which field names are retrieved,formatted and output as an array
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

Function retrieve_data_from_db(report_type As String, combined_wb_bool As Boolean, contract_code As String) As Variant
'===================================================================================================================
'Retrieves filtered data from database and returns as an array
'===================================================================================================================
    Dim record_set As Object, cn As Object, table_name As String

    Dim SQL As String, db_data As Variant, column_names_sql As String, field_names() As Variant

    On Error GoTo Close_Connection

    Set cn = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, report_type, cn, table_name

    With cn
        .Open
        Set record_set = .Execute("SELECT TOP 1 * FROM " & table_name & ";")
    End With
    
    field_names = generate_field_names(record_set, use_brackets:=True)
    record_set.Close
    column_names_sql = generate_filtered_columns_string(field_names, report_type:=report_type)

    SQL = "SELECT " & column_names_sql & " FROM " & table_name & " WHERE [CFTC_Contract_Market_Code]='" & contract_code & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
    
    With record_set
        .Open SQL, cn
        db_data = .GetRows 'Returns a 0 based 2D array
    End With
     
    retrieve_data_from_db = db_columns_to_array(db_data)

Close_Connection:

    record_set.Close
    cn.Close
    Set cn = Nothing
    Set record_set = Nothing

End Function

Public Sub Update_DataBase(data_array As Variant, combined_wb_bool As Boolean, report_type As String, Price_Symbols As Collection)
'===================================================================================================================
'Uodates a given data table one row at a time
'===================================================================================================================
    Dim table_name As String, cn As Object, legacy_combined_price_query As Object, number_of_records_command As Object, _
    combined_price_record As Object, field_names() As Variant, X As Long, record As Object, number_of_records_returned As Object, _
    row_data() As Variant, Y As Long, legacy_combined_table_name As String, combined_cn As Object, legacy_combined_data As Boolean, Records_to_Update As Long
    
    Dim yymmdd As Long, contract_code As String
    
    On Error GoTo Close_Connection
    
    Const date_column As Long = 2
    Const contract_code_column As Long = 4
    Const legacy_abbreviation As String = "L"

    ReDim row_data(LBound(data_array, 2) To UBound(data_array, 2))
    
    Set cn = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    Set number_of_records_command = CreateObject("ADODB.Command")
    Set number_of_records_returned = CreateObject("ADODB.RecordSet")

    If report_type = legacy_abbreviation And combined_wb_bool = True Then legacy_combined_data = True

    database_details combined_wb_bool, report_type, cn, table_name

    With cn
        .Open
        Set record = .Execute("SELECT TOP 1 * FROM " & table_name & ";") 'This record will be used to retrieve field names
        .CursorLocation = adUseClient                                    'Batch update won't work otherwise
    End With
    
    field_names = generate_field_names(record, use_brackets:=False)     'field names from database returned as an array
    
    record.Close
    
    If Not legacy_combined_data Then
        'A new connection is used otherwise only 1 update can be done at a time using the cn object because of the .cursorlocation property
        'legacy combined data may alo be in a different database
        Set combined_cn = CreateObject("ADODB.Connection")
        
        database_details True, legacy_abbreviation, table_name:=legacy_combined_table_name, cn:=combined_cn
        
        combined_cn.Open
        
        Set legacy_combined_price_query = CreateObject("ADODB.Command")
        
        With legacy_combined_price_query
            'Command will be used to get price data from the Legacy combined database
            .ActiveConnection = combined_cn
            .CommandText = "SELECT Price FROM " & legacy_combined_table_name & " WHERE [As_of_Date_In_Form_YYMMDD] = ? AND [CFTC_Contract_Market_Code] = ?;"
            .CommandType = adCmdText
            .Prepared = True
            
            With .Parameters
                .Append legacy_combined_price_query.CreateParameter("YYMMDD", adDouble, adParamInput, 6)
                .Append legacy_combined_price_query.CreateParameter("Contract_Code", adVarWChar, adParamInput, 6)
            End With
            
        End With
        
    End If
    
    With number_of_records_command
        'Command will be used to ensure that there aren't duplicate entries in the database
        .ActiveConnection = cn
        .CommandText = "SELECT Count([As_of_Date_In_Form_YYMMDD]) FROM " & table_name & " WHERE [As_of_Date_In_Form_YYMMDD] = ? AND [CFTC_Contract_Market_Code] = ?;"
        .CommandType = adCmdText
        .Prepared = True
        
        With .Parameters
            .Append number_of_records_command.CreateParameter("YYMMDD", adDouble, adParamInput, 6)
            .Append number_of_records_command.CreateParameter("Contract_Code", adVarWChar, adParamInput, 6)
        End With

    End With
    
    With record
        'This Recordset will be used to add new data to the database table via the batchupdate method
        .CursorLocation = adUseClient
        .Open table_name, cn, adOpenForwardOnly, adLockBatchOptimistic
        
        For X = LBound(data_array, 1) To UBound(data_array, 1)
            
            'If data_array(X, 3) > DateSerial(2000, 1, 1) Then GoTo next_row
            
            contract_code = data_array(X, contract_code_column)
            yymmdd = data_array(X, date_column)
            
            number_of_records_command.Parameters("Contract_Code").value = contract_code
            number_of_records_command.Parameters("YYMMDD").value = yymmdd

            Set number_of_records_returned = number_of_records_command.Execute

            If number_of_records_returned(0) = 0 Then
                'Recordset contains a single field that is the result of how many rows match a given combo of date and contract code
                Records_to_Update = Records_to_Update + 1
                
                For Y = LBound(data_array, 2) To UBound(data_array, 2) - 1
                    'loop a row from the input variable array and assign values
                    'Last array value is designated for price data and is conditionally retrieved outside of this for loop
                    If Not (IsError(data_array(X, Y)) Or IsEmpty(data_array(X, Y))) Then
                        
                        If IsNumeric(data_array(X, Y)) Then
                            row_data(Y) = data_array(X, Y)
                        ElseIf data_array(X, Y) = "." Or Len(Replace(data_array(X, Y), " ", vbNullString)) = 0 Then
                            row_data(Y) = Null
                        Else
                            row_data(Y) = data_array(X, Y)
                        End If

                    Else
                        row_data(Y) = Null
                    End If

                Next Y
                
                If Not legacy_combined_data Then 'Get price data from legacy combined database
                    
                    If HasKey(Price_Symbols, contract_code) Then
                    
                        legacy_combined_price_query.Parameters("Contract_Code").value = contract_code
                        legacy_combined_price_query.Parameters("YYMMDD").value = yymmdd
                    
                        Set combined_price_record = legacy_combined_price_query.Execute
                    
                        If Not combined_price_record.EOF And Not combined_price_record.BOF Then
                            'This just tests if it is not an empty record
                            row_data(UBound(data_array, 2)) = combined_price_record(0)
                        Else
                            row_data(UBound(data_array, 2)) = Null
                        End If
                        
                    Else
                        row_data(UBound(data_array, 2)) = Null
                    End If
                    
                Else
                    Y = UBound(data_array, 2)
                    row_data(Y) = IIf(IsEmpty(data_array(X, Y)), Null, data_array(X, Y))
                End If
                
                .AddNew field_names, row_data

            End If
             
            If Records_to_Update = 500 Then
                Records_to_Update = 0
                .UpdateBatch
            End If
next_row:
        Next X
        
        If Records_to_Update > 0 Then .UpdateBatch
        
    End With

Close_Connection:
    
    If Err.Number <> 0 Then
    
        MsgBox "An error occurred while attempting to update table [ " & table_name & " ] in database " & cn.Properties("Data Source") & _
        vbNewLine & vbNewLine & _
        "Error description: " & Err.Description

    End If
    
    If Not number_of_records_returned Is Nothing Then
        'RecordSet object
        number_of_records_returned.Close
        Set number_of_records_returned = Nothing
    End If
    
    Set number_of_records_command = Nothing
    
    If Not record Is Nothing Then
        record.Close
        Set record = Nothing
    End If
    
    If Not legacy_combined_data Then
    
        If Not combined_price_record Is Nothing Then
            combined_price_record.Close
            Set combined_price_record = Nothing
        End If
        
        combined_cn.Close
        Set legacy_combined_price_query = Nothing
        Set combined_cn = Nothing
    
    End If
    
    cn.Close
    Set cn = Nothing
    
End Sub

Public Sub Latest_Contracts(report_type As String, combined_wb_bool As Boolean, LO As ListObject)
'===================================================================================================================
'Returns a 2D array (contract_name,contract code) for the latest data in the database
'===================================================================================================================
    Dim table_name As String, SQL As String, cn As Object, record As Object, _
    query_return() As Variant, returned_data As Variant

    Set cn = CreateObject("ADODB.Connection")

On Error GoTo Close_Connection
    
    database_details combined_wb_bool, report_type, cn, table_name

    SQL = "SELECT [Market_and_Exchange_Names],[CFTC_Contract_Market_Code] FROM " & table_name & _
          " as T_Name WHERE [Report_Date_as_YYYY-MM-DD] in (SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & ")" & _
          " ORDER BY [Market_and_Exchange_Names] ASC;"
    
    With cn
        .Open
        Set record = .Execute(SQL)
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

    Set record = Nothing
    cn.Close
    Set cn = Nothing
    
    If Err.Number <> 0 Then
        MsgBox "An error occurred while attempting to query the " & table_name & " table for the latest contract names." & vbNewLine & vbNewLine & _
        "Error Description :" & Err.Description
    End If
    
End Sub

Public Function db_columns_to_array(ByRef data As Variant) As Variant
'===================================================================================================================
'Since recordset.getrows returns each array row as a database column, data will need to be parsed into rows for display
'===================================================================================================================
    Dim X As Long, Y As Long, output() As Variant

    ReDim output(1 To UBound(data, 2) + 1, 1 To UBound(data, 1) + 1)

    For Y = 1 To UBound(output, 2)
    
        For X = 1 To UBound(output, 1)
            
            output(X, Y) = IIf(IsNull(data(Y - 1, X - 1)) And Not Y = UBound(output, 2), 0, data(Y - 1, X - 1))
        
        Next X
        
    Next Y
    
    db_columns_to_array = output
    
End Function

Public Function change_table_data(LO As ListObject, combined_workbook As Boolean, report_type As String, contract_code As String, triggered_by_linked_charts As Boolean) As Variant
'===================================================================================================================
'Retrieves data and updates a given listobject
'===================================================================================================================
    Dim data() As Variant, Script_2_Run As String, Last_Calculated_Column As Long, _
    First_Calculated_Column As Long, table_filters() As Variant
    
    First_Calculated_Column = 3 + Evaluate("=VLOOKUP(""" & report_type & """,Report_Abbreviation,5,FALSE)")
     
    Last_Calculated_Column = Evaluate("=VLOOKUP(""" & report_type & """,Report_Abbreviation,3,FALSE)")
        
    data = retrieve_data_from_db(report_type, combined_workbook, contract_code)
    
    ReDim Preserve data(1 To UBound(data, 1), 1 To Last_Calculated_Column)
    
    Select Case report_type
        Case "L":
            data = Legacy_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
        Case "D":
            data = Disaggregated_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
        Case "T":
            data = TFF_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26, 52)
        
    End Select
    
    Application.ScreenUpdating = False
    
    ChangeFilters LO, table_filters
    
    With LO
    
        With .DataBodyRange
            On Error Resume Next
            .SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0
            .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
        End With
        
        .Resize .Range.Resize(UBound(data, 1) + 1, .Range.Columns.Count)
    
        Reset_Worksheet_UsedRange .Range
        
    End With
    
    If Not triggered_by_linked_charts Then
        RestoreFilters LO, table_filters
        Application.ScreenUpdating = True
    End If
    
End Function

Sub delete_cftc_data_from_database(yymmdd As Long, report_type As String)

    Dim SQL As String, table_name As String, cn As Object, combined_wb_bool As Boolean, X As Integer
    
    For X = 1 To 2
        
        If X = 2 Then combined_wb_bool = True
        
        Set cn = CreateObject("ADODB.Connection")
    
        database_details combined_wb_bool, report_type, cn, table_name
        
        On Error GoTo No_Table
        SQL = "DELETE FROM " & table_name & " WHERE [As_of_Date_In_Form_YYMMDD] >= " & yymmdd & ";"
    
        With cn
            .Open
            .Execute SQL
        End With
    
        cn.Close
        Set cn = Nothing
    
    Next X
    
Exit Sub
    
No_Table:
    MsgBox "TableL " & table_name & " not found within database."
    cn.Close
    Set cn = Nothing

End Sub

Public Function Latest_Date(report_type As String, combined_wb_bool As Boolean, ICE_Query As Boolean) As Date
'===================================================================================================================
'Returns the date for the most recent data within a table
'===================================================================================================================
    Dim table_name As String, SQL As String, cn As Object, record As Object, var_str As String
    
    Const filter As String = "('Cocoa','B','RC','G','Wheat','W');"
    
    Set cn = CreateObject("ADODB.Connection")
    
    database_details combined_wb_bool, report_type, cn, table_name

    If DataBase_Not_Found Then
        Set cn = Nothing
        Latest_Date = 0
        Exit Function
    End If
    
    var_str = IIf(ICE_Query = False, "NOT ", vbNullString)
    
    SQL = "SELECT TOP 1 MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & _
    " WHERE " & var_str & "[CFTC_Contract_Market_Code] IN " & filter
    
    On Error GoTo Table_Not_Exist
    
    With cn
        .Open
        Set record = .Execute(SQL)
    End With
    
    If Not IsNull(record(0)) Then
        Latest_Date = record(0)
    Else
        Latest_Date = 0
    End If
    
    record.Close
    cn.Close
    Set record = Nothing
    Set cn = Nothing

Exit Function

Table_Not_Exist:
    Latest_Date = 0
    cn.Close
    Set cn = Nothing
    
End Function
Sub update_database_prices(data As Variant, report_type As String, combined_wb_bool As Boolean, price_column As Long)
'===================================================================================================================
'Updates database with price data from a given array. Array should come from a worksheet
'===================================================================================================================
    Dim SQL As String, table_name As String, X As Long, cn As Object, cmd As Object, CC_Column As Long
    
    Const date_column As Integer = 1
    
    CC_Column = price_column - 1

    Set cn = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, report_type, cn, table_name

    SQL = "UPDATE " & table_name & _
        " SET [Price] = ? " & _
        " WHERE [CFTC_Contract_Market_Code] = ? AND [Report_Date_as_YYYY-MM-DD] = ?;"
    
    cn.Open
    
    Set cmd = CreateObject("ADODB.Command")

    With cmd
    
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = SQL
        .Prepared = True

        With .Parameters
            .Append cmd.CreateParameter("Price", adDouble, adParamInput, 20)
            .Append cmd.CreateParameter("Contract Code", adChar, adParamInput, 6)
            .Append cmd.CreateParameter("Date", adDBDate, adParamInput, 8)
        End With
        
    End With

    For X = LBound(data, 1) To UBound(data, 1)

        If Not IsEmpty(data(X, price_column)) Then

            On Error GoTo Exit_Code
            
            With cmd

                With .Parameters
                    .Item("Price").value = data(X, price_column)
                    .Item("Contract Code").value = data(X, CC_Column)
                    .Item("Date").value = data(X, date_column)
                End With
                
                .Execute
                
            End With

        End If
        
    Next X
    
Exit_Code:
    cn.Close
    Set cn = Nothing
    Set cmd = Nothing

End Sub
Public Sub Retrieve_Price_From_Source_Upload_To_DB()
Attribute Retrieve_Price_From_Source_Upload_To_DB.VB_Description = "Takes the contract code from a currently active data sheet and retrieves price data and uploads it to each database where needed."
Attribute Retrieve_Price_From_Source_Upload_To_DB.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
'Retrieves dates from a given data table, retrieves accompanying dates and then uploads to database
'===================================================================================================================
    Dim Worksheet_Data() As Variant, WS As Variant, price_column As Long, _
    report_type As String, Price_Symbols As Collection, contract_code As String, _
    Source_Ws As Worksheet, D As Long, current_filters() As Variant, LO As ListObject
    
    Const legacy_initial As String = "L"
    
    For Each WS In Array(LC, DC, TC)
        
        If ThisWorkbook.ActiveSheet Is WS Then
            report_type = Array("L", "D", "T")(D)
            Set Source_Ws = WS
            Exit For
        End If
        D = D + 1
    Next WS
    
    If Source_Ws Is Nothing Then Exit Sub
    
    Set LO = Source_Ws.ListObjects(report_type & "_Data")
    
    price_column = Evaluate("=VLOOKUP(""" & report_type & """,Report_Abbreviation,5,FALSE)") + 1
    
    With LO.DataBodyRange
        Worksheet_Data = .Resize(.Rows.Count, price_column).value
    End With
    
    contract_code = Worksheet_Data(1, price_column - 1)
    
    Set Price_Symbols = Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")
    
    If HasKey(Price_Symbols, contract_code) Then
    
        Retrieve_Tuesdays_CLose Worksheet_Data, price_column, Price_Symbols(contract_code), True
        
        'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
        update_database_prices Worksheet_Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column
        
        'Overwrites all other database tables with price data from Legacy_Combined
        
        overwrite_with_legacy_combined_prices contract_code
        
        ChangeFilters LO, current_filters
            
        LO.DataBodyRange.Columns(price_column).value = WorksheetFunction.Index(Worksheet_Data, 0, price_column)
        
        RestoreFilters LO, current_filters
        
    Else
        MsgBox "A symbol is unavailable for: [ " & contract_code & " ] on worksheet " & Symbols.Name & "."
    End If
    
End Sub

Sub overwrite_with_legacy_combined_prices(Optional specific_contract As String = ";")
'===========================================================================================================
' Overwrites a given table found within a database with price data from the legacy combined table in the legacy database
'===========================================================================================================
Dim SQL As String, table_name As String, cn As Object, Legacy_DataBase_Path As String
  
Dim report_type As Variant, Combined_Version As Variant, contract_filter As String
    
    Const legacy_initial As String = "L"
    
    database_details True, legacy_initial, db_path:=Legacy_DataBase_Path
    
    If Not specific_contract = ";" Then
        contract_filter = " WHERE F.[CFTC_Contract_Market_Code]='" & specific_contract & "';"
    End If
    
    For Each report_type In Array(legacy_initial, "D", "T")
        
        For Each Combined_Version In Array(True, False)
            
            If Combined_Version = True Then
            
                Set cn = CreateObject("ADODB.Connection")
                
                database_details CBool(Combined_Version), CStr(report_type), cn
                
                cn.Open
                
            End If
            
            If Not (report_type = legacy_initial And Combined_Version = True) Then
                
                database_details CBool(Combined_Version), CStr(report_type), table_name:=table_name
            
                SQL = "UPDATE " & table_name & _
                    " as T INNER JOIN [" & Legacy_DataBase_Path & "].Legacy_Combined as F ON (F.[Report_Date_as_YYYY-MM-DD] = T.[Report_Date_as_YYYY-MM-DD] AND T.[CFTC_Contract_Market_Code] = F.[CFTC_Contract_Market_Code])" & _
                    " SET T.[Price] = F.[Price]" & contract_filter
                
                cn.Execute SQL

            End If
            
        Next Combined_Version
        
        cn.Close
        Set cn = Nothing
        
    Next report_type

End Sub
Sub contract_change(Data_Ws As Worksheet, report_type As String, Linked_Charts_WS As Worksheet, change_contract_name As Boolean, change_combined_box As Boolean, triggered_by_charts As Boolean)
'===========================================================================================================
' This Subroutine interfaces with COmbobox change events on data/chart worksheets to update table data for a given
'Combination of Report Type and Version
'===========================================================================================================
    Dim combined_wb As Boolean, contract_code As Variant, target_table_name As String, CB As Object, Chart_ListBox As Object, Data_LB As Object
     
    On Error GoTo Exit_Procedure
    
    Set CB = Data_Ws.OLEObjects("Select_Contract_Name").Object
    
    contract_code = WorksheetFunction.VLookup(CB.value, Range(report_type & "_Contract_TBL"), 2, False)
     
    combined_wb = IIf(Data_Ws.OLEObjects("COMB_List").Object.value = "Combined", True, False)
    
    target_table_name = report_type & "_Data"
    
    change_table_data Data_Ws.ListObjects(target_table_name), combined_wb, report_type, CStr(contract_code), triggered_by_charts

    If Data_Ws Is ThisWorkbook.ActiveSheet Then Data_Ws.Range("A1").Select
    
    If change_contract_name Then 'Update Combobox values on chart sheet
    
        With ThisWorkbook.Worksheets(Linked_Charts_WS.Name) 'Can't access worksheet defined variables directly

            .Disable_Data_Change = True
                .OLEObjects("Sheet_Selection").Object.value = CB.value
                .Range("A4").Value2 = CB.value 'save value from combobox to worksheet
            .Disable_Data_Change = False
            
        End With
    
    End If
    
    If change_combined_box Then 'Update Combined ListBox on Chart Sheet
    
        With ThisWorkbook.Worksheets(Linked_Charts_WS.Name)
        
            Set Chart_ListBox = .OLEObjects("Report_Version").Object
            
            Set Data_LB = Data_Ws.OLEObjects("COMB_List").Object
            
            If Chart_ListBox.value <> Data_LB.value Then
            
                .Disable_Data_Change = True
                    Chart_ListBox.value = Data_LB.value
                .Disable_Data_Change = False
                
            End If
            
        End With
    
    End If
        
Exit_Procedure:

End Sub

Sub Replace_All_Prices()
Attribute Replace_All_Prices.VB_Description = "Retrieves price data for all available contracts where a price symbol is available and uploads it to each database."
Attribute Replace_All_Prices.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim Symbol_Info As Collection, Z As Long, SQL As String, cn As Object, _
    table_name As String, record As Object, Symbols_P() As Variant, data() As Variant, contract_code As String

    Const legacy_initial As String = "L"
    Const combined_Bool As Boolean = True
    Const price_column = 3
    
    Symbols_P = Symbols.ListObjects("Symbols_TBL").DataBodyRange.Columns(1).value
    
    Set Symbol_Info = Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")

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
                
                    data = db_columns_to_array(.GetRows)
                    .Close
                    
                    Call Retrieve_Tuesdays_CLose(data, price_column, Symbol_Info(contract_code), dates_in_column_1:=True)
                    
                    Call update_database_prices(data, legacy_initial, combined_wb_bool:=True, price_column:=price_column)
                
                Else
                    .Close
                End If
                
            End With

            'Overwrites all other database tables with price data from Legacy_Combine
        End If
        
    Next Z
    
    cn.Close
    Set cn = Nothing
    Set record = Nothing
    
    overwrite_with_legacy_combined_prices
    
End Sub
