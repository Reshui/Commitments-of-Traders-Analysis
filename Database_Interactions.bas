Attribute VB_Name = "Database_Interactions"
Private AfterEventHolder As ClassQTE

Option Explicit

 Sub database_details(combined_wb_bool As Boolean, reportType As String, Optional ByRef adodbConnection As Object, _
                    Optional ByRef table_name As String, Optional ByRef databasePath As String, Optional ByRef doesDatabaseExist As Boolean = False)
'===================================================================================================================
'Determines if database exists. IF it does the appropriate variables or properties are assigned values if needed.
'===================================================================================================================
    Dim Report_Name As String, userSpecifiedDatabasePath As String ', T As Variant
    
    If reportType = "T" Then
        Report_Name = "TFF"
    Else
        Report_Name = Evaluate("VLOOKUP(""" & reportType & """,Report_Abbreviation,2,FALSE)")
    End If
    
    If UUID Then
        databasePath = Environ$("USERPROFILE") & "\Documents\" & Report_Name & ".accdb"
        doesDatabaseExist = True
    Else

        userSpecifiedDatabasePath = Variable_Sheet.Range(reportType & "_Database_Path").value
            
        If userSpecifiedDatabasePath = vbNullString Or Not FileOrFolderExists(userSpecifiedDatabasePath) Then
            
            doesDatabaseExist = False
            
            If Data_Retrieval.Running_Weekly_Retrieval Then
                Exit Sub
            Else
                MsgBox Report_Name & " database not found."
                Re_Enable
                End
            End If
            
        ElseIf FileOrFolderExists(userSpecifiedDatabasePath) Then
            databasePath = userSpecifiedDatabasePath
            doesDatabaseExist = True
        Else
            doesDatabaseExist = False
            Exit Sub
        End If

    End If
    
    If Not IsMissing(table_name) Then table_name = Report_Name & IIf(combined_wb_bool = True, "_Combined", "_Futures_Only")
    
    If Not adodbConnection Is Nothing Then adodbConnection.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath & ";"
    
    'Set T = adodbConnection.Properties
    
End Sub
Private Function FilterColumnsAndDelimit(fieldsInDatabase As Variant, reportType As String, includePriceColumn As Boolean) As String
'===================================================================================================================
'Loops table found on Variables Worksheet that contains True/False values for wanted columns
'An array of wanted columns with some re-ordering is returned
'===================================================================================================================
    Dim wantedColumns() As Variant
    
    wantedColumns = Filter_Market_Columns(False, True, convert_skip_col_to_general:=False, reportType:=reportType, Create_Filter:=True, InputA:=fieldsInDatabase)
    
    Dim filterforDashboard, finalIndex As Byte
    
    If filterforDashboard Then
    
        Select Case reportType
            Case "L"
                finalIndex = 8
            Case "D"
                finalIndex = 13
            Case "T"
                finalIndex = 14
        End Select
        
        
    End If
    
    ReDim Preserve wantedColumns(LBound(wantedColumns) To UBound(wantedColumns) + IIf(includePriceColumn = True, 1, 0))
    
    If includePriceColumn Then wantedColumns(UBound(wantedColumns)) = "Price"

    FilterColumnsAndDelimit = WorksheetFunction.TextJoin(",", True, wantedColumns)
    
End Function

Function FieldsFromRecordSet(record As Object, encloseFieldsInBrackets As Boolean) As Variant
'===================================================================================================================
'record is a RecordSET object containing a single row of data from which field names are retrieved,formatted and output as an array
'===================================================================================================================
    Dim X As Integer, Z As Byte, fieldNamesInRecord() As Variant, currentFieldName As String, lcaseVersion As String

    ReDim fieldNamesInRecord(1 To record.Fields.count - 1)
    
    For X = 0 To record.Fields.count - 1
        
        currentFieldName = record(X).name
        
        If Not currentFieldName = "ID" Then
            Z = Z + 1
            If encloseFieldsInBrackets Then
                fieldNamesInRecord(Z) = "[" & currentFieldName & "]"
            Else
                fieldNamesInRecord(Z) = currentFieldName
            End If
            
        End If
            
    Next X
    
    FieldsFromRecordSet = fieldNamesInRecord

End Function

Function QueryDatabaseForContract(reportType As String, combined_wb_bool As Boolean, contract_code As String) As Variant
'===================================================================================================================
'Retrieves filtered data from database and returns as an array
'===================================================================================================================
    Dim record As Object, adodbConnection As Object, tableNameWithinDatabase As String

    Dim SQL As String, delimitedWantedColumns As String, allFieldNames() As Variant
    
'    Dim retrievalTimer As TimedTask
'    Set retrievalTimer = New TimedTask: retrievalTimer.Start "Contract Retrieval(" & contract_code & ") ~ " & Time

    On Error GoTo Close_Connection

    Set adodbConnection = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, reportType, adodbConnection, tableNameWithinDatabase

    With adodbConnection
        '.CursorLocation = adUseServer
        .Open
        Set record = .Execute(tableNameWithinDatabase, , adCmdTable)
    End With
    
    allFieldNames = FieldsFromRecordSet(record, encloseFieldsInBrackets:=True)
    
    record.Close
    
    delimitedWantedColumns = FilterColumnsAndDelimit(allFieldNames, reportType:=reportType, includePriceColumn:=True)

    SQL = "SELECT " & delimitedWantedColumns & " FROM " & tableNameWithinDatabase & " WHERE [CFTC_Contract_Market_Code]='" & contract_code & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
    
    With record
    
        .Open SQL, adodbConnection
         QueryDatabaseForContract = TransposeData(.GetRows)

    End With

    'If Not retrievalTimer Is Nothing Then retrievalTimer.DPrint
    
Close_Connection:

    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Function

Public Sub Update_DataBase(dataToUpload As Variant, uploadingCombinedData As Boolean, reportType As String, debugOnly As Boolean)
'===================================================================================================================
'Uodates a given data table one row at a time
'===================================================================================================================
    Dim tableToUpdateName As String, databaseFieldNames() As Variant, X As Long, _
    rowToUpload() As Variant, Y As Byte, legacyCombinedTableName As String, _
    Legacy_Combined_Data As Boolean, oldest_added_date As Date, uploadToDatabase As Boolean
    
    If debugOnly Then
        If MsgBox("Debug Active: Do you want to upload data to databse?", vbYesNo) = vbYes Then uploadToDatabase = True
    Else
        uploadToDatabase = True
    End If
    
    'dim  Records_to_Update As Long,
    Dim contract_code As String, SQL As String, legacyDatabasePath As String
    
    Dim record As Object, adodbConnection As Object ',row_date As Date, number_of_records_command As Object, number_of_records_returned As Object
    
    On Error GoTo Close_Connection
    
    Const yyyy_mm_dd_column As Byte = 3, legacy_abbreviation As String = "L"
    'Const contractCodeColumn As Byte = 4

    Set adodbConnection = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    'Set number_of_records_command = CreateObject("ADODB.Command")
    'Set number_of_records_returned = CreateObject("ADODB.RecordSet")

    If reportType = legacy_abbreviation And uploadingCombinedData = True Then Legacy_Combined_Data = True

    Call database_details(uploadingCombinedData, reportType, adodbConnection, tableToUpdateName)   'Generates a connection string and assigns a table to modify

    With adodbConnection
        '.CursorLocation = adUseServer                                   'Batch update won't work otherwise
        .Open
        Set record = .Execute(CommandText:=tableToUpdateName, Options:=adCmdTable) 'This record will be used to retrieve field names
    End With
    
    Dim fieldNameKeys As New Collection
    
    databaseFieldNames = FieldsFromRecordSet(record, encloseFieldsInBrackets:=False)  'Field names from database returned as an array
    
    ReDim rowToUpload(LBound(dataToUpload, 2) To UBound(dataToUpload, 2))
    
    record.Close

'    With number_of_records_command
'        'Command will be used to ensure that there aren't duplicate entries in the database
'        .ActiveConnection = adodbConnection
'        .CommandText = "SELECT Count([Report_Date_as_YYYY-MM-DD]) FROM " & tableToUpdateName & " WHERE [Report_Date_as_YYYY-MM-DD] = ? AND [CFTC_Contract_Market_Code] = ?;"
'        .CommandType = adCmdText
'        .Prepared = True
'
'        With .Parameters
'            .Append number_of_records_command.CreateParameter("YYYY-MM-DD", adDate, adParamInput)
'            .Append number_of_records_command.CreateParameter("Contract_Code", adVarWChar, adParamInput, 6)
'        End With
'
'    End With
    
    With record
        'This Recordset will be used to add new data to the database table via the batchupdate method
        .Open tableToUpdateName, adodbConnection, adOpenForwardOnly, adLockOptimistic
        
        oldest_added_date = dataToUpload(LBound(dataToUpload, 1), yyyy_mm_dd_column)

        For X = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
            
            If Not (debugOnly Or Legacy_Combined_Data) Then
                If dataToUpload(X, yyyy_mm_dd_column) < oldest_added_date Then oldest_added_date = dataToUpload(X, yyyy_mm_dd_column)
            End If
            
            'If dataToUpload(X, yyyy_mm_dd_column) > DateSerial(2000, 1, 1) Then GoTo next_row
            
'            contract_code = dataToUpload(X, contractCodeColumn)
'            row_date = dataToUpload(X, yyyy_mm_dd_column)
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
                
                For Y = LBound(dataToUpload, 2) To UBound(dataToUpload, 2)
                    'loop a row from the input variable array and assign values
                    'Last array value is designated for price data and is conditionally retrieved outside of this for loop
                    If Not (IsError(dataToUpload(X, Y)) Or IsEmpty(dataToUpload(X, Y))) Then
                        
                        If IsNumeric(dataToUpload(X, Y)) Then
                            rowToUpload(Y) = dataToUpload(X, Y)
                        ElseIf dataToUpload(X, Y) = "." Or Trim(dataToUpload(X, Y)) = vbNullString Then
                            rowToUpload(Y) = Null
                        Else
                            rowToUpload(Y) = dataToUpload(X, Y)
                        End If

                    Else
                        rowToUpload(Y) = Null
                    End If

                Next Y
                
                If uploadToDatabase Then
                    .AddNew databaseFieldNames, rowToUpload
                    .Update
                End If
                
            'Else
            '    number_of_records_returned.Close
            'End If
next_row:
        Next X
        
        'If Records_to_Update > 0 And Not debugOnly Then .UpdateBatch
        
    End With

    If uploadToDatabase And Not Legacy_Combined_Data Then 'retrieve price data from the legacy combined table
        'Legacy COmbined Data should be the first data retrieved
        Call database_details(True, legacy_abbreviation, table_name:=legacyCombinedTableName, databasePath:=legacyDatabasePath)
    
        'T alias is for table that is being updated
        SQL = "Update " & tableToUpdateName & " as T INNER JOIN [" & legacyDatabasePath & "]." & legacyCombinedTableName & " as Source_TBL ON Source_TBL.[Report_Date_as_YYYY-MM-DD]=T.[Report_Date_as_YYYY-MM-DD] AND Source_TBL.[CFTC_Contract_Market_Code]=T.[CFTC_Contract_Market_Code]" & _
            " SET T.[Price] = Source_TBL.[Price] WHERE T.[Report_Date_as_YYYY-MM-DD]>=CDate('" & Format(oldest_added_date, "YYYY-MM-DD") & "');"
        
        adodbConnection.Execute CommandText:=SQL, Options:=adCmdText + adExecuteNoRecords

    End If
    
    If Not debugOnly Then
        If Range(reportType & "_Combined").value = uploadingCombinedData Then
            'This will signal to worksheet activate events to update the currently visible data
            COT_ABR_Match(reportType).Cells(1, 8).value = True
        End If
    End If
    
Close_Connection:
    
    If Err.Number <> 0 Then
        
        MsgBox "An error occurred while attempting to update table [ " & tableToUpdateName & " ] in database " & adodbConnection.Properties("Data Source") & _
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
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Sub
Sub DeleteAllCFTCDataFromDatabaseByDate()
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_Description = "Deletes all data from each database available that is greater than or equal to a user-inputted date."
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim wantedDate As Date, reportType As Variant, combinedType As Variant
    
    wantedDate = InputBox("Input date for which all data >= will be deleted in the format YYYY,MM,DD (year,month,day).")
    
    If MsgBox("Is this date correct? " & wantedDate, vbYesNo) = vbYes Then
    
        For Each reportType In Array("L", "D", "T")
            For Each combinedType In Array(True, False)
                DeleteCftcDataFromSpecificDatabase wantedDate, CStr(reportType), CBool(combinedType)
            Next
        Next
        
    End If
    
End Sub
Sub DeleteCftcDataFromSpecificDatabase(smallest_date As Date, reportType As String, retrieveCombinedData As Boolean)

    Dim SQL As String, table_name As String, adodbConnection As Object, combined_wb_bool As Boolean
    
    Set adodbConnection = CreateObject("ADODB.Connection")

    database_details retrieveCombinedData, reportType, adodbConnection, table_name
    
    On Error GoTo No_Table
    SQL = "DELETE FROM " & table_name & " WHERE [Report_Date_as_YYYY-MM-DD] >= Cdate('" & Format(smallest_date, "YYYY-MM-DD") & "');"

    With adodbConnection
        .Open
        .Execute SQL, , adExecuteNoRecords
        .Close
    End With

    Set adodbConnection = Nothing

    Exit Sub
    
No_Table:
    
    MsgBox "TableL " & table_name & " not found within database."
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
    
End Sub

Public Function Latest_Date(reportType As String, combined_wb_bool As Boolean, ICE_Query As Boolean, ByRef databaseExists As Boolean) As Date
'===================================================================================================================
'Returns the date for the most recent data within a database
'===================================================================================================================
    Dim table_name As String, SQL As String, adodbConnection As Object, record As Object, var_str As String
    
    Const filter As String = "('Cocoa','B','RC','G','Wheat','W');"
    
    On Error GoTo Connection_Unavailable

    Set adodbConnection = CreateObject("ADODB.Connection")
    
    database_details combined_wb_bool, reportType, adodbConnection, table_name, , databaseExists

    If Not databaseExists Then
        Set adodbConnection = Nothing
        Latest_Date = 0
        Exit Function
    End If
    
    If Not ICE_Query Then var_str = "NOT "
    
    SQL = "SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & table_name & _
    " WHERE " & var_str & "[CFTC_Contract_Market_Code] IN " & filter

    With adodbConnection
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

    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Function
Sub UpdateDatabasePrices(data As Variant, reportType As String, combined_wb_bool As Boolean, price_column As Byte)
'===================================================================================================================
'Updates database with price data from a given array. Array should come from a worksheet
'===================================================================================================================
    Dim SQL As String, table_name As String, X As Integer, adodbConnection As Object, price_update_command As Object, CC_Column As Byte
    
    Const date_column As Byte = 1
    
    CC_Column = price_column - 1

    Set adodbConnection = CreateObject("ADODB.Connection")

    database_details combined_wb_bool, reportType, adodbConnection, table_name

    SQL = "UPDATE " & table_name & _
        " SET [Price] = ? " & _
        " WHERE [CFTC_Contract_Market_Code] = ? AND [Report_Date_as_YYYY-MM-DD] = ?;"
    
    adodbConnection.Open
    
    Set price_update_command = CreateObject("ADODB.Command")

    With price_update_command
    
        .ActiveConnection = adodbConnection
        .CommandType = adCmdText
        .CommandText = SQL
        .Prepared = True
        
        With .Parameters
            .Append price_update_command.CreateParameter("Price", adDouble, adParamInput, 20)
            .Append price_update_command.CreateParameter("Contract Code", adChar, adParamInput, 6)
            .Append price_update_command.CreateParameter("Date", adDBDate, adParamInput, 8)
        End With
        
    End With

    For X = LBound(data, 1) To UBound(data, 1)

        On Error GoTo Exit_Code
        
        With price_update_command

            With .Parameters
            
                If Not IsEmpty(data(X, price_column)) Then
                    .Item("Price").value = data(X, price_column)
                Else
                    .Item("Price").value = Null
                End If
                
                .Item("Contract Code").value = data(X, CC_Column)
                .Item("Date").value = data(X, date_column)
                
            End With
            
            .Execute
            
        End With
        
    Next X
    
Exit_Code:

    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
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
    reportType As String, Price_Symbols As Collection, contractCode As String, _
    Source_Ws As Worksheet, D As Byte, current_Filters() As Variant, LO As ListObject, Price_Data_Found As Boolean
    
    Const legacy_initial As String = "L"
    
    For Each WS In Array(LC, DC, TC)
        
        If ThisWorkbook.ActiveSheet Is WS Then
            reportType = Array("L", "D", "T")(D)
            Set Source_Ws = WS
            Exit For
        End If
        D = D + 1
    Next WS
    
    If Source_Ws Is Nothing Then Exit Sub
    
    Set LO = Source_Ws.ListObjects(reportType & "_Data")
    
    price_column = Evaluate("=VLOOKUP(""" & reportType & """,Report_Abbreviation,5,FALSE)") + 1
    
    With LO.DataBodyRange
        Worksheet_Data = .Resize(.Rows.count, price_column).value
    End With
    
    contractCode = Worksheet_Data(1, price_column - 1)
    
    Set Price_Symbols = ContractDetails
    
    If HasKey(Price_Symbols, contractCode) Then
    
        Retrieve_Tuesdays_CLose Worksheet_Data, price_column, Price_Symbols(contractCode), overwrite_all_prices:=True, dates_in_column_1:=True, Data_Found:=Price_Data_Found
        
        If Price_Data_Found Then
            
            Price_Data_Found = False
            
            'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
            UpdateDatabasePrices Worksheet_Data, legacy_initial, combined_wb_bool:=True, price_column:=price_column
            
            'Overwrites all other database tables with price data from Legacy_Combined
            
            overwrite_with_legacy_combined_prices contractCode
            
            ChangeFilters LO, current_Filters
                
            LO.DataBodyRange.columns(price_column).value = WorksheetFunction.Index(Worksheet_Data, 0, price_column)
            
            RestoreFilters LO, current_Filters
        Else
            MsgBox "Unable to retrieve data."
        End If
        
    Else
        MsgBox "A symbol is unavailable for: [ " & contractCode & " ] on worksheet " & Symbols.name & "."
    End If
    
End Sub

Sub overwrite_with_legacy_combined_prices(Optional specific_contract As String = ";", Optional minimum_date As Variant)
'===========================================================================================================
' Overwrites a given table found within a database with price data from the legacy combined table in the legacy database
'===========================================================================================================
    Dim SQL As String, table_name As String, adodbConnection As Object, legacy_database_path As String
      
    Dim reportType As Variant, retrieveCombinedData As Variant, contract_filter As String
        
    Const legacy_initial As String = "L"
    
    On Error GoTo Close_Connections

    database_details True, legacy_initial, databasePath:=legacy_database_path
    
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
    
    For Each reportType In Array(legacy_initial, "D", "T")
        
        For Each retrieveCombinedData In Array(True, False)
            
            If retrieveCombinedData = True Then
                'Related Report tables currently share the same database so only 1 connecton is needed between the 2
                Set adodbConnection = CreateObject("ADODB.Connection")
                Call database_details(CBool(retrieveCombinedData), CStr(reportType), adodbConnection)
                adodbConnection.Open
                
            End If
            
            If Not (reportType = legacy_initial And retrieveCombinedData = True) Then
                
                database_details CBool(retrieveCombinedData), CStr(reportType), table_name:=table_name
            
                SQL = "UPDATE " & table_name & _
                    " as T INNER JOIN [" & legacy_database_path & "].Legacy_Combined as F ON (F.[Report_Date_as_YYYY-MM-DD] = T.[Report_Date_as_YYYY-MM-DD] AND T.[CFTC_Contract_Market_Code] = F.[CFTC_Contract_Market_Code])" & _
                    " SET T.[Price] = F.[Price]" & contract_filter
                
                adodbConnection.Execute SQL, , adExecuteNoRecords

            End If
            
        Next retrieveCombinedData
        
        adodbConnection.Close
        Set adodbConnection = Nothing
        
    Next reportType

Close_Connections:
    
    If Not adodbConnection Is Nothing Then
        With adodbConnection
            If .State = adStateOpen Then .Close
        End With
        Set adodbConnection = Nothing
    End If

End Sub


Sub Replace_All_Prices()
Attribute Replace_All_Prices.VB_Description = "Retrieves price data for all available contracts where a price symbol is available and uploads it to each database."
Attribute Replace_All_Prices.VB_ProcData.VB_Invoke_Func = " \n14"
'=======================================================================================================
'For every contract code for which a price symbol is available, query new prices and upload to every database
'=======================================================================================================
    Dim Symbol_Info As Collection, CO As Variant, SQL As String, adodbConnection As Object, New_Data_Available As Boolean, _
    table_name As String, record As Object, data() As Variant

    Const legacy_initial As String = "L"
    Const combined_Bool As Boolean = True
    Const price_column As Byte = 3
    
    If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then
        Exit Sub
    End If
    
    On Error GoTo Close_Connection
    
    Set Symbol_Info = ContractDetails

    Set adodbConnection = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    database_details combined_Bool, legacy_initial, adodbConnection, table_name
    
    adodbConnection.Open
    
    For Each CO In Symbol_Info
        
        If CO.PriceSymbol <> vbNullString Then
    
            SQL = "SELECT [Report_Date_as_YYYY-MM-DD],[CFTC_Contract_Market_Code],[Price] FROM " & table_name & " WHERE [CFTC_Contract_Market_Code] = '" & CO.contractCode & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
            
            With record
            
                .Open SQL, adodbConnection
                
                If Not .EOF And Not .BOF Then
                
                    data = TransposeData(.GetRows)
                    .Close
                    
                    Call Retrieve_Tuesdays_CLose(data, price_column, Symbol_Info(CO.contractCode), overwrite_all_prices:=True, dates_in_column_1:=True, Data_Found:=New_Data_Available)
                    
                    If New_Data_Available Then
                        
                        New_Data_Available = False
                        
                        Call UpdateDatabasePrices(data, legacy_initial, combined_wb_bool:=True, price_column:=price_column)
                        
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
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Sub
Public Function change_table_data(LO As ListObject, retrieve_combined_data As Boolean, reportType As String, contract_code As String, triggered_by_linked_charts As Boolean) As Variant
'===================================================================================================================
'Retrieves data and updates a given listobject
'===================================================================================================================
    Dim data() As Variant, Script_2_Run As String, Last_Calculated_Column As Integer, _
    First_Calculated_Column As Byte, table_filters() As Variant, Data_Variables As Range
    
    Dim DebugTasks As New TimerC, queryForCode As String
    
    Const calculateFieldTask As String = "Calculations", outputToSheetTask As String = "Output to worksheet."
    
    Const resetUsedRangeTask As String = "Reset used range." ',applyFiltersTask As String = "Re-apply worksheet filters."
    
    queryForCode = "Query database for (" & contract_code & ")"
    
    DebugTasks.description = "Retrieve data from database and place on worksheet."
    
    Set Data_Variables = COT_ABR_Match(reportType)
    
    data = Data_Variables.value
    
    First_Calculated_Column = 3 + data(1, 5) 'Raw data coluumn count + (price) + (Empty) + (start)
    Last_Calculated_Column = data(1, 3)
    
    With DebugTasks
    
        .StartTask queryForCode
        
         data = QueryDatabaseForContract(reportType, retrieve_combined_data, contract_code)
         
        .EndTask queryForCode
    
        ReDim Preserve data(1 To UBound(data, 1), 1 To Last_Calculated_Column)
            
        .StartTask calculateFieldTask
        
        Select Case reportType
            Case "L":
                data = Legacy_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
            Case "D":
                data = Disaggregated_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
            Case "T":
                data = TFF_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26, 52)
            
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
            .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
        End With
        
        'DebugTasks.StartTask resetUsedRangeTask
        
        .Resize .Range.Resize(UBound(data, 1) + 1, .Range.columns.count)
        
        Reset_Worksheet_UsedRange .Range
        
        'DebugTasks.EndTask resetUsedRangeTask
        
    End With
    
    DebugTasks.EndTask outputToSheetTask
    
    With LO.Sort
        If .SortFields.count > 0 Then .Apply
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
    
    'Debug.Print DebugTasks.ToString
    
    Application.Calculation = xlCalculationAutomatic
    
End Function
Public Sub Manage_Table_Visual(reportType As String, Calling_Worksheet As Worksheet)
'==================================================================================================
'This sub is used to update the GUI after contracts have been updated upon activation of the calling worksheet
'==================================================================================================
    Dim Current_Details() As Variant
    
    Current_Details = COT_ABR_Match(reportType).value
    
    If Current_Details(1, 8) = True Then
        Call change_table_data(Calling_Worksheet.ListObjects(reportType & "_Data"), CBool(Current_Details(1, 7)), reportType, CStr(Current_Details(1, 9)), False)
    End If
    
End Sub
Sub Latest_Contracts()
Attribute Latest_Contracts.VB_Description = "Queries available databases for the latest contracts in a specified timeframe."
Attribute Latest_Contracts.VB_ProcData.VB_Invoke_Func = " \n14"
'=======================================================================================================
' Queries the database for the latest contracts within the database.
'=======================================================================================================
    Dim L_Table As String, L_Path As String, D_Path As String, D_Table As String, queryAvailable As Boolean
     
    Dim SQL_2 As String, date_cutoff As String, connectionString As String, QT As QueryTable

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
     
        For Each QT In .QueryTables
            If QT.name Like queryName & "*" Then
                queryAvailable = True
                Exit For
            End If
        Next QT
        
        If Not queryAvailable Then
            Set QT = .QueryTables.Add(connectionString, .Range("G1"))
        End If
        
    End With
    
    With QT
    
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
        
        AfterEventHolder.HookUpLatestContracts QT
        
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
            
            .Resize .Range.Cells(1, 1).Resize(UBound(results, 1) + 1, .ListColumns.count)
    
        End With

    End If
    
    '=IF(COUNTIF(Symbols_TBL[Contract Code -CFTC],[@[Contract Code]]),true,False)
    
End Sub
Sub Interpolator(inputStr As String, ParamArray values() As Variant)
'=======================================================================================================
' Replaces text within {} with a value in the paramArray values.
'=======================================================================================================
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

Function GetAllContractDataInFavorites(reportType As String, getCombinedData As Boolean, _
                        minWeeks As Long, Optional useAlternateCodes As Boolean = False, _
                        Optional alternateCodes As Variant) As Collection

    Dim SQL As String, tableName As String, adodbConnection As Object, record As Object, SQL2 As String, _
    favoritedContractCodes As String, queryResult() As Variant, fieldNames As String, _
    contractClctn As Collection, allContracts As New Collection, databaseExists As Boolean
     
    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]", dateColumn As Byte = 1
    
    If Not useAlternateCodes Then
        ' Get a list of all contract codes that have been favorited.
        queryResult = WorksheetFunction.Transpose(Variable_Sheet.ListObjects("Current_Favorites").DataBodyRange.columns(1).value)
    Else
        ' Use a supplied list of contract codes.
        queryResult = alternateCodes
    End If
    
    favoritedContractCodes = Join(QuotedForm(queryResult, "'"), ",")
          
    Set adodbConnection = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    Call database_details(getCombinedData, reportType, adodbConnection, tableName, , databaseExists) 'Generates a connection string and assigns a table to modify
    
    If Not databaseExists Then Exit Function
    
    With adodbConnection
        .Open
        Set record = .Execute(CommandText:=tableName, Options:=adCmdTable) 'This record will be used to retrieve field names
    End With
    
    fieldNames = FilterColumnsAndDelimit(FieldsFromRecordSet(record, encloseFieldsInBrackets:=True), reportType, includePriceColumn:=False)   'Field names from database returned as an array
    record.Close
'
    SQL2 = "SELECT " & codeField & " FROM " & tableName & " WHERE " & dateField & " = CDATE('" & Format(Variable_Sheet.Range("Most_Recently_Queried_Date").value, "yyyy-mm-dd") & "') AND " & codeField & " in (" & favoritedContractCodes & ");"
    
    'datefilter = "AND " & dateField & " >= CDATE('" & Format(DateAdd("ww", -minWeeks, Variable_Sheet.Range("Most_Recently_Queried_Date").value), "yyyy-mm-dd") & _
    "')
    
    SQL = "SELECT " & fieldNames & " FROM " & tableName & _
    " WHERE " & codeField & " in (" & SQL2 & ") Order BY " & codeField & " ASC," & dateField & " ASC;"
    
    Erase queryResult
    
    With record
        .Open SQL, adodbConnection, adOpenStatic, adLockReadOnly, adCmdText
        queryResult = TransposeData(.GetRows)
        .Close
    End With
    
    adodbConnection.Close
    
    Dim codeColumn As Byte, nameColumn As Byte, rowIndex As Long, columnIndex As Byte, _
    queryRow() As Variant, cc As Variant, output As New Collection
    
    codeColumn = UBound(queryResult, 2)
    nameColumn = 2
    
    ReDim queryRow(1 To codeColumn)
    
    With allContracts
        'Group contracts into separate collections for further processing
        For rowIndex = LBound(queryResult, 1) To UBound(queryResult, 1)
        
            For columnIndex = 1 To codeColumn
                queryRow(columnIndex) = queryResult(rowIndex, columnIndex)
            Next columnIndex
        
            On Error GoTo Create_Contract_Collection
            Set contractClctn = .Item(queryRow(codeColumn))
            
            On Error GoTo Duplicate_Date_Found
            ' Use dates as a key
            contractClctn.Add queryRow, CStr(queryRow(dateColumn))
Next_QueryRow_Iterator:
        Next rowIndex
        
        Erase queryResult

    End With
    
    With output
        For rowIndex = 1 To allContracts.count
            .Add Multi_Week_Addition(allContracts(rowIndex), Append_Type.Multiple_1d), allContracts(rowIndex)(1)(codeColumn)
        Next rowIndex
    End With
    
    Set GetAllContractDataInFavorites = output
    
    Exit Function
    
Create_Contract_Collection:

    Set contractClctn = New Collection
    allContracts.Add contractClctn, queryRow(codeColumn)
    
    Resume Next
Duplicate_Date_Found:

    Debug.Print "Duplicate found " & queryRow(1) & " " & queryRow(nameColumn) & "   " & queryRow(codeColumn)
    Resume Next_QueryRow_Iterator
End Function

Public Sub Generate_Database_Dashboard(callingWorksheet As Worksheet, ReportChr As String)

    Dim contractClctn As Collection, tempData As Variant, output() As Variant, totalStoch() As Variant, _
    outputRow As Integer, tempRow As Integer, tempCol As Byte, commercialNetColumn As Byte, _
    dateRange As Integer, Z As Byte, targetColumn As Integer, queryFutOnly As Boolean
    
    Dim dealerNetColumn As Byte, assetNetColumn As Byte, levFundNet As Byte, otherNet As Byte, _
    nonCommercialNetColumn As Byte, totalNetColumns As Byte, _
    reportGroup As Variant, reportedGroups() As Variant, producerNet As Byte, swapNet As Byte, managedNet As Byte
    
    Const threeYearsInWeeks As Integer = 156, sixMonthsInWeeks As Byte = 26, oneYearInWeeks As Byte = 52, _
    previousWeeksToCalculate As Byte = 1
    
    On Error GoTo No_Data
    
    If callingWorksheet.Shapes("FUT Only").OLEFormat.Object.value = 1 Then queryFutOnly = True
    
    Set contractClctn = GetAllContractDataInFavorites(ReportChr, Not queryFutOnly, threeYearsInWeeks + previousWeeksToCalculate + 2)
    
    With contractClctn
        If .count = 0 Then Exit Sub
        ReDim output(1 To .count, 1 To callingWorksheet.ListObjects("Dashboard_Results" & ReportChr).ListColumns.count)
    End With
    
    On Error GoTo 0
    
    Select Case ReportChr
        Case "L"
            totalNetColumns = 2
            commercialNetColumn = UBound(contractClctn(1), 2) + 1
            nonCommercialNetColumn = commercialNetColumn + 1
            totalStoch = Array(3, 7, 8, commercialNetColumn, 4, 5, nonCommercialNetColumn)
            reportedGroups = Array(commercialNetColumn, nonCommercialNetColumn)
        Case "D"
            totalNetColumns = 4
            producerNet = UBound(contractClctn(1), 2) + 1
            swapNet = producerNet + 1
            managedNet = swapNet + 1
            otherNet = managedNet + 1
            totalStoch = Array(3, 4, 5, producerNet, 6, 7, swapNet, 9, 10, managedNet, 12, 13, otherNet)
            reportedGroups = Array(producerNet, swapNet, managedNet, otherNet)
        Case "T"
            totalNetColumns = 4
            dealerNetColumn = UBound(contractClctn(1), 2) + 1
            assetNetColumn = dealerNetColumn + 1
            levFundNet = assetNetColumn + 1
            otherNet = levFundNet + 1
            totalStoch = Array(3, 4, 5, dealerNetColumn, 7, 8, assetNetColumn, 10, 11, levFundNet, 13, 14, otherNet)
            reportedGroups = Array(dealerNetColumn, assetNetColumn, levFundNet, otherNet)
    End Select

    For Each tempData In contractClctn
        
        contractClctn.Remove tempData(1, UBound(tempData, 2))

        outputRow = outputRow + 1
        'Contract name without exchange name.
        output(outputRow, 1) = Left$(tempData(UBound(tempData, 1), 2), InStrRev(tempData(UBound(tempData, 1), 2), "-") - 2)
        
        ReDim Preserve tempData(1 To UBound(tempData, 1), 1 To UBound(tempData, 2) + totalNetColumns)
        'Net Position calculations.
        For tempRow = LBound(tempData, 1) To UBound(tempData, 1)

            Select Case ReportChr
                Case "L"
                    tempData(tempRow, commercialNetColumn) = tempData(tempRow, 7) - tempData(tempRow, 8)
                    tempData(tempRow, nonCommercialNetColumn) = tempData(tempRow, 4) - tempData(tempRow, 5)
                Case "D"
                    tempData(tempRow, producerNet) = tempData(tempRow, 4) - tempData(tempRow, 5)
                    tempData(tempRow, swapNet) = tempData(tempRow, 6) - tempData(tempRow, 7)
                    tempData(tempRow, managedNet) = tempData(tempRow, 9) - tempData(tempRow, 10)
                    tempData(tempRow, otherNet) = tempData(tempRow, 12) - tempData(tempRow, 13)
                Case "T"
                    tempData(tempRow, dealerNetColumn) = tempData(tempRow, 4) - tempData(tempRow, 5)
                    tempData(tempRow, assetNetColumn) = tempData(tempRow, 7) - tempData(tempRow, 8)
                    tempData(tempRow, levFundNet) = tempData(tempRow, 10) - tempData(tempRow, 11)
                    tempData(tempRow, otherNet) = tempData(tempRow, 13) - tempData(tempRow, 14)
            End Select
            
        Next tempRow
        'Index calculations using all available data.
        For Z = LBound(totalStoch) To UBound(totalStoch)
            targetColumn = totalStoch(Z)
            output(outputRow, 2 + Z) = Stochastic_Calculations(targetColumn, UBound(tempData, 1), tempData, previousWeeksToCalculate, True)(1)
        Next Z
        'Variable Index calculations.
        'tempCol is used to track the column that correlates with the given calculation.
        tempCol = 3 + UBound(totalStoch)
        For Each reportGroup In reportedGroups
            
            For Z = 0 To 2
                dateRange = Array(threeYearsInWeeks, oneYearInWeeks, sixMonthsInWeeks)(Z)
                If UBound(tempData, 1) >= dateRange Then
                    output(outputRow, tempCol) = Stochastic_Calculations(CInt(reportGroup), dateRange, tempData, previousWeeksToCalculate, True)(1)
                End If
                tempCol = tempCol + 1
            Next Z
            
        Next reportGroup
        
    Next tempData
    
    On Error GoTo 0
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    With callingWorksheet
    
        .Range("A1").value = Variable_Sheet.Range("Most_Recently_Queried_Date").value
        
        With .ListObjects("Dashboard_Results" & ReportChr)
            
            With .DataBodyRange
                .ClearContents
                With .Resize(UBound(output, 1), UBound(output, 2))
                    .value = output
                    .Sort key1:=.columns(1), Orientation:=xlSortColumns, ORder1:=xlAscending, header:=xlNo, MatchCase:=False
                End With
            End With
            
            If UBound(output, 1) <> .ListRows.count Then
                .Resize .Range.Resize(UBound(output, 1) + 1, .ListColumns.count)
            End If
            
        End With

    End With
    
    Re_Enable
    
    Exit Sub
    
No_Data:
    MsgBox "An error occurred. " & Err.description
End Sub


Public Function Assign_Charts_WS(reportType As String) As Worksheet
    
    Dim WSA() As Variant, T As Byte
    
    WSA = Array(L_Charts, D_Charts, T_Charts)
    
    T = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
    
    Set Assign_Charts_WS = WSA(T)

End Function

Public Function Assign_Linked_Data_Sheet(reportType As String) As Worksheet
'=======================================================================================================
' Returns a COT data worksheet based on the value of reportType.
'=======================================================================================================
    Dim WSA() As Variant, T As Byte
    
    WSA = Array(LC, DC, TC)
    
    T = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
    
    Set Assign_Linked_Data_Sheet = WSA(T)
    
End Function
    
Public Sub Save_For_Github()
'=======================================================================================================
' Toggles range value that marks the workboook for upload to github.
'=======================================================================================================
    If UUID Then
        Range("Github_Version").value = True
        Custom_SaveAS
    End If

End Sub
Public Sub Handle_Contract_GUI()
'=======================================================================================================
' Unloads the Contract_Selection userform if not on a data worksheet.
'=======================================================================================================
    Dim WS As Worksheet, Userform_Active As Boolean
    
    Set WS = ThisWorkbook.ActiveSheet
    
    Userform_Active = IsLoadedUserform("Contract_Selection")
    
    Select Case True
        
        Case WS Is LC, WS Is TC, WS Is DC
        
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
'=======================================================================================================
' Loops through worksheets containing COT data and moves the Contract_Selection launching shape to a point
' on the worksheet.
'=======================================================================================================
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
    Dim Symbols As Collection, SQL As String, adodbConnection As Object, tableName As String, queryResult() As Variant, cc As Integer

    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]"
    
    Dim codeColumn As Byte, rowIndex As Long, columnIndex As Byte, contractClctn As Collection, _
    queryRow() As Variant, allContracts As New Collection, minDate As String, succesfulPriceRetrieval As Boolean
    
    minDate = InputBox("Input date in form YYYY-MM-DD")

    Set adodbConnection = CreateObject("ADODB.COnnection")
    
    database_details True, "L", adodbConnection, tableName
    
    SQL = "SELECT " & Join(Array(dateField, codeField, "Price"), ",") & " FROM " & tableName & " WHERE " & dateField & " >=Cdate('" & minDate & "');"
    
    codeColumn = 2

    With adodbConnection
        .Open
         queryResult = TransposeData(.Execute(SQL, , adCmdText).GetRows)
        .Close
    End With
    
    Set Symbols = ContractDetails
    
    ReDim queryRow(1 To UBound(queryResult, 2))
    
    With allContracts
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
    
    With allContracts
    
        For cc = .count To 1 Step -1
            
            Set contractClctn = .Item(cc)
            
            queryResult = Multi_Week_Addition(contractClctn, Append_Type.Multiple_1d)
            
            .Remove queryResult(1, codeColumn)
            
            If HasKey(Symbols, CStr(queryResult(1, codeColumn))) Then
            
                Retrieve_Tuesdays_CLose queryResult, 3, Symbols(queryResult(1, codeColumn)), True, True, succesfulPriceRetrieval
                
                If succesfulPriceRetrieval Then .Add queryResult, queryResult(1, codeColumn)
            
            End If
        
        Next cc
    
    End With
    
    queryResult = Multi_Week_Addition(allContracts, Append_Type.Multiple_2d)
    
    On Error GoTo 0
    
    UpdateDatabasePrices queryResult, "L", True, 3
    
    overwrite_with_legacy_combined_prices minimum_date:=CDate(minDate)
    
    Exit Sub
    
Create_Contract_Collection:

    Set contractClctn = New Collection
    allContracts.Add contractClctn, queryRow(codeColumn)
    
    Resume Next

End Sub
Sub FindDatabasePathInSameFolder()
'===========================================================================================================
' Looks for MS Access Database files that haven't been renamed within the same folder as the Excel workbook.
'===========================================================================================================
    Dim Legacy As New LoadedData, TFF As New LoadedData, DGG As New LoadedData, _
    strfile As String, foundCount As Byte, folderPath As String
    
    On Error GoTo Prompt_User_About_UserForm
    
    Legacy.InitializeClass "L"
    DGG.InitializeClass "D"
    TFF.InitializeClass "T"
    
    folderPath = ThisWorkbook.Path & Application.PathSeparator
    ' Filter for Microsoft Access databases.
    strfile = Dir(folderPath & "*.accdb")
    
    Do While Len(strfile) > 0
        
        If LCase(strfile) Like "*disaggregated.accdb" And IsEmpty(DGG.CurrentDatabasePath) Then
            DGG.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        ElseIf LCase(strfile) Like "*legacy.accdb" And IsEmpty(Legacy.CurrentDatabasePath) Then
            Legacy.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        ElseIf LCase(strfile) Like "*tff.accdb" And IsEmpty(TFF.CurrentDatabasePath) Then
            TFF.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        End If
        
        strfile = Dir
    Loop
    
Prompt_User_About_UserForm:

    If foundCount <> 3 And Not UUID Then
        MsgBox "Database paths couldn't be auto-retrieved." & vbNewLine & vbNewLine & _
        "Please use the Database Paths USerform to fill in the needed data."
    End If
    
End Sub
