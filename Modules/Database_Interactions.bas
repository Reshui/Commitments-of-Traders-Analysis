Attribute VB_Name = "Database_Interactions"
Private AfterEventHolder As ClassQTE

#Const engageTimers = False

Option Explicit

Sub database_details(getFuturesAndOptions As Boolean, reportType As String, Optional ByRef adodbConnection As Object, _
                    Optional ByRef tableNameToReturn As String, Optional ByRef databasePath As String, Optional ByRef doesDatabaseExist As Boolean = False)
'===================================================================================================================
    'Purpose: Determines if database exists. If it does the appropriate variables or properties are assigned values if needed.
    'Inputs:
    '        reportType - One of L,D,T to repersent which database to delete from.
    '        getFuturesAndOptions - true for futures+options and false for futures only.
    'Outputs:
'===================================================================================================================
    
    Dim Report_Name As String, userSpecifiedDatabasePath As String
    
    If reportType = "T" Then
        Report_Name = "TFF"
    Else
        Report_Name = GetStoredReportDetails(reportType).FullReportName.Value2
    End If
    
    If UUID Then
        databasePath = Environ$("USERPROFILE") & "\Documents\" & Report_Name & ".accdb"
        doesDatabaseExist = True
    Else

        userSpecifiedDatabasePath = Variable_Sheet.Range(reportType & "_Database_Path").Value2
            
        If LenB(userSpecifiedDatabasePath) = 0 Or Not FileOrFolderExists(userSpecifiedDatabasePath) Then
            
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
    
    If Not IsMissing(tableNameToReturn) Then tableNameToReturn = Report_Name & IIf(getFuturesAndOptions = True, "_Combined", "_Futures_Only")
    
    If Not adodbConnection Is Nothing Then adodbConnection.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath & ";"
    
End Sub
Private Function FilterColumnsAndDelimit(fieldsInDatabase As Variant, reportType As String, includePriceColumn As Boolean) As String
'===================================================================================================================
'Loops table found on Variables Worksheet that contains True/False values for wanted columns
'An array of wanted columns with some re-ordering is returned
'===================================================================================================================
    Dim wantedColumns() As Variant
    
    wantedColumns = Filter_Market_Columns(False, True, convert_skip_col_to_general:=False, reportType:=reportType, Create_Filter:=True, InputA:=fieldsInDatabase)
    
    ReDim Preserve wantedColumns(LBound(wantedColumns) To UBound(wantedColumns) + IIf(includePriceColumn = True, 1, 0))
    
    If includePriceColumn Then wantedColumns(UBound(wantedColumns)) = "Price"

    FilterColumnsAndDelimit = WorksheetFunction.TextJoin(",", True, wantedColumns)
    
End Function
Function FilteredFieldsFromRecordSet(record As Object, fieldInfoByEditedName As Collection) As Collection
        
    Dim Item As Variant, EditedName As String, output As New Collection, FI As FieldInfo
    
    On Error GoTo MissingKey
    
    For Each Item In record.Fields
    
        EditedName = EditDatabaseNames(Item.name)
        
        Set FI = fieldInfoByEditedName(EditedName)
        
        With FI
            If .IsMissing = False And Not .EditedName = "id" Then
                Call .EditDatabaseName(Item.name)
                output.Add FI, .EditedName
            End If
        End With
        
AttemptNextField:
    Next Item
    
    Set FilteredFieldsFromRecordSet = output
    Exit Function
    
MissingKey:
    Resume AttemptNextField
End Function
Function FieldsFromRecordSet(record As Object, encloseFieldsInBrackets As Boolean) As Variant
'===================================================================================================================
'record is a RecordSET object containing a single row of data from which field names are retrieved,formatted and output as an array
'===================================================================================================================
    Dim X As Integer, Z As Byte, fieldNamesInRecord() As Variant, currentFieldName As String

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

Function QueryDatabaseForContract(reportType As String, getFuturesAndOptions As Boolean, wantedContractCode As String) As Variant
'===================================================================================================================
'Retrieves filtered data from database and returns as an array
'===================================================================================================================
    Dim record As Object, adodbConnection As Object, tableNameWithinDatabase As String

    Dim SQL As String, delimitedWantedColumns As String, allFieldNames() As Variant

    On Error GoTo Close_Connection
    
    Application.StatusBar = "Querying database for " & wantedContractCode
    Set adodbConnection = CreateObject("ADODB.Connection")

    database_details getFuturesAndOptions, reportType, adodbConnection, tableNameWithinDatabase

    With adodbConnection
        .Open
        Set record = .Execute(tableNameWithinDatabase, , adCmdTable)
    End With
    
    allFieldNames = FieldsFromRecordSet(record, encloseFieldsInBrackets:=True)
    
    record.Close
    
    delimitedWantedColumns = FilterColumnsAndDelimit(allFieldNames, reportType:=reportType, includePriceColumn:=True)

    SQL = "SELECT " & delimitedWantedColumns & " FROM " & tableNameWithinDatabase & " WHERE [CFTC_Contract_Market_Code]='" & wantedContractCode & "' ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
    
    With record
        .Open SQL, adodbConnection
         QueryDatabaseForContract = TransposeData(.GetRows)
    End With
    
Close_Connection:
    
    Application.StatusBar = vbNullString
    
    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Function

Public Sub Update_Database(dataToUpload As Variant, uploadingFuturesAndOptions As Boolean, reportType As String, debugOnly As Boolean, fieldInfoByEditedName As Collection)
'===================================================================================================================
    'Purpose: Uploads rows contained within dataToUpload to a database determined by other Parameters.
    'Inputs:
    '       dataToUpload  - 2D array of rows to be uploaded.
    '       uploadingFuturesAndOptions - True if data being uploaded is Futures + Options combined.
    '       reportType - One of L,D,T to repersent which database to upload to.
    '       fieldInfoByEditedName - A Collection of FieldInfo instances used to describe columns contained within dataToUpload.
'===================================================================================================================
   
    Dim tableToUpdateName As String, wantedDatabaseFields As Collection, row As Long, _
    legacyCombinedTableName As String, legacyDatabasePath As String, useYear3000 As Boolean, _
    uploadingLegacyCombinedData As Boolean, oldestDateInUpload As Date, uploadToDatabase As Boolean
    
    If debugOnly Then
        If MsgBox("Debug Active: Do you want to upload data to databse?", vbYesNo) = vbYes Then uploadToDatabase = True
        If MsgBox("Replace dates with year 3000?", vbYesNo) = vbYes Then useYear3000 = True
    Else
        uploadToDatabase = True
    End If
    
    Dim databaseFieldNamesRecord As Object, adodbConnection As Object, SQL As String
    
    On Error GoTo Close_Connection
    
    Const legacy_abbreviation As String = "L"

    Set adodbConnection = CreateObject("ADODB.Connection")
    Set databaseFieldNamesRecord = CreateObject("ADODB.RecordSet")
    
    If reportType = legacy_abbreviation And uploadingFuturesAndOptions = True Then uploadingLegacyCombinedData = True

    Call database_details(uploadingFuturesAndOptions, reportType, adodbConnection, tableToUpdateName)   'Generates a connection string and assigns a table to modify

    With adodbConnection
        '.CursorLocation = adUseServer                                   'Batch update won't work otherwise
        .Open
        'Get a record of all field names within the database.
        Set databaseFieldNamesRecord = .Execute(CommandText:=tableToUpdateName, Options:=adCmdTable)
    End With
    ' Get a ccollection of FieldInfo instances with matching fields for input and target.
    Set wantedDatabaseFields = FilteredFieldsFromRecordSet(databaseFieldNamesRecord, fieldInfoByEditedName)
    
    databaseFieldNamesRecord.Close
    
    Dim uploadCommand As Object, FI As FieldInfo, Item As Variant, obj As Object
    Dim fieldNames As New Collection, fieldValues As New Collection, startedTransaction As Boolean
    
    Set uploadCommand = CreateObject("ADODB.Command")
    
    adodbConnection.BeginTrans
    startedTransaction = True
    
    With uploadCommand
    
        .ActiveConnection = adodbConnection
        .commandType = adCmdText
        .Prepared = True
         
        With .Parameters
            For Each FI In wantedDatabaseFields
            
                Select Case FI.DataType
                    Case adNumeric, adCurrency
                    
                        Set obj = uploadCommand.CreateParameter(FI.EditedName, FI.DataType, adParamInput)
                        
                        With obj
                            .NumericScale = 5
                            .Precision = 15
                        End With
                        
                        .Append obj
                        
                    Case Else
                        .Append uploadCommand.CreateParameter(FI.EditedName, FI.DataType, adParamInput)
                End Select

                fieldValues.Add "?"
                fieldNames.Add FI.DatabaseNameForSQL
                
            Next FI
            
        End With
        
        .CommandText = "Insert Into " + tableToUpdateName + "(" + Join(ConvertCollectionToArray(fieldNames), ",") + ") Values (" + Join(ConvertCollectionToArray(fieldValues), ",") + ");"
        Set fieldValues = Nothing: Set fieldNames = Nothing
        Dim wantedColumn As Byte
        
        For row = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
        
            For Each Item In uploadCommand.Parameters
            
                wantedColumn = wantedDatabaseFields(Item.name).ColumnIndex

                If Not (IsError(dataToUpload(row, wantedColumn)) Or IsEmpty(dataToUpload(row, wantedColumn))) Then
                    
                    If IsNumeric(dataToUpload(row, wantedColumn)) Then
                        Item.value = dataToUpload(row, wantedColumn)
                    ElseIf dataToUpload(row, wantedColumn) = "." Or LenB(Trim$(dataToUpload(row, wantedColumn))) = 0 Then
                        Item.value = Null
                    Else
                        Item.value = dataToUpload(row, wantedColumn)
                    End If
                    
                    If Item.Type = adDate And (oldestDateInUpload = TimeSerial(0, 0, 0) Or dataToUpload(row, wantedColumn) < oldestDateInUpload) Then
                        oldestDateInUpload = dataToUpload(row, wantedColumn)
                    End If
                Else
                    Item.value = Null
                End If

            Next Item
            
            If useYear3000 Then
                uploadCommand.Parameters("report_date_as_yyyy_mm_dd").value = DateSerial(3000, 1, 1)
            End If
            
            If uploadToDatabase Then .Execute
        Next row
        
    End With
    
    adodbConnection.CommitTrans
    startedTransaction = False
    Set uploadCommand = Nothing
    
    If uploadToDatabase And Not uploadingLegacyCombinedData Then 'retrieve price data from the legacy combined table
        'Legacy COmbined Data should be the first data retrieved
        Call database_details(True, legacy_abbreviation, tableNameToReturn:=legacyCombinedTableName, databasePath:=legacyDatabasePath)
    
        'T alias is for table that is being updated
        SQL = "Update " & tableToUpdateName & " as T INNER JOIN [" & legacyDatabasePath & "]." & legacyCombinedTableName & " as Source_TBL ON Source_TBL.[Report_Date_as_YYYY-MM-DD]=T.[Report_Date_as_YYYY-MM-DD] AND Source_TBL.[CFTC_Contract_Market_Code]=T.[CFTC_Contract_Market_Code]" & _
            " SET T.[Price] = Source_TBL.[Price] WHERE T.[Report_Date_as_YYYY-MM-DD]>=CDate('" & Format(oldestDateInUpload, "YYYY-MM-DD") & "');"
        
        adodbConnection.Execute CommandText:=SQL, Options:=adCmdText + adExecuteNoRecords

    End If
    
    If Not debugOnly Then
        
        With GetStoredReportDetails(reportType)
            If .UsingCombined.Value2 = uploadingFuturesAndOptions Then
                'This will signal to worksheet activate events to update the currently visible data
                .PendingUpdateInDatabase.Value2 = True
            End If
        End With
        
    End If
    
Close_Connection:
    
    If Err.Number <> 0 Then
        
        MsgBox "An error occurred while attempting to update table [ " & tableToUpdateName & " ] in database " & adodbConnection.Properties("Data Source") & _
        vbNewLine & vbNewLine & _
        "Error description: " & Err.description
        
        If startedTransaction Then adodbConnection.RollbackTrans
        
    End If

    If Not databaseFieldNamesRecord Is Nothing Then
        If databaseFieldNamesRecord.State = adStateOpen Then databaseFieldNamesRecord.Close
        Set databaseFieldNamesRecord = Nothing
    End If
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Sub
Sub DeleteAllCFTCDataFromDatabaseByDate()
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_Description = "Deletes all data from each database available that is greater than or equal to a user-inputted date."
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_ProcData.VB_Invoke_Func = " \n14"
'===================================================================================================================
    'Purpose: Asks the user for a minimum date and then deletes all data greater
    '           than or equal to that in all available datanases.
'===================================================================================================================
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
Sub DeleteCftcDataFromSpecificDatabase(smallest_date As Date, reportType As String, deleteFuturesAndOptions As Boolean)
'===================================================================================================================
    'Purpose: Deletes COT data from database that is as recent as smallest_date.
    'Inputs: smallest_date - all rows with a date value >= to this will be deleted.
    '        reportType - One of L,D,T to repersent which database to delete from.
    '        deleteFuturesAndOptions - true for futures+options and false for futures only.
    'Outputs:
'===================================================================================================================

    Dim SQL As String, tableName As String, adodbConnection As Object, adodbCommand As Object
    
    Set adodbConnection = CreateObject("ADODB.Connection")
    Set adodbCommand = CreateObject("ADODB.Command")
    
    database_details deleteFuturesAndOptions, reportType, adodbConnection, tableName
    
    On Error GoTo No_Table
    SQL = "DELETE FROM " & tableName & " WHERE [Report_Date_as_YYYY-MM-DD] >= ?;"
    
    adodbConnection.Open
        With adodbCommand
            .ActiveConnection = adodbConnection
            .CommnadText = SQL
            .commandType = adCmdText
            .Parameters.Add .CreateParameter("@smallestDate", adDate, adParamInput, value:=smallest_date)
            
            .Execute
        End With
    adodbConnection.Close
    
    Set adodbConnection = Nothing
    Set adodbCommand = Nothing
    Exit Sub
    
No_Table:
    MsgBox "TableL " & tableName & " not found within database."
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Sub

Public Function Latest_Date(reportType As String, getFuturesAndOptions As Boolean, queryIceContracts As Boolean, ByRef databaseExists As Boolean) As Date
'===================================================================================================================
'Returns the date for the most recent data within a database
'===================================================================================================================

    Dim tableName As String, SQL As String, adodbConnection As Object, record As Object, var_str As String
    
    Const filter As String = "('Cocoa','B','RC','G','Wheat','W');"
    
    On Error GoTo Connection_Unavailable

    Set adodbConnection = CreateObject("ADODB.Connection")
    
    database_details getFuturesAndOptions, reportType, adodbConnection, tableName, , databaseExists

    If Not databaseExists Then
        Set adodbConnection = Nothing
        Latest_Date = 0
        Exit Function
    End If
    
    If Not queryIceContracts Then var_str = "NOT "
    
    SQL = "SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & tableName & _
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
Sub UpdateDatabasePrices(data As Variant, reportType As String, targetFuturesAndOptions As Boolean, price_column As Byte)
'===================================================================================================================
'Updates database with price data from a given array. Array should come from a worksheet
'===================================================================================================================
    Dim SQL As String, tableName As String, X As Integer, adodbConnection As Object, price_update_command As Object, CC_Column As Byte
    
    Const date_column As Byte = 1
    
    CC_Column = price_column - 1

    Set adodbConnection = CreateObject("ADODB.Connection")

    database_details targetFuturesAndOptions, reportType, adodbConnection, tableName

    SQL = "UPDATE " & tableName & _
        " SET [Price] = ? " & _
        " WHERE [CFTC_Contract_Market_Code] = ? AND [Report_Date_as_YYYY-MM-DD] = ?;"
    
    adodbConnection.Open
    
    Set price_update_command = CreateObject("ADODB.Command")

    With price_update_command
    
        .ActiveConnection = adodbConnection
        .commandType = adCmdText
        .CommandText = SQL
        .Prepared = True
        
        With .Parameters
            .Append price_update_command.CreateParameter("Price", adCurrency, adParamInput)
            .Append price_update_command.CreateParameter("Contract Code", adBSTR, adParamInput)
            .Append price_update_command.CreateParameter("Date", adDate, adParamInput)
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
    reportType As String, availableContractInfo As Collection, ContractCode As String, _
    Source_Ws As Worksheet, D As Byte, current_Filters() As Variant, LO As ListObject, currentData As New LoadedData
    
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
    
    Set LO = CftcOutputTable(reportType)
    
    currentData.InitializeClass reportType
    
    price_column = currentData.RawDataCount + 1
    
    With LO.DataBodyRange
        Worksheet_Data = .Resize(.Rows.count, price_column).value
    End With
    
    ContractCode = currentData.CurrentContractCode
    
    Set availableContractInfo = GetAvailableContractInfo
    
    If HasKey(availableContractInfo, ContractCode) Then
    
        If TryGetPriceData(Worksheet_Data, price_column, availableContractInfo(ContractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
            
            'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
            UpdateDatabasePrices Worksheet_Data, legacy_initial, targetFuturesAndOptions:=True, price_column:=price_column
            
            'Overwrites all other database tables with price data from Legacy_Combined
            
            overwrite_with_legacy_combined_prices ContractCode
            
            ChangeFilters LO, current_Filters
                
            LO.DataBodyRange.columns(price_column).Value2 = WorksheetFunction.index(Worksheet_Data, 0, price_column)
            
            RestoreFilters LO, current_Filters
        Else
            MsgBox "Unable to retrieve data."
        End If
        
    Else
        MsgBox "A symbol is unavailable for: [ " & ContractCode & " ] on worksheet " & Symbols.name & "."
    End If
    
End Sub

Sub overwrite_with_legacy_combined_prices(Optional specific_contract As String = ";", Optional minimum_date As Variant)
'===========================================================================================================
' Overwrites a given table found within a database with price data from the legacy combined table in the legacy database
'===========================================================================================================
    Dim SQL As String, tableName As String, adodbConnection As Object, legacy_database_path As String
      
    Dim reportType As Variant, overwritingFuturesAndOptions As Variant, contract_filter As String
        
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
        
        For Each overwritingFuturesAndOptions In Array(True, False)
            
            If overwritingFuturesAndOptions = True Then
                'Related Report tables currently share the same database so only 1 connecton is needed between the 2
                Set adodbConnection = CreateObject("ADODB.Connection")
                Call database_details(CBool(overwritingFuturesAndOptions), CStr(reportType), adodbConnection)
                adodbConnection.Open
                
            End If
            
            If Not (reportType = legacy_initial And overwritingFuturesAndOptions = True) Then
                
                database_details CBool(overwritingFuturesAndOptions), CStr(reportType), tableNameToReturn:=tableName
            
                SQL = "UPDATE " & tableName & _
                    " as T INNER JOIN [" & legacy_database_path & "].Legacy_Combined as F ON (F.[Report_Date_as_YYYY-MM-DD] = T.[Report_Date_as_YYYY-MM-DD] AND T.[CFTC_Contract_Market_Code] = F.[CFTC_Contract_Market_Code])" & _
                    " SET T.[Price] = F.[Price]" & contract_filter
                
                adodbConnection.Execute SQL, , adExecuteNoRecords

            End If
            
        Next overwritingFuturesAndOptions
        
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
    Dim availableContractInfo As Collection, CO As Variant, SQL As String, adodbConnection As Object, _
    tableName As String, recordSet As Object, data() As Variant, cmd As Object

    Const legacy_initial As String = "L"
    Const combined_Bool As Boolean = True
    Const price_column As Byte = 3
    
    If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then
        Exit Sub
    End If
    
    Set availableContractInfo = GetAvailableContractInfo

    Set adodbConnection = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    database_details combined_Bool, legacy_initial, adodbConnection, tableName
    
    On Error GoTo Close_Connection
    adodbConnection.Open
    
    With cmd
    
        .commandType = adCmdText
        .CommandText = "SELECT [Report_Date_as_YYYY-MM-DD],[CFTC_Contract_Market_Code],[Price] FROM " & tableName & " WHERE [CFTC_Contract_Market_Code] = ? ORDER BY [Report_Date_as_YYYY-MM-DD] ASC;"
        .ActiveConnection = adodbConnection
        .Prepared = True

        With .Parameters
            .Append cmd.CreateParameter("@ContractCode", adBSTR, adParamInput)
        End With
        
    End With
            
    For Each CO In availableContractInfo
        
        If CO.HasSymbol Then
            
            With cmd
                .Parameters("@ContractCode").value = CO.ContractCode
                Set recordSet = .Execute
            End With
            
            With recordSet
                
                If Not .EOF And Not .BOF Then
                
                    data = TransposeData(.GetRows)
                    
                    If TryGetPriceData(data, price_column, availableContractInfo(CO.ContractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
                        Call UpdateDatabasePrices(data, legacy_initial, targetFuturesAndOptions:=True, price_column:=price_column)
                        overwrite_with_legacy_combined_prices CO.ContractCode
                    End If

                End If
                
                 .Close
                 
            End With

        End If
        
    Next CO

Close_Connection:
    
    Set cmd = Nothing
    
    If Not recordSet Is Nothing Then
        If recordSet.State = adStateOpen Then recordSet.Close
        Set recordSet = Nothing
    End If
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
End Sub
Public Sub ExchangeTableData(LO As ListObject, getFuturesAndOptions As Boolean, reportType As String, ContractCode As String, maintainCurrentTableFilters As Boolean, recalculateWorksheetFormulas As Boolean)
'===================================================================================================================
'Retrieves data and updates a given listobject
'===================================================================================================================
    Dim data() As Variant, Last_Calculated_Column As Integer, rawDataCountForReport As Byte, _
    First_Calculated_Column As Byte, table_filters() As Variant, reportDetails As LoadedData, worksheetForTable As Worksheet
    
    Dim DebugTasks As New Timers, statusBarMessage As String, appProperties As Collection, contractUnitsColumn As Range
    
    Const calculateFieldTask As String = "Calculations", outputToSheetTask As String = "Output to worksheet."
    
    Const resizeTableTask As String = "Resize Table.", clearCertainCells As String = "Clear Constants" ',applyFiltersTask As String = "Re-apply worksheet filters."
    
    Set appProperties = DisableApplicationProperties(False, True, True)
    
    statusBarMessage = "Query database for (" & ContractCode & ")"
    
    On Error GoTo Unhandled_Error_Discovered
    
    Set reportDetails = GetStoredReportDetails(reportType)
    
    With reportDetails
        rawDataCountForReport = .RawDataCount.Value2
        First_Calculated_Column = 3 + rawDataCountForReport 'Raw data coluumn count + (price) + (Empty) + (start)
        Last_Calculated_Column = .LastCalculatedColumn.Value2
    End With
    
    With DebugTasks
    
        #If engageTimers Then
        
            .description = "Retrieve data from database and place on worksheet."
            
            .StartTask statusBarMessage
            data = QueryDatabaseForContract(reportType, getFuturesAndOptions, ContractCode)
            .EndTask statusBarMessage
        #Else
            data = QueryDatabaseForContract(reportType, getFuturesAndOptions, ContractCode)
        #End If
    
        ReDim Preserve data(1 To UBound(data, 1), 1 To Last_Calculated_Column)
            
        #If engageTimers Then
            .StartTask calculateFieldTask
        #End If
        
        Select Case reportType
            Case "L":
                data = Legacy_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
            Case "D":
                data = Disaggregated_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
            Case "T":
                data = TFF_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26, 52)
        End Select
        
        #If engageTimers Then
            .EndTask calculateFieldTask
        #End If
        
    End With
        
    With LO
        
        Set worksheetForTable = .Parent
        
        worksheetForTable.DisplayPageBreaks = False
        
        ChangeFilters LO, table_filters
        
        With .DataBodyRange
            'Clear get quantity array formula
            '.columns(1).Offset(0, -1).EntireColumn.ClearContents
                        
            #If engageTimers Then
                DebugTasks.StartTask outputToSheetTask
                .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
                DebugTasks.EndTask outputToSheetTask
            #Else
                .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
            #End If
            
            .columns(1).Offset(0, -1).ClearContents
            
        End With
        
        #If engageTimers Then
            DebugTasks.StartTask resizeTableTask
            .Resize .Range.Resize(UBound(data, 1) + 1, .Range.columns.count)
            DebugTasks.EndTask resizeTableTask
        #Else
            .Resize .Range.Resize(UBound(data, 1) + 1, .Range.columns.count)
        #End If
        
        With LO.Sort
            If .SortFields.count > 0 Then .Apply
        End With
        
        With .DataBodyRange
        
            Set contractUnitsColumn = .columns(rawDataCountForReport - 1) '.columns(First_Calculated_Column - 4)
            
            With .columns(1).Offset(0, -1)
                .FormulaArray = "=GetNumbers(" & contractUnitsColumn.Address & ")"
            End With
            
        End With
        
    End With
    
    On Error GoTo Unhandled_Error_Discovered
    
    If maintainCurrentTableFilters Then
        With DebugTasks
            '.StartTask applyFiltersTask
                RestoreFilters LO, table_filters
            '.EndTask applyFiltersTask
        End With
    End If
    
    With reportDetails
        .CurrentContractCode.Value2 = ContractCode
        .PendingUpdateInDatabase.Value2 = False
    End With
    
    Const formulaCalculation As String = "Formula Calculation for Worksheet"
            
    #If engageTimers Then
    
        With DebugTasks
        
            .StartTask clearCertainCells
                ClearRegionBeneathTable LO
            .EndTask clearCertainCells
        
            If recalculateWorksheetFormulas Then
                .StartTask formulaCalculation
                    worksheetForTable.Calculate
                .EndTask formulaCalculation
            End If
            
        End With
        
    #Else
    
        ClearRegionBeneathTable LO
        
        If recalculateWorksheetFormulas Then
            worksheetForTable.Calculate
        End If
        
    #End If
        
Finally:
    
    #If engageTimers Then
        Debug.Print DebugTasks.ToString
    #End If
    
    EnableApplicationProperties appProperties
    
    Exit Sub
Unhandled_Error_Discovered:
    DisplayErrorIfAvailable Err, "ExchangeTableData()"
    Resume Finally
End Sub
Public Sub RefreshTableData(reportType As String)
'==================================================================================================
'This sub is used to update the GUI after contracts have been updated upon activation of the calling worksheet
'==================================================================================================
    Dim tableToRefresh As ListObject
    
    Set tableToRefresh = CftcOutputTable(reportType)
    
    With GetStoredReportDetails(reportType)
        If .PendingUpdateInDatabase = True Then
            Call ExchangeTableData(tableToRefresh, .UsingCombined, reportType, .CurrentContractCode, True, True)
        End If
    End With
    
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
            Set QT = .QueryTables.Add(connectionString, .Range("K1"))
        End If
        
    End With
    
    With QT
    
        .CommandText = SQL_2
        .BackgroundQuery = True
        .Connection = connectionString
        .commandType = xlCmdSql
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
        
    Dim results() As Variant, commodityGroups As Collection, Item As Variant, iRow As Integer, LO As ListObject
    Dim appProperties As Collection
    Const commodityColumn As Byte = 4, subGroupColumn As Byte = 5
    
    Set AfterEventHolder = Nothing
    
    Set appProperties = DisableApplicationProperties(True, False, True)
    
    If Success Then
        
        With RefreshedQueryTable.ResultRange
            results = .Value2
            .ClearContents
        End With
        
        Set commodityGroups = CFTC_CommodityGroupings
        
        ReDim Preserve results(LBound(results, 1) To UBound(results, 1), 1 To UBound(results, 2) + 2)
        
        On Error GoTo Code_Not_Found
        For iRow = LBound(results, 1) To UBound(results, 1)
            results(iRow, commodityColumn) = commodityGroups(results(iRow, 1))(1)
            results(iRow, subGroupColumn) = commodityGroups(results(iRow, 1))(2)
Next_Commodity_Assignment:
        Next iRow
        
        On Error GoTo 0
        
        Set LO = Available_Contracts.ListObjects("Contract_Availability")
        
        With LO
    
            With .DataBodyRange
                .SpecialCells(xlCellTypeConstants).ClearContents
                .Cells(1, 1).Resize(UBound(results, 1), UBound(results, 2)).Value2 = results
            End With
            
            .Resize .Range.Cells(1, 1).Resize(UBound(results, 1) + 1, .ListColumns.count)

        End With
        
        ClearRegionBeneathTable LO
        
    End If
    
    EnableApplicationProperties appProperties
    
    Exit Sub
    
Code_Not_Found:
     Resume Next_Commodity_Assignment
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

Function GetDataForMultipleContractsFromDatabase(reportType As String, getCombinedData As Boolean, sortDataAscending As Boolean, _
                        Optional maxWeeksInPast As Integer = -1, Optional alternateCodes As Variant, _
                        Optional includePriceColumn As Boolean = False) As Collection
'====================================================================================================================================
'   Summary: Retrieves data for all favorites or select contracts from the database and stores an array for each contract keyed to its contract code.
'   Inputs:
'       reportType: One of L,D or T to select which database to target.
'       getCombinedData: true if Futures + Options data is wanted; otherwise, false.
'       sortDataAscending: true to sort data in ascending order by date otherwise false for descending.
'       maxWeeksInPast: Number of weeks in the past in addition to the current week to query for.
'       alternateCodes: Specific contract codes to filter for from the database.
'       includePriceColumn: true if you want to return prices as well.
'   Returns: A collection of arrays keyed to that contracts contract code.
'====================================================================================================================================
    Dim SQL As String, tableName As String, adodbConnection As Object, record As Object, SQL2 As String, _
    favoritedContractCodes As String, queryResult() As Variant, fieldNames As String, _
    contractClctn As Collection, allContracts As New Collection, databaseExists As Boolean, oldestWantedDate As Date, mostRecentDate As Date
     
    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]", dateColumn As Byte = 1
    
    On Error GoTo Finally
    
    If IsMissing(alternateCodes) Then
        ' Get a list of all contract codes that have been favorited.
        queryResult = WorksheetFunction.Transpose(Variable_Sheet.ListObjects("Current_Favorites").DataBodyRange.columns(1).Value2)
    Else
        queryResult = alternateCodes
    End If
    
    favoritedContractCodes = Join(QuotedForm(queryResult, "'"), ",")
          
    Set adodbConnection = CreateObject("ADODB.Connection")
    Set record = CreateObject("ADODB.RecordSet")
    
    Call database_details(getCombinedData, reportType, adodbConnection, tableName, , databaseExists)
    
    If Not databaseExists Then Exit Function
    
    With adodbConnection
        .Open
        'Get a record of all field names in tha database.
        Set record = .Execute(CommandText:=tableName, Options:=adCmdTable)
    End With
    
    fieldNames = FilterColumnsAndDelimit(FieldsFromRecordSet(record, encloseFieldsInBrackets:=True), reportType, includePriceColumn:=includePriceColumn)    'Field names from database returned as an array
    record.Close
        
    mostRecentDate = Variable_Sheet.Range("Most_Recently_Queried_Date").Value2
    
    SQL2 = "SELECT " & codeField & " FROM " & tableName & " WHERE " & dateField & " = CDATE('" & Format(mostRecentDate, "yyyy-mm-dd") & "') AND " & codeField & " in (" & favoritedContractCodes & ");"
    
    oldestWantedDate = IIf(maxWeeksInPast > 0, DateAdd("ww", -maxWeeksInPast, mostRecentDate), DateSerial(1970, 1, 1))
    
    SQL = "SELECT " & fieldNames & " FROM " & tableName & _
    " WHERE " & codeField & " in (" & SQL2 & ") AND " & dateField & " >=CDATE('" & oldestWantedDate & "') Order BY " & codeField & " ASC," & dateField & " " & IIf(sortDataAscending, "ASC;", "DESC;")
    
    Erase queryResult
    
    With record
        .Open SQL, adodbConnection, adOpenStatic, adLockReadOnly, adCmdText
        queryResult = TransposeData(.GetRows)
        .Close
    End With
    
    adodbConnection.Close
    
    Dim codeColumn As Byte, nameColumn As Byte, iRow As Long, iColumn As Byte, _
    queryRow() As Variant, CC As Variant, output As New Collection
    
    codeColumn = UBound(queryResult, 2) - IIf(includePriceColumn, 1, 0)
    nameColumn = 2
    
    ReDim queryRow(1 To codeColumn + IIf(includePriceColumn, 1, 0))
    
    With allContracts
        'Group contracts into separate collections for further processing
        For iRow = LBound(queryResult, 1) To UBound(queryResult, 1)
        
            For iColumn = 1 To UBound(queryResult, 2)
                queryRow(iColumn) = IIf(IsNull(queryResult(iRow, iColumn)), Empty, queryResult(iRow, iColumn))
            Next iColumn
        
            On Error GoTo Catch_CollectionMissing
            Set contractClctn = .Item(queryRow(codeColumn))
            
            On Error GoTo Catch_DuplicateKeyAttempt
            ' Use dates as a key
            contractClctn.Add queryRow, CStr(queryRow(dateColumn))
Next_QueryRow_Iterator:
        Next iRow
        
        Erase queryResult

    End With
    
    On Error GoTo Finally
    
    With output
        For iRow = 1 To allContracts.count
            .Add CombineArraysInCollections(allContracts(iRow), Append_Type.Multiple_1d), allContracts(iRow)(1)(codeColumn)
        Next iRow
    End With
    
    Set GetDataForMultipleContractsFromDatabase = output
    
Finally:
    
    If Not record Is Nothing Then
        If record.State = adStateOpen Then record.Close
        Set record = Nothing
    End If
    
    If Not adodbConnection Is Nothing Then
        If adodbConnection.State = adStateOpen Then adodbConnection.Close
        Set adodbConnection = Nothing
    End If
    
    If Err.Number <> 0 Then
        DisplayErrorIfAvailable Err, "GetDataForMultipleContractsFromDatabase()"
        Err.Raise Err.Number, "GetDataForMultipleContractsFromDatabase()", Err.description
    End If
    
    Exit Function
    
Catch_CollectionMissing:

    Set contractClctn = New Collection
    allContracts.Add contractClctn, queryRow(codeColumn)
    
    Resume Next
Catch_DuplicateKeyAttempt:

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
    
    Set contractClctn = GetDataForMultipleContractsFromDatabase(ReportChr, Not queryFutOnly, True, threeYearsInWeeks + previousWeeksToCalculate + 2)
    
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
    
        .Range("A1").Value2 = Variable_Sheet.Range("Most_Recently_Queried_Date").Value2
        
        With .ListObjects("Dashboard_Results" & ReportChr)
            
            With .DataBodyRange
                .ClearContents
                With .Resize(UBound(output, 1), UBound(output, 2))
                    .Value2 = output
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
Public Function CftcOutputTable(report As String) As ListObject
'==================================================================================================
'   Returns the ListObject used to store data for the report abbreviated by the report paramater.
'   Paramater:
'       - report: One of L,D or T.
'==================================================================================================
    Set CftcOutputTable = Assign_Linked_Data_Sheet(report).ListObjects(report & "_Data")
End Function

Public Sub Save_For_Github()
'=======================================================================================================
' Toggles range value that marks the workboook for upload to github.
'=======================================================================================================
    If UUID Then
        Range("Github_Version").Value2 = True
        Custom_SaveAS
    End If

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
    Dim availableContractInfo As Collection, SQL As String, adodbConnection As Object, tableName As String, queryResult() As Variant, CC As Integer

    Const dateField As String = "[Report_Date_as_YYYY-MM-DD]", _
          codeField As String = "[CFTC_Contract_Market_Code]", _
          nameField As String = "[Market_and_Exchange_Names]"
    
    Dim codeColumn As Byte, rowIndex As Long, ColumnIndex As Byte, recordsWithSameContractCode As Collection, _
    queryRow() As Variant, recordsByDateByCode As New Collection, minDate As String
    
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
    
    Set availableContractInfo = GetAvailableContractInfo
    
    ReDim queryRow(1 To UBound(queryResult, 2))
    
    With recordsByDateByCode
        'Group contracts into separate collections for further processing
        For rowIndex = LBound(queryResult, 1) To UBound(queryResult, 1)
        
            For ColumnIndex = 1 To UBound(queryResult, 2)
                queryRow(ColumnIndex) = queryResult(rowIndex, ColumnIndex)
            Next ColumnIndex
        
            On Error GoTo Create_Contract_Collection
            Set recordsWithSameContractCode = .Item(queryRow(codeColumn))
            
            On Error GoTo 0
            'Use dates as a key
            recordsWithSameContractCode.Add queryRow, CStr(queryRow(1))
            
        Next rowIndex
        
        Erase queryResult
        Erase queryRow
        
    End With
    
    With recordsByDateByCode
    
        For CC = .count To 1 Step -1
            
            Set recordsWithSameContractCode = .Item(CC)
            
            queryResult = CombineArraysInCollections(recordsWithSameContractCode, Append_Type.Multiple_1d)
            
            .Remove queryResult(1, codeColumn)
            
            If HasKey(availableContractInfo, CStr(queryResult(1, codeColumn))) Then
            
                If TryGetPriceData(queryResult, 3, availableContractInfo(queryResult(1, codeColumn)), True, True) Then
                    .Add queryResult, queryResult(1, codeColumn)
                End If
            
            End If
        
        Next CC
    
    End With
    
    queryResult = CombineArraysInCollections(recordsByDateByCode, Append_Type.Multiple_2d)
    
    On Error GoTo 0
    
    UpdateDatabasePrices queryResult, "L", True, 3
    
    overwrite_with_legacy_combined_prices minimum_date:=CDate(minDate)
    
    Exit Sub
    
Create_Contract_Collection:

    Set recordsWithSameContractCode = New Collection
    recordsByDateByCode.Add recordsWithSameContractCode, queryRow(codeColumn)
    
    Resume Next

End Sub
Sub FindDatabasePathInSameFolder()
'===========================================================================================================
' Looks for MS Access Database files that haven't been renamed within the same folder as the Excel workbook.
'===========================================================================================================
    Dim legacy As New LoadedData, TFF As New LoadedData, DGG As New LoadedData, _
    strfile As String, foundCount As Byte, folderPath As String
    
    On Error GoTo Prompt_User_About_UserForm
    
    legacy.InitializeClass "L"
    DGG.InitializeClass "D"
    TFF.InitializeClass "T"
    
    folderPath = ThisWorkbook.Path & Application.PathSeparator
    ' Filter for Microsoft Access databases.
    strfile = Dir(folderPath & "*.accdb")
    
    Do While LenB(strfile) > 0
        
        If LCase$(strfile) Like "*disaggregated.accdb" And IsEmpty(DGG.CurrentDatabasePath) Then
            DGG.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        ElseIf LCase$(strfile) Like "*legacy.accdb" And IsEmpty(legacy.CurrentDatabasePath) Then
            legacy.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        ElseIf LCase$(strfile) Like "*tff.accdb" And IsEmpty(TFF.CurrentDatabasePath) Then
            TFF.CurrentDatabasePath = folderPath & strfile
            foundCount = foundCount + 1
        End If
        
        strfile = Dir
    Loop
    
Prompt_User_About_UserForm:

    If foundCount <> 3 And Not UUID Then
        MsgBox "Database paths couldn't be auto-retrieved." & vbNewLine & vbNewLine & _
        "Please use the Database Paths USerform to fill in the needed data."
        
        Err.Raise 17, "FindDatabasePathInSameFolder", "Missing Database(s)"
    End If
    
End Sub
Public Function GetStoredReportDetails(reportType As String) As LoadedData
    
    Dim storedData As New LoadedData
    storedData.InitializeClass reportType
    Set GetStoredReportDetails = storedData
    
End Function
Private Sub AttemptCross()

    Dim tableLayout As Collection, notionalValue() As Double, iRow As Integer, _
    dataFromDatabase As Collection, reportType As String, notionalValuesByCode As New Collection, _
    Code As Variant, contractUnits As Variant, prices As Variant
    
    On Error GoTo Finally
    ' Setting equal to -1 will allow all data to be retrieved.
    Const maxWeeksInPast As Integer = -1
    
    IncreasePerformance
    
    Dim codeToLong As String, codeToShort As String, selectionTable As Variant
    
    With ForexCross
    
        selectionTable = .ListObjects("Long_Short").DataBodyRange.Value2
        
        With .ListObjects("ForexTickers")
            codeToLong = WorksheetFunction.VLookup(selectionTable(1, 1), .DataBodyRange, 2, False)
            codeToShort = WorksheetFunction.VLookup(selectionTable(1, 2), .DataBodyRange, 2, False)
        End With
        
        If LenB(codeToLong) = 0 Or LenB(codeToShort) = 0 Or codeToLong = codeToShort Then
            MsgBox "Invalid input paramaters."
            Exit Sub
        End If
        
        selectionTable = Array(codeToLong, codeToShort)
        
    End With
    
    reportType = "L"
    
    Set dataFromDatabase = GetDataForMultipleContractsFromDatabase(reportType, True, True, maxWeeksInPast, selectionTable, True)

    Set tableLayout = GetExpectedLocalFieldInfo(reportType, True, True)

    For Each Code In selectionTable

        contractUnits = WorksheetFunction.index(dataFromDatabase(Code), 0, tableLayout("contract_units").ColumnIndex)  '
        contractUnits = GetNumbers(contractUnits)

        With notionalValuesByCode

            .Add New Collection, Code

            With .Item(Code)

                prices = Application.index(dataFromDatabase(Code), 0, tableLayout(tableLayout.count).ColumnIndex + 1)

                ReDim notionalValue(LBound(contractUnits, 1) To UBound(contractUnits, 1))

                For iRow = LBound(contractUnits, 1) To UBound(contractUnits, 1)
                    If Not IsEmpty(prices(iRow, 1)) Then
                        notionalValue(iRow) = prices(iRow, 1) * contractUnits(iRow, 1)
                    End If
                Next
                
                .Add notionalValue, "Notional"

            End With

        End With

    Next
    
    'Calculate hedge ratio and combine.
    
    Dim contractToLong As Variant, contractToShort As Variant, iColumn As Byte, _
    hedgeRatio As Double, output() As Variant, nonCommLong As Byte, commLong As Byte, commShort As Byte, _
    nonCommShort As Byte, iShortRow As Integer, iReduction As Integer
    
    contractToLong = dataFromDatabase(codeToLong)
    contractToShort = dataFromDatabase(codeToShort)
    
'    commLong = tableLayout("comm_positions_long_all").ColumnIndex
'    commShort = tableLayout("comm_positions_short_all").ColumnIndex
'    nonCommLong = tableLayout("noncomm_positions_long_all").ColumnIndex
'    nonCommShort = tableLayout("noncomm_positions_short_all").ColumnIndex

    commLong = tableLayout("pct_of_oi_comm_long_all").ColumnIndex
    commShort = tableLayout("pct_of_oi_comm_short_all").ColumnIndex
    
    nonCommLong = tableLayout("pct_of_oi_noncomm_long_all").ColumnIndex
    nonCommShort = tableLayout("pct_of_oi_noncomm_short_all").ColumnIndex
    
    ReDim output(1 To UBound(contractToLong, 1), 1 To 5)
    
    iShortRow = UBound(contractToShort, 1)
    
    On Error GoTo Exit_Loop
    
    For iRow = UBound(contractToLong, 1) To LBound(contractToLong, 1) Step -1
                    
        If contractToLong(iRow, 1) = contractToShort(iShortRow, 1) Then
        
            hedgeRatio = notionalValuesByCode(codeToLong)("Notional")(iRow) / notionalValuesByCode(codeToShort)("Notional")(iShortRow)
            'hedgeRatio = notionalValuesByCode(codeToShort)("Notional")(iShortRow) / notionalValuesByCode(codeToLong)("Notional")(iRow)
            
'            For iColumn = LBound(output, 2) + 2 To UBound(output, 2)
'
'                If InStrB(1, tableLayout(iColumn).EditedName, "spread") = 0 Then
'
'                    If InStrB(1, tableLayout(iColumn).EditedName, "comm") = 1 Then
'                        nonCommLong = tableLayout("comm_positions_long_all").ColumnIndex
'                        nonCommShort = tableLayout("comm_positions_short_all").ColumnIndex
'                    Else
'                        nonCommLong = tableLayout("noncomm_positions_long_all").ColumnIndex
'                        nonCommShort = tableLayout("noncomm_positions_short_all").ColumnIndex
'                    End If
'
'                End If
'
'                output(iRow, 2) = CLng((contractToLong(iRow, nonCommLong) * hedgeRatio) + contractToShort(iShortRow, nonCommShort))
'                output(iRow, 3) = CLng((contractToLong(iRow, nonCommShort) * hedgeRatio) + contractToShort(iShortRow, nonCommLong))
'            Next


'            output(iRow, 2) = CLng((contractToLong(iRow, nonCommLong) * hedgeRatio) + contractToShort(iShortRow, nonCommShort))
'            output(iRow, 3) = CLng((contractToLong(iRow, nonCommShort) * hedgeRatio) + contractToShort(iShortRow, nonCommLong))
'            output(iRow, 4) = CLng((contractToLong(iRow, commLong) * hedgeRatio) + contractToShort(iShortRow, commShort))
'            output(iRow, 5) = CLng((contractToLong(iRow, commShort) * hedgeRatio) + contractToShort(iShortRow, commLong))
            
            
            output(iRow, 2) = contractToLong(iRow, nonCommLong) - contractToLong(iRow, nonCommShort)
            output(iRow, 3) = contractToShort(iShortRow, nonCommLong) - contractToShort(iShortRow, nonCommShort)
            
            
            output(iRow, 4) = contractToLong(iRow, commLong) - contractToLong(iRow, commShort)
            output(iRow, 5) = contractToShort(iShortRow, commLong) - contractToShort(iShortRow, commShort)
            
            output(iRow, 1) = contractToLong(iRow, 1)
            
        End If
        
        iShortRow = iShortRow - 1
        
    Next iRow
    
PlaceOnSheet:
    
    On Error GoTo 0
    
    Dim bb As Variant, LO As ListObject
    
    Set LO = ForexCross.ListObjects("CrossTable")
    
    With LO.DataBodyRange
        
        ChangeFilters LO, bb
        .Range(.Cells(1, 1), .Cells(.Rows.count, UBound(output, 2))).ClearContents
        .Resize(UBound(output, 1), UBound(output, 2)).Value2 = Reverse_2D_Array(output)
        ResizeTableBasedOnColumn .ListObject, .columns(1)
        ClearRegionBeneathTable .ListObject
        RestoreFilters LO, bb
        
    End With
    
Finally:
    DisplayErrorIfAvailable Err, "AttemptCross()"
    Re_Enable
    
    Exit Sub
    
Exit_Loop:
    Resume PlaceOnSheet
End Sub

Public Function GetContractInfo_DbVersion() As Collection
'==============================================================================================
' Creates a collection of Contract objects keyed to their contract code for each
' available contract within the database.
'==============================================================================================

    Dim Available_Data() As Variant, CD As ContractInfo, T As Integer, _
    pAllContracts As New Collection, priceSymbol As String, usingYahoo As Boolean, symbolsRange As Range

    Available_Data = Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.value
    
    Const codeColumn As Byte = 1, nameColumn As Byte = 2, availabileColumn As Byte = 3, _
    commodityGroupColumn As Byte = 4, subGroupColumn As Byte = 5, hasSymbolColumn As Byte = 6, isFavoriteColumn As Byte = 7
    
    Set symbolsRange = Symbols.ListObjects("Symbols_TBL").DataBodyRange
    
    For T = LBound(Available_Data) To UBound(Available_Data)
        
        priceSymbol = vbNullString
        usingYahoo = False
        
        If Available_Data(T, hasSymbolColumn) = True Then
            priceSymbol = WorksheetFunction.VLookup(Available_Data(T, codeColumn), symbolsRange, 3, False)
            usingYahoo = LenB(priceSymbol) > 0
        End If
        
        Set CD = New ContractInfo
        
        With CD
            
            .InitializeBasicVersion CStr(Available_Data(T, codeColumn)), CStr(Available_Data(T, nameColumn)), CStr(Available_Data(T, availabileColumn)), CBool(Available_Data(T, isFavoriteColumn)), priceSymbol, usingYahoo
            
            On Error GoTo Possible_Duplicate_Key
                pAllContracts.Add CD, .ContractCode
            On Error GoTo 0
            
       End With

    Next T
    
    Set GetContractInfo_DbVersion = pAllContracts

    Exit Function
    
Possible_Duplicate_Key:
    Resume Next

End Function

Public Sub DeactivateContractSelection()

    If IsLoadedUserform("Contract_Selection") Then
       Unload Contract_Selection
    End If
    
End Sub

Public Sub Open_Contract_Selection()
    Dim reportToLoad As String
            
    On Error GoTo Failed_To_Get_Type
        With ThisWorkbook
            reportToLoad = .Worksheets(.ActiveSheet.name).WorksheetReportType
        End With
    On Error GoTo 0
    
    With Contract_Selection
        .SetReport reportToLoad
        .Show
    End With
Finally:
    Exit Sub
    
Failed_To_Get_Type:
    MsgBox ThisWorkbook.ActiveSheet.name & " does not have a publicly available WorksheetReportType Function."
    Resume Finally
End Sub

Public Sub Show_Client_Differences()
    With ClientAvn
        .Visible = xlSheetVisible
        .Activate
    End With
    ReversalCharts.Visible = xlSheetVisible
End Sub
Private Sub CFTC_CalculateWeeklyChanges()

    Dim contractData As Variant, outputA() As Variant, contractDataByCode As Collection, iRow As Integer, mostRecentContractCodes As Variant
    
    Dim localFields As Collection, availableContracts As Collection, currentWeekNet As Long, previousWeekNet As Long
    Const maxWeeksInPast As Byte = 1
    
    mostRecentContractCodes = Application.Transpose(Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.columns(1).Value2)
    
    Set contractDataByCode = GetDataForMultipleContractsFromDatabase("L", True, False, maxWeeksInPast, mostRecentContractCodes)
    
    Set localFields = GetExpectedLocalFieldInfo("L", True, True)
    Set availableContracts = GetAvailableContractInfo
    
    Dim commLong As Byte, commShort As Byte, nonCommLong As Byte, nonCommShort As Byte, codeColumn As Byte, _
    iColumn As Byte, columnLong As Byte, columnShort As Byte, oiColumn As Byte
    
    Const currentWeek As Byte = 1, previousWeek As Byte = 2
    
    commLong = localFields("comm_positions_long_all").ColumnIndex
    commShort = localFields("comm_positions_short_all").ColumnIndex
    nonCommLong = localFields("noncomm_positions_long_all").ColumnIndex
    nonCommShort = localFields("noncomm_positions_short_all").ColumnIndex
    codeColumn = localFields("cftc_contract_market_code").ColumnIndex
    oiColumn = localFields("oi_all").ColumnIndex
    
    ReDim outputA(1 To contractDataByCode.count, 1 To 8)
    
    For Each contractData In contractDataByCode
    
        On Error GoTo Catch_CodeMissing
        If UBound(contractData, 1) = 2 Then
            
            iRow = iRow + 1
            On Error GoTo 0
            outputA(iRow, 1) = contractData(currentWeek, codeColumn)
            outputA(iRow, 2) = availableContracts(contractData(currentWeek, codeColumn)).ContractNameWithoutMarket
            
            For iColumn = 0 To 3
            
                Dim columnTarget As Variant
                
                columnTarget = Array(nonCommLong, nonCommShort, commLong, commShort)(iColumn)
                
                If contractData(previousWeek, columnTarget) <> 0 Then
                    outputA(iRow, 3 + iColumn) = (contractData(currentWeek, columnTarget) - contractData(previousWeek, columnTarget)) / contractData(previousWeek, columnTarget)
                End If
                
                If iColumn Mod 2 = 0 Then
                    
                    currentWeekNet = contractData(currentWeek, columnTarget) - contractData(currentWeek, columnTarget + 1)
                    previousWeekNet = contractData(previousWeek, columnTarget) - contractData(previousWeek, columnTarget + 1)
                    
                    Dim calc As Byte
                    
                    calc = 0
                    
                    If columnTarget = nonCommLong Then
                        calc = 7
                    ElseIf columnTarget = commLong Then
                        calc = 8
                    End If
                    
                    If calc <> 0 Then
                    
                        On Error Resume Next
                        'Commercial net change from previouse / (previous longs+shorts)
                        outputA(iRow, calc) = (currentWeekNet - previousWeekNet) / (contractData(previousWeek, columnTarget) + contractData(previousWeek, columnTarget + 1))
                                                            
                        '% difference in net position.
                        'outputA(iRow, calc) = (currentWeekNet - previousWeekNet) / contractData(previousWeek, 3)
                        
                        If Err.Number <> 0 Then
                            outputA(iRow, calc) = 0
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                    End If
                    
                End If
                
                '=SUM( IF( TRUE , IF(K36<[Column2],1,FALSE),FALSE ))
                
                'columnLong = IIf(iColumn = 3, nonCommLong, commLong)
                'columnShort = IIf(iColumn = 5, nonCommShort, commShort)
                
                'currentWeekNet = contractData(1, columnLong) - contractData(1, columnShort)
                'previousWeekNet = contractData(2, columnLong) - contractData(2, columnShort)
                
                'If currentWeekNet > previousWeekNet Then
                '    outputA(iRow, iColumn) = (currentWeekNet - previousWeekNet) / previousWeekNet
                'Else
                '    outputA(iRow, iColumn) = (previousWeekNet - currentWeekNet) / Abs(previousWeekNet)
                'End If
                
                'outputA(iRow, 3 + iColumn) = Abs((currentWeekNet - previousWeekNet)) / ((Abs(previousWeekNet) + Abs(currentWeekNet)) / 2)
                
            Next iColumn
            
        End If
        
Next_Array:
    Next
    
    Dim tableDataRng As Range, LO As ListObject, currentFilters As Variant, appProperties As Collection
    
    Set LO = WeeklyChanges.ListObjects("PctNetChange")
    Set tableDataRng = LO.DataBodyRange
    
    With tableDataRng
    
        Set appProperties = DisableApplicationProperties(True, False, True)
        
        ChangeFilters LO, currentFilters
        
        .SpecialCells(xlCellTypeConstants).ClearContents
        .columns(4).Resize(UBound(outputA, 1), UBound(outputA, 2)).Value2 = outputA
        
        ResizeTableBasedOnColumn LO, .columns(4)
        
        ClearRegionBeneathTable LO
        With LO.Sort
            With .SortFields
                .Clear
                'Group
                .Add tableDataRng.columns(2), xlSortOnValues, xlAscending
                'SubGroup
                .Add tableDataRng.columns(3), xlSortOnValues, xlAscending
                'Name
                '.Add tableDataRng.columns(5), xlSortOnValues, xlAscending
                'Rank
                .Add tableDataRng.columns(11), xlSortOnValues, xlDescending
            End With
            .Apply
        End With
        RestoreFilters LO, currentFilters
        
        WeeklyChanges.Range("reflectedDate").Value2 = Variable_Sheet.Range("Most_Recently_Queried_Date").Value2
        EnableApplicationProperties appProperties
        
        '=SUM(IF(SUBTOTAL(103,OFFSET([Commercial Net change/Total Position],ROW([Commercial Net change/Total Position])-ROW($A$3),0,1))>0,IF(K10<[Commercial Net change/Total Position],1)))+1
    End With
        
    Exit Sub

Catch_CodeMissing:
    Resume Next_Array
End Sub

