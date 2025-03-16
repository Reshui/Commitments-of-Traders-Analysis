Attribute VB_Name = "Database_Interactions"

#If Not Mac Then
    Private AfterEventHolder As ClassQTE
    Private SQLSERVER_DatabaseExists_COT As Boolean
    Private SQL_Server_TableExistance As Object
    Private Const SqlServerDatabaseName As String = "Commitments_Of_Traders_MoshiM"
    Private Const PriceDataTableName As String = "PriceData"

    Public Enum DbError

        VersionUnacceptable = vbObjectError + 600
        UserSelectedFieldsEqualsZero = vbObjectError + 601
        UseMasterCatalog = vbObjectError + 602
        UseDisaggregatedReport = vbObjectError + 603
        UnfinishedCode = vbObjectError + 604
        NoMatchingFields = vbObjectError + 605
        DatabaseConnectionFailed = vbObjectError + 606
        ServerNameMissing = vbObjectError + 607
        InvalidFunctionParameter = vbObjectError + 608
        ExcelTableMissing = vbObjectError + 609
        DatabaseTableMissing = vbObjectError + 610
        NotImplemented = vbObjectError + 611
        
        ' Specific to SQL Server -----------------------------------------------------------------------------------------
        PrimaryKeyViolation_SQLSERVER = -2147217873
        InvalidCast = -2147217887
        ' SQL Server may be unavailable Control Panel > Administrative Tools > Services.
        InvalidConnectionStringSqlServer = -2147467259
        '-----------------------------------------------------------------------------------------------------------------
        ' https://learn.microsoft.com/en-us/sql/ado/guide/data/provider-errors?view=sql-server-ver16
        ProviderError = -2147467259

        'https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/errorvalueenum
        InvalidParameterAssignment = adErrDataConversion '3421

        NoRecords = 3021

    End Enum

    Private StoredAdoObjects As New Scripting.Dictionary

    Option Explicit
    Private Function TryGetDatabaseDetails(openInterestSelection As OpenInterestEnum, eReport As ReportEnum, _
        Optional ByRef databaseConnection As ADODB.Connection, Optional ByRef tableNameToReturn$, Optional ByRef databasePath$, _
        Optional ByRef suppressMsgBoxIfUnavailable As Boolean = False, _
        Optional ignoreSqlServerDetails As Boolean = False, _
        Optional ByRef isSqlServerDetail As Boolean, _
        Optional ByRef doesPriceTableExist As Boolean) As Boolean
    '===================================================================================================================
    'Summary: Determines if database exists. If it does the appropriate variables or properties are assigned values if needed.
    'Inputs:
    '        eReport - ReportEnum used to select a database.
    '        openInterestSelection - Enum used to select a specific table within the database.
    '        databaseConnection - If supplied then a connection string will be assigned to this object.
    '        tableNameToReturn - If supplied then the wanted table within the selected database will be returned to this variable.
    '        databasePath - If supplied then the path to the database will be stored in this variable.
    '        ignoreSqlServerDetails - If true then attempts to retrieve SQL Server details will be denied.
    '        isSqlServerDetail - Returned as TRUE if SQL SERVER connection.
    'Returns: True if a database exists for the given inputs; othewise, false.
    '===================================================================================================================
        Dim isDatabaseAvailable As Boolean, setConnectionToNothingOnExit As Boolean, sqlServerConnectionSuccess As Boolean

        If databaseConnection Is Nothing Then
            Set databaseConnection = GetStoredAdoClass(eReport).Connection
            setConnectionToNothingOnExit = True
        End If
        
        If LenB(tableNameToReturn) <> 0 Then tableNameToReturn = vbNullString
        
        If Not openInterestSelection = OpenInterestEnum.OptionsOnly Then
            On Error GoTo Finally

            If Not ignoreSqlServerDetails And DoesUserPermit_SqlServer() Then
                
                If Not databaseConnection Is Nothing Then sqlServerConnectionSuccess = IsSqlServerConnection(databaseConnection)
                
                If databaseConnection.State = adStateClosed Or Not sqlServerConnectionSuccess Then
                   sqlServerConnectionSuccess = TryConnectingToSqlServer(databaseConnection Is Nothing, databaseConnection, True, eReport, openInterestSelection, tableNameToReturn)
                End If
                
                If sqlServerConnectionSuccess Then
                    TryGetDatabaseDetails = True: isSqlServerDetail = True: doesPriceTableExist = True
                    If LenB(tableNameToReturn) = 0 Then tableNameToReturn = GetSqlServerTableName(eReport, openInterestSelection)
                End If
            Else
                Dim tableNamePrefix$
                ' Attempt Microsoft Access connection.
                With GetStoredReportDetails(eReport)
                    If eReport = eTFF Then
                        tableNamePrefix = "TFF"
                    Else
                        tableNamePrefix = .FullReportName.Value2
                    End If
                    databasePath = .CurrentDatabasePath.Value2
                End With
                
                isDatabaseAvailable = FileOrFolderExists(databasePath) And databasePath Like "*.accdb"

                If Not isDatabaseAvailable And Not suppressMsgBoxIfUnavailable Then
                    MsgBox ConvertReportTypeEnum(eReport) & " ... Unable to locate Microsoft Access database."
                ElseIf isDatabaseAvailable Then
                    With databaseConnection
                        If .State <> adStateOpen Then .Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath & ";"
                    End With
                    doesPriceTableExist = DoesTableExist(databaseConnection, PriceDataTableName)
                End If

                If doesPriceTableExist Then
                    tableNameToReturn = GetSqlServerTableName(eReport, openInterestSelection)
                Else
                    tableNameToReturn = tableNamePrefix & IIf(openInterestSelection = OpenInterestEnum.FuturesAndOptions, "_Combined", "_Futures_Only")
                End If

                TryGetDatabaseDetails = isDatabaseAvailable
            End If
        Else
            Err.Raise DbError.InvalidFunctionParameter, Description:="OpenInterestEnum.OptionsOnly is an invalid value for the parameter [openInterestSelection]."
        End If
Finally:
        If Not databaseConnection Is Nothing And (Err.Number <> 0 Or setConnectionToNothingOnExit) Then
'            With databaseConnection
'                If .State = adStateOpen Then .Close
'            End With
            If setConnectionToNothingOnExit Then Set databaseConnection = Nothing
        End If
        If Err.Number <> 0 Then PropagateError Err, "TryGetDatabaseDetails"

    End Function
    Public Sub DisposeConnections()
Attribute DisposeConnections.VB_Description = "Disposes of any active connections to the Commitments of Traders database."
Attribute DisposeConnections.VB_ProcData.VB_Invoke_Func = " \n14"
    '===================================================================================================================
    'Summary: Disposes of all AdoContainer objects stored within the public collection storedAdoObjects.
    '===================================================================================================================
        Dim AdoContainer As Variant
        
        For Each AdoContainer In StoredAdoObjects.Items
            Call AdoContainer.CloseConnection
        Next AdoContainer
        Set StoredAdoObjects = Nothing
        
    End Sub
    Private Function FilterDatabaseFieldsWithLocalInfo(record As ADODB.recordSet, fieldInfoByEditedName As Collection) As Collection
    '===================================================================================================================
    'Summary: Generates FieldInfo instances for fields contained within [record]
    'Inputs:
    '   record : ADODB.Record that contains all fields for a table within a database.
    '   fieldInfoByEditedName :
    'Returns: A collection of FieldInfo instances generated from [record] keyed to a standardized name.
    '===================================================================================================================
        Dim item As Field, output As New Collection, FI As FieldInfo

        On Error GoTo Propagate
        For Each item In record.Fields

            On Error GoTo Catch_MissingKey
            Set FI = fieldInfoByEditedName(StandardizedDatabaseFieldNames(item.name))

            With FI
                If Not (.IsMissing Or .EditedName = "id") Then
                    Call .EditDatabaseName(item.name)
                    output.Add FI, .EditedName
                End If
            End With
AttemptNextField:
        Next item

        On Error GoTo 0
        If output.Count = 0 Then Err.Raise DbError.NoMatchingFields, "FilterDatabaseFieldsWithLocalInfo", "No matching field names between local database and supplied FieldIno collection."

        Set FilterDatabaseFieldsWithLocalInfo = output
        Exit Function
Propagate:
    PropagateError Err, "FilterDatabaseFieldsWithLocalInfo"
Catch_MissingKey:
        Resume AttemptNextField
    End Function
    Private Function GetFieldNamesFromRecord(record As ADODB.recordSet, encloseFieldsInBrackets As Boolean) As String()
    '===================================================================================================================
    'record is a ADODB.Record object containing a single row of data from which field names are retrieved,formatted and output as an array
    '===================================================================================================================
        Dim Z As Long, fieldNamesInRecord$(), databaseField As Field

        On Error GoTo Propagate

        With record
            ReDim fieldNamesInRecord(0 To .Fields.Count - 1)
            For Each databaseField In .Fields
                With databaseField
                    If Not .name = "ID" Then
                        If encloseFieldsInBrackets Then
                            fieldNamesInRecord(Z) = "[" & .name & "]"
                        Else
                            fieldNamesInRecord(Z) = .name
                        End If
                        Z = Z + 1
                    End If
                End With
            Next databaseField
        End With
        ReDim Preserve fieldNamesInRecord(LBound(fieldNamesInRecord) To Z - 1)
        GetFieldNamesFromRecord = fieldNamesInRecord
        Exit Function
Propagate:
        PropagateError Err, "GetFieldNamesFromRecord"
    End Function
    Private Function FilterCollectionOnFieldInfoKey(databaseFields As Collection, localFieldInfo As Collection) As Collection
    '====================================================================================================================================
    '   Summary: Filters [databaseFields] based on FieldInfo found in [localFieldInfo]
    '   Inputs:
    '       databaseFields: FieldInfo collection generated from a database query.
    '       localFieldInfo: FieldInfo collection generated from local storage.
    '   Returns: A filtered collection.
    '====================================================================================================================================
        Dim cc As New Collection, FI As FieldInfo, editedFieldName$, i&

        On Error GoTo NEXT_FIELD
        i = 0
        For Each FI In localFieldInfo
            editedFieldName = FI.EditedName
            With databaseFields
                If TypeName(.item(editedFieldName)) = "FieldInfo" Then
                    i = i + 1
                    .item(editedFieldName).ColumnIndex = i
                End If
                cc.Add .item(editedFieldName), editedFieldName
            End With
NEXT_FIELD:
            If Err.Number <> 0 Then On Error GoTo -1
        Next FI

        Set FilterCollectionOnFieldInfoKey = cc

    End Function

    Private Function GetTableFieldsRecordSet(databaseConnection As ADODB.Connection, tableName$) As Object
    '====================================================================================================================================
    '   Summary: Queries [tableName] within the database connected by [databaseConnection] for a record of all fields contained within.
    '   Inputs:
    '       databaseConnection: ADODB.Connection for a database.
    '       tableName: Table name to get fields for.
    '   Returns: A record of fields within a table.
    '====================================================================================================================================
        Dim record As New ADODB.recordSet
        On Error GoTo Propagate

        'Set record = CreateObject("ADODB.RecordSet")
        record.Open tableName, databaseConnection, adOpenForwardOnly, adLockReadOnly, adCmdTable

        Set GetTableFieldsRecordSet = record
        Exit Function
Propagate:
        PropagateError Err, "GetTableFieldsRecordset"
    End Function
    Private Function GetStoredAdoClass(eReport As ReportEnum) As AdoContainer
    '====================================================================================================================================
    'Summary: Returns a AdoContainer object for the selected {eReport{ enum.
    'Inputs:
    '    eReport: Used to select which AdoContainer class to return.
    'Returns: An AdoContainer class.
    '====================================================================================================================================
        Dim value As AdoContainer, firstAvailableConnection As ADODB.Connection
        
        On Error GoTo Propagate
        
        With StoredAdoObjects
            If Not .Exists(eReport) Then .Add eReport, New AdoContainer
            Set firstAvailableConnection = .Items(0).Connection
            Set value = .item(eReport)
            
            With value
                ' Share the same connection for SQL Server connections otherwise create a new connection if one isn't available.
                If .Connection Is Nothing Then
                    Set .Connection = IIf(IsSqlServerConnection(firstAvailableConnection), firstAvailableConnection, New ADODB.Connection)
                End If
            End With
        End With
        Set GetStoredAdoClass = value
        Exit Function
Propagate:
        PropagateError Err, "GetStoredAdoClass"
    End Function
    Private Function QueryDatabaseForContract(eReport As ReportEnum, ByVal wantedReportEnum As OpenInterestEnum, wantedContractCode$, _
        Optional sortOrder As XlSortOrder = xlAscending, _
        Optional profiler As TimedTask, _
        Optional ByVal combinePositionsWithSpread As Boolean = False, _
        Optional ByVal disableNegativePositions As Boolean = False) As Variant()
    '====================================================================================================================================
    'Summary: Queries data within a database for output to a worksheet.
    'Inputs:
    '    eReport: Selects which database to query for [wantedContractCode].
    '    wantedReportEnum: OpenInterestEnum to query for.
    '    wantedContractCode: Contract code to query for.
    '    sortOrder:  Order returned data should be sorted in by date.
    '    disableNegativePositions: If True and [wantedReportEnum] = Options Only then negative position calculations will be set to 0,
    '       abs() added to the opposite position and removed from spread.
    '   combinePositionsWithSpread: If true and [wantedReportEnum] = Options Only then spread positions will be added back into gross positions.
    'Returns: An array of wanted data.
    '====================================================================================================================================
        Dim databaseConnection As ADODB.Connection, retrievalCommand As ADODB.command, tableNameWithinDatabase$, filteredDatabaseInfoByEditedName As Collection

        Dim sql$, delimitedWantedColumns$, futuresOnlyTableName$, retainedADO As AdoContainer, wantedFieldNamesSQL As Collection, optionsOnlyFields$(), _
        iCount As Long, sqlFieldName$, createCommand As Boolean, priceTableSqlJoin As String, disableSqlServerConnection As Boolean

        Dim currentFieldEdited$, groupedTraderData As Collection, traderGroupName$, wantedField As FieldInfo, swappedToFuturesAndOptions As Boolean, _
        connectedToSqlServer As Boolean, priceTableAvailable As Boolean, commandKey$

        Const FutOnly$ = "FutOnly", FutOpt$ = "FutOpt", ContractCodeFieldName$ = "cftc_contract_market_code", DateFieldName$ = "report_date_as_yyyy_mm_dd"

        On Error GoTo Finally
        
        Set retainedADO = GetStoredAdoClass(eReport)
        Set databaseConnection = retainedADO.Connection
        
        If TryGetDatabaseDetails(IIf(wantedReportEnum = OpenInterestEnum.OptionsOnly, OpenInterestEnum.FuturesAndOptions, wantedReportEnum), eReport, databaseConnection, tableNameWithinDatabase, ignoreSqlServerDetails:=disableSqlServerConnection, isSqlServerDetail:=connectedToSqlServer, doesPriceTableExist:=priceTableAvailable) Then

            With databaseConnection
                If .State = adStateClosed Then .Open
            End With
            
            If wantedReportEnum = OptionsOnly Then
                If combinePositionsWithSpread And disableNegativePositions Then
                    disableNegativePositions = False
                End If
            End If
            
            commandKey = wantedReportEnum & IIf(wantedReportEnum = OptionsOnly, ";Combined Spread=" & combinePositionsWithSpread & ";NegativPositionsEnabled=" & Not disableNegativePositions, vbNullString)
            
            With retainedADO
                Set retrievalCommand = .GetCommand(commandKey)
                If retrievalCommand Is Nothing Then
                    Set retrievalCommand = New ADODB.command
                    .SetCommand retrievalCommand, commandKey
                    createCommand = True
                End If
            End With
            ' If wantedReportEnum = OpenInterestEnum.OptionsOnly then column names need to be parsed and used to calculate after query return.
            If createCommand Or wantedReportEnum = OpenInterestEnum.OptionsOnly Then
                ' wantedReportEnum = OpenInterestEnum.OptionsOnly then the sql statement needs to be generated to get a collection
                ' of columns that need to calculated manually after the database returns data.
                Const GetFieldsKey$ = "Get wanted database fields"
                
                If Not profiler Is Nothing Then profiler.StartSubTask GetFieldsKey
                
                Set filteredDatabaseInfoByEditedName = FilterCollectionOnFieldInfoKey(GetFieldInfoForDatabaseTable(databaseConnection, tableNameWithinDatabase), GetExpectedLocalFieldInfo(eReport, filterUnwantedFields:=True, reArrangeToReflectSheet:=True, includePrice:=True, adjustIndexes:=True))
                
                If Not profiler Is Nothing Then profiler.StopSubTask GetFieldsKey

                If filteredDatabaseInfoByEditedName.Count = 0 Then
                    Err.Raise DbError.UserSelectedFieldsEqualsZero, "No wanted fields have been selected."
                End If

                Set wantedFieldNamesSQL = New Collection

                With wantedFieldNamesSQL
                    ' Store field names for use in SQL query.
                    For Each wantedField In filteredDatabaseInfoByEditedName
                        sqlFieldName = wantedField.EditedName
                        If priceTableAvailable Then
                            .Add sqlFieldName, sqlFieldName
                        Else
                            .Add wantedField.DatabaseNameForSQL, sqlFieldName
                        End If
                    Next wantedField
                End With

                If priceTableAvailable And createCommand Then
                    ' Price data is held in a separate table and needs to be joined with query.
                    priceTableSqlJoin = "LEFT JOIN " & PriceDataTableName & " as P on P.report_date_as_yyyy_mm_dd = T.report_date_as_yyyy_mm_dd AND P.cftc_contract_market_code = T.cftc_contract_market_code"
                End If

            End If

Create_SQL_Statement:
            If wantedReportEnum <> OpenInterestEnum.OptionsOnly Then

                If createCommand Then
                    delimitedWantedColumns = "T." & Replace$(Join(ConvertCollectionToArray(wantedFieldNamesSQL), ","), ",", ",T.")
                    
                    If priceTableAvailable Then delimitedWantedColumns = delimitedWantedColumns & ",P.Price"

                    With wantedFieldNamesSQL
                        sql = "SELECT " & delimitedWantedColumns & " FROM " & tableNameWithinDatabase & " as T" & _
                                IIf(priceTableAvailable, vbNewLine & priceTableSqlJoin, vbNullString) & _
                                vbNewLine & "WHERE T." & .item(ContractCodeFieldName) & " = ?" & _
                                vbNewLine & "ORDER BY T." & .item(DateFieldName) & " " & IIf(sortOrder = xlAscending, "ASC;", "DESC;")
                    End With
                End If

            ElseIf TryGetDatabaseDetails(OpenInterestEnum.FuturesOnly, eReport, tableNameToReturn:=futuresOnlyTableName, ignoreSqlServerDetails:=disableSqlServerConnection) Then

                'Spread column will now be the number of options offsseting an equivalent futures or option position.

                Dim futOptField$, isTotalColumn As Boolean, isTraderColumn As Boolean, _
                isLongColumn As Boolean, isSpreadColumn As Boolean, tempRef$, oiSelectionIndex As Long

                ReDim optionsOnlyFields(1 To filteredDatabaseInfoByEditedName.Count + IIf(priceTableAvailable, 1, 0))

                Set groupedTraderData = New Collection
                ' If TRUE then the spread will be removed and an equal number of contracts will be added to longs and shorts.
                ' Else if FUT + OPT - FUT < 0 then -1 from spread and + 1 to opposite side.

                iCount = 0
                
                For Each wantedField In filteredDatabaseInfoByEditedName
                    iCount = iCount + 1
                    With wantedField

                        currentFieldEdited = .EditedName
                        'This is effectively an inner join.
                        On Error GoTo Catch_WantedFieldMissing_OptionsOnly
                        sqlFieldName = wantedFieldNamesSQL(currentFieldEdited)
                        On Error GoTo Finally

                        futOptField = FutOpt & "." & sqlFieldName

                        isLongColumn = InStrB(currentFieldEdited, "long") <> 0
                        isSpreadColumn = InStrB(currentFieldEdited, "spread") <> 0

                        Select Case .DataType

                            Case adInteger, adBigInt, adDouble, adSmallInt, adTinyInt, adBigInt, adBinary
                                ' Assign a default value of Futures & Options - Futures Only.
                                optionsOnlyFields(iCount) = futOptField & "-" & FutOnly & "." & sqlFieldName
                                
                                isTraderColumn = InStrB(currentFieldEdited, "trader") > 0

                                ' If not oi or trader column
                                If Not (currentFieldEdited Like "*oi*" Or isTraderColumn) Then

                                    isTotalColumn = InStrB(currentFieldEdited, "tot") > 0
                                    ' If not change column.
                                    If InStrB(currentFieldEdited, "change") = 0 Then
                                        ' Calculate difference with a minimum value of 0. Exclude spread columns.
                                        ' Store column name in relevant collection.
                                        If Not (isSpreadColumn Or isTotalColumn) And createCommand Then
                                        
                                            If combinePositionsWithSpread Or disableNegativePositions Then
                                                Select Case Split(currentFieldEdited, "_", 2)(0)
                                                    Case "prod", "comm", "nonrept"
                                                        'These groups can't spread.
                                                    Case Else
                                                        'if FutOpt-FutOnly < 0 then the trader added positions that ended up in spread. Add the abs(of negative number) to opposite group.
                                                        'If above condition then ensure contract is removed from spread.
                                                        If isLongColumn Then
                                                            tempRef = "." & wantedFieldNamesSQL(Replace$(currentFieldEdited, "long", IIf(combinePositionsWithSpread, "spread", "short")))
                                                        Else
                                                            tempRef = "." & wantedFieldNamesSQL(Replace$(currentFieldEdited, "short", IIf(combinePositionsWithSpread, "spread", "long")))
                                                        End If
    
                                                        If combinePositionsWithSpread Then
                                                            'Column + Spread
                                                            optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & "+(" & FutOpt & tempRef & "-" & FutOnly & tempRef & ")"
                                                        ElseIf disableNegativePositions Then
                                                            'Column with min value of 0 + IIF(Opposing Side Options Only count< 0,ABS(Options Only opposing side),0)
                                                            optionsOnlyFields(iCount) = "(IIF((" & optionsOnlyFields(iCount) & ")>=0, " & optionsOnlyFields(iCount) & ", 0) + IIF((" & FutOpt & tempRef & "-" & FutOnly & tempRef & ")<0, ABS(" & FutOpt & tempRef & "-" & FutOnly & tempRef & "), 0))"
                                                        End If
                                                        'optionsOnlyFields(iCount) = Replace$(optionsOnlyFields(iCount), "FuncNull", IIf(connectedToSqlServer, "IsNull", "Nz"))
                                                End Select
                                            End If
                                            
                                        ElseIf isSpreadColumn And createCommand Then

                                            If combinePositionsWithSpread Then
                                                optionsOnlyFields(iCount) = "NULL"
                                            ElseIf disableNegativePositions Then
                                                ' If Long < 0 or short<0 subtract abs(value) from spread column for current trader group.
                                                For oiSelectionIndex = 0 To 1
                                                    tempRef = "." & wantedFieldNamesSQL(Replace$(currentFieldEdited, "spread", Array("long", "short")(oiSelectionIndex)))
                                                    optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & "-IIF(" & FutOpt & tempRef & "-" & FutOnly & tempRef & " < 0, ABS(" & FutOpt & tempRef & "-" & FutOnly & tempRef & "), 0)"
                                                Next oiSelectionIndex
                                            End If
                                        End If
                                        ' Store column with raw positions
                                        If Not (isSpreadColumn And combinePositionsWithSpread) Then
                                            traderGroupName = Split(currentFieldEdited, "_", 2)(0)
                                            On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                groupedTraderData(traderGroupName).Add currentFieldEdited, IIf(isLongColumn, "long", IIf(isSpreadColumn, "spread", "short"))
                                            On Error GoTo Finally
                                        End If

                                    ElseIf Not isTotalColumn And (disableNegativePositions Or combinePositionsWithSpread) Then
                                        ' If not change in total and negative positions have been disabled.
                                        ' Store change column name in relevant collection.
                                        
                                        traderGroupName = Split(currentFieldEdited, "_", 4)(2)

                                        Select Case traderGroupName
                                            Case "comm", "prod", "nonrept"
                                                'These groups don't have spreading to effect changes.
                                            Case Else
                                                If Not (isSpreadColumn And combinePositionsWithSpread) Then
                                                    On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                    groupedTraderData(traderGroupName).Add currentFieldEdited, IIf(isLongColumn, "longChange", IIf(isSpreadColumn, "spreadChange", "shortChange"))
                                                    On Error GoTo Finally
                                                End If
                                                
                                                If createCommand Then optionsOnlyFields(iCount) = "NULL"
                                        End Select

                                    End If
                                ElseIf createCommand And isTraderColumn Then
                                    optionsOnlyFields(iCount) = "NULL"
                                End If

                            Case NumericField
                                Select Case Split(currentFieldEdited, "_", 2)(0)
                                    Case "pct"
                                        Select Case currentFieldEdited
                                            Case "pct_of_oi_all", "pct_of_oi_old", "pct_of_oi_other"
                                                If createCommand Then optionsOnlyFields(iCount) = 100
                                            Case Else
                                                If Not (isSpreadColumn And combinePositionsWithSpread) Then
                                                    traderGroupName = Split(currentFieldEdited, "_", 5)(3)
                                                    On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                    groupedTraderData(traderGroupName).Add currentFieldEdited, IIf(isLongColumn, "longPct", IIf(isSpreadColumn, "spreadPct", "shortPct"))
                                                    On Error GoTo Finally
                                                End If
                                                If createCommand Then optionsOnlyFields(iCount) = "NULL"
                                        End Select
                                    Case "conc"
                                        'Concentration
                                        If createCommand Then optionsOnlyFields(iCount) = "NULL"
                                End Select
                            Case Else
                                If createCommand Then optionsOnlyFields(iCount) = futOptField
                        End Select

                    End With
OptionsOnly_AssignAlias:
                    If createCommand Then optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & " as " & currentFieldEdited
                    'Debug.Print optionsOnlyFields(iCount)
                Next wantedField

                If createCommand Then
                    If priceTableAvailable Then optionsOnlyFields(UBound(optionsOnlyFields)) = "P.Price as Price"

                    With wantedFieldNamesSQL
                        sql = " SELECT " & Join(optionsOnlyFields, ",") & " FROM (" & tableNameWithinDatabase & " as " & FutOpt & _
                                vbNewLine & "INNER JOIN " & futuresOnlyTableName & " as " & FutOnly & _
                                vbNewLine & "ON " & FutOpt & "." & .item(DateFieldName) & "=" & FutOnly & "." & .item(DateFieldName) & " AND " & FutOpt & "." & .item(ContractCodeFieldName) & "=" & FutOnly & "." & .item(ContractCodeFieldName) & ")" & _
                                IIf(priceTableAvailable, vbNewLine & Replace$(priceTableSqlJoin, "T.", FutOpt & "."), vbNullString) & _
                                vbNewLine & "WHERE " & FutOpt & "." & .item(ContractCodeFieldName) & "= ?" & _
                                vbNewLine & "ORDER BY " & FutOpt & "." & .item(DateFieldName) & " " & IIf(sortOrder = xlAscending, "ASC;", "DESC;")
                    End With
                End If
            End If

            delimitedWantedColumns = vbNullString: Set wantedFieldNamesSQL = Nothing
            On Error GoTo Finally
            
            Const ContractCodeParameterKey$ = "@ContractCode"
            
            With retrievalCommand
                If createCommand Then
                    .Parameters.Append .CreateParameter(ContractCodeParameterKey, IIf(connectedToSqlServer, adVarChar, adVarWChar), adParamInput, 10)
                    .Prepared = True
                    .CommandText = sql
                    .CommandType = adCmdText
                    Set .ActiveConnection = databaseConnection
                End If
                .Parameters(ContractCodeParameterKey).value = wantedContractCode
            End With
            
            Dim returnedData() As Variant, record As New ADODB.recordSet, serverResponseTimer As TimedTask
            
            With record
                .cursorLocation = adUseClient
                
                If Not profiler Is Nothing Then Set serverResponseTimer = profiler.StartSubTask(IIf(connectedToSqlServer, "SQL Server", "MS Access"))
                
                .Open retrievalCommand, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly
                
                If Not serverResponseTimer Is Nothing Then serverResponseTimer.EndTask
                
                On Error GoTo Data_Unavailable
                returnedData = TransposeData(.GetRows())
                On Error GoTo Finally
                .Close
            End With
            
            Set record = Nothing

            ' Calculate Changes and percent of OI.
            If wantedReportEnum = OpenInterestEnum.OptionsOnly Then

                Dim item As Collection, columnTarget As Long, pctOiColumn As Long, offsetN As Long, _
                calculatePctOI As Boolean, calculateChange As Boolean, oiSelectionForGroup$, positionColumn As Long
                ' Loop trader groups

                Const oiColumn As Long = 3
                offsetN = IIf(sortOrder = xlAscending, -1, 1)
                For Each item In groupedTraderData

                    On Error GoTo OptionsOnly_Next_GroupOiSelection
                    For oiSelectionIndex = 0 To 2

                        oiSelectionForGroup = Array("long", "short", "spread")(oiSelectionIndex)

                        calculatePctOI = HasKey(item, oiSelectionForGroup & "Pct")
                        calculateChange = HasKey(item, oiSelectionForGroup & "Change")

                        If calculatePctOI Then pctOiColumn = filteredDatabaseInfoByEditedName(item(oiSelectionForGroup & "Pct")).ColumnIndex
                        If calculateChange Then columnTarget = filteredDatabaseInfoByEditedName(item(oiSelectionForGroup & "Change")).ColumnIndex

                        If calculatePctOI Or calculateChange Then

                            positionColumn = filteredDatabaseInfoByEditedName(item(oiSelectionForGroup)).ColumnIndex
                            On Error GoTo Catch_OptionsOnlyCalculationError
                            For iCount = UBound(returnedData, 1) To LBound(returnedData, 1) Step -1
                                
                                If calculatePctOI Then
                                    If returnedData(iCount, oiColumn) <> 0 And returnedData(iCount, positionColumn) > 0 Then
                                        returnedData(iCount, pctOiColumn) = Round(100 * (returnedData(iCount, positionColumn) / returnedData(iCount, oiColumn)), 1)
                                    Else
                                        returnedData(iCount, pctOiColumn) = 0
                                    End If
                                End If
                                'This line won't generate an error unless missing data in database.
                                If calculateChange Then returnedData(iCount, columnTarget) = returnedData(iCount, positionColumn) - returnedData(iCount + offsetN, positionColumn)
                            Next iCount
                            On Error GoTo OptionsOnly_Next_GroupOiSelection

                        End If
OptionsOnly_Next_GroupOiSelection:
                        On Error GoTo -1

                    Next oiSelectionIndex
                Next item

                On Error GoTo Finally
            End If
            QueryDatabaseForContract = returnedData
        Else
            Err.Raise DbError.DatabaseConnectionFailed, Description:="Unable to locate database."
        End If
Finally:
        If Not record Is Nothing Then
            With record
                If .State = adStateOpen Then .Close
            End With
            Set record = Nothing
        End If

        If Not databaseConnection Is Nothing Then Set databaseConnection = Nothing

        If Err.Number <> 0 Then
            If Not retrievalCommand Is Nothing Then retainedADO.SetCommand Nothing, commandKey
            Call PropagateError(Err, "QueryDatabaseForContract", ConvertReportTypeEnum(eReport) & "_" & wantedContractCode & " - " & ConvertOpenInterestTypeToName(IIf(swappedToFuturesAndOptions, OpenInterestEnum.OptionsOnly, wantedReportEnum)))
        End If

        Exit Function
Data_Unavailable:
        If Err.Number = DbError.NoRecords Then
            If wantedReportEnum <> OpenInterestEnum.OptionsOnly Then
                AppendErrorDescription Err, "No data available for the current contract [" & wantedContractCode & "]."
            Else
                ' It's likely that the wanted contract doesn't exist in Futures Only so SQL statement fails.
                wantedReportEnum = OpenInterestEnum.FuturesAndOptions
                swappedToFuturesAndOptions = True
                With record
                    If .State = adStateOpen Then .Close
                End With

                Resume Create_SQL_Statement
            End If
        End If

        GoTo Finally
Catch_OptionsOnly_TraderGroup_Missing:
        Select Case Err.Number
            Case 5 ' Wanted collection with key [traderGroupName] not found.
                'On Error GoTo Finally
                groupedTraderData.Add New Collection, traderGroupName
                Resume
            Case Else
                AppendErrorDescription Err, "Error caught by > Catch_OptionsOnly_TraderGroup_Missing."
                GoTo Finally
        End Select
Catch_OptionsOnlyCalculationError:
        Select Case Err.Number
            Case 9
                'Subscript out of range when calculating change.
                Resume Next
            Case 6
                'Overflow error: Division by 0.
                returnedData(iCount, pctOiColumn) = 0
                Resume Next
            Case Else
                Resume OptionsOnly_Next_GroupOiSelection
        End Select
Catch_WantedFieldMissing_OptionsOnly:
        optionsOnlyFields(iCount) = IIf(currentFieldEdited = "price" And priceTableAvailable, "P.Price", "NULL")
        Resume OptionsOnly_AssignAlias
    End Function
    Public Sub Update_Database(dataToUpload() As Variant, versionToUpdate As OpenInterestEnum, eReport As ReportEnum, debugOnly As Boolean, suppliedFieldInfoByEditedName As Collection, Optional enableProgressBar As Boolean = False)
     '===================================================================================================================
    'Summary: Uploads the contents of dataToUpload to a database determined by other parameters.
    'Inputs:
    '       dataToUpload  - 2D array of C.O.T data to be uploaded.
    '       versionToUpdate - True if data being uploaded is Futures + Options combined.
    '       eReport - A reportTypeEnum used to specify which database to target.
    '       suppliedFieldInfoByEditedName - A Collection of FieldInfo instances used to describe columns contained within dataToUpload.
    '===================================================================================================================

        Dim matchedDatabaseFieldNamesByStandardName As Collection, iRow As Long, _
        legacyCombinedTableName$, legacyDatabasePath$, overwriteDatesForDebugging As Boolean, doesPriceTableExist As Boolean, _
        uploadingLegacyCombinedData As Boolean, uploadToDatabase As Boolean, dateColumnIndex&

        #Const DebugActive = False
        #Const UseLoop = False
        
        Const dateFieldKey$ = "report_date_as_yyyy_mm_dd"
        
        dateColumnIndex = suppliedFieldInfoByEditedName(dateFieldKey).ColumnIndex
        
        If debugOnly Then
            uploadToDatabase = MsgBox("Debug Active: Do you want to upload data to database?", vbYesNo) = vbYes
            If uploadToDatabase Then
            
                Const DebuggingDate As Date = #1/1/3000#
                
                If MsgBox("Replace dates with year 3000?", vbYesNo) = vbYes Then
                    On Error GoTo 0
                    If Not HasKey(suppliedFieldInfoByEditedName, dateFieldKey) Then Err.Raise vbObjectError + 587, "Update_Database", "Date field key not found."
                    overwriteDatesForDebugging = True
                End If
                
            End If
        Else
            uploadToDatabase = True
        End If

        Dim databaseFieldNamesRecord As ADODB.recordSet, databaseConnection As ADODB.Connection, tableToUpdateName$, connectingToSqlServer As Boolean

        On Error GoTo Finally

        Set databaseConnection = GetStoredAdoClass(eReport).Connection

        uploadingLegacyCombinedData = (eReport = ReportEnum.eLegacy And versionToUpdate = OpenInterestEnum.FuturesAndOptions)

        If TryGetDatabaseDetails(versionToUpdate, eReport, databaseConnection, tableToUpdateName, isSqlServerDetail:=connectingToSqlServer, doesPriceTableExist:=doesPriceTableExist) Then

            With databaseConnection
                If .State = adStateClosed Then .Open
                'Gets a Record of all field names within the database.
                Set databaseFieldNamesRecord = GetTableFieldsRecordSet(databaseConnection, tableToUpdateName)
                ' Get a collection of FieldInfo instances with matching fields for input and target.
                Set matchedDatabaseFieldNamesByStandardName = FilterDatabaseFieldsWithLocalInfo(databaseFieldNamesRecord, suppliedFieldInfoByEditedName)
            End With

            databaseFieldNamesRecord.Close
                        
            Dim uploadCommand As ADODB.command, wantedField As FieldInfo, cmdParameter As ADODB.Parameter, _
            iColumn&, wantedColumn As Long, isDateParameter As Boolean, minimumDate As Date
            
            Dim tableFieldNames$(), fieldValues$(), errorDuringTransaction As Boolean, standardName$, _
            csvFilteredData() As Variant, tableKeyByCsvHeader As Object, isPrimaryKey As Boolean

            Set uploadCommand = New ADODB.command ' CreateObject("ADODB.Command")

            With uploadCommand

                .ActiveConnection = databaseConnection
                .CommandType = adCmdText
                .Prepared = True

                ReDim fieldValues(matchedDatabaseFieldNamesByStandardName.Count - 1)
                ReDim tableFieldNames(UBound(fieldValues))
                
                #If Not UseLoop Then
                    Set tableKeyByCsvHeader = GetDictionaryObject()
                    ReDim csvFilteredData(LBound(dataToUpload, 1) To UBound(dataToUpload, 1) + 1, LBound(dataToUpload, 2) To matchedDatabaseFieldNamesByStandardName.Count)
                #End If
                
                On Error GoTo Catch_ParamaterCreationFailure
                ' Create a parameter for each field present in matchedDatabaseFieldNamesByStandardName
                With .Parameters
                    iColumn = LBound(fieldValues)
                    For Each wantedField In matchedDatabaseFieldNamesByStandardName
                        
                        With wantedField
                            standardName = .EditedName
                            ' Create parameter based on Field object stored inside recordset.
                            With databaseFieldNamesRecord.Fields(IIf(connectingToSqlServer, standardName, .DatabaseName))
                                Set cmdParameter = uploadCommand.CreateParameter(standardName, .Type, adParamInput, size:=.DefinedSize, value:=Null)
                                Select Case .Type
                                    Case adNumeric, adDecimal
                                        With cmdParameter
                                            .NumericScale = 2
                                            .Precision = 5
                                        End With
                                End Select
                            End With

                            tableFieldNames(iColumn) = .DatabaseNameForSQL
                            fieldValues(iColumn) = "?"
                            
                            #If Not UseLoop Then
                                wantedColumn = .ColumnIndex
                                ' If primary key.
                                If InStr(standardName, "yyyy") > 0 Or standardName = "cftc_contract_market_code" Then
                                    tableKeyByCsvHeader.item(standardName) = tableFieldNames(iColumn)
                                End If
                            #End If
                            
                        End With
                        
                        iColumn = iColumn + 1
                        
                        #If UseLoop Then
                            .Append cmdParameter
                        #Else
                            'iColumn started at 0 so the +1 above will set the index at the correct value.
                            Dim csvCodeColumn&, csvDateColumn&
                            With cmdParameter
                                isDateParameter = InStr(.name, "yyyy") > 0
                                If .name = "cftc_contract_market_code" Then csvCodeColumn = iColumn
                                
                                'Place header in first row of array.
                                csvFilteredData(LBound(csvFilteredData, 1), iColumn) = .name
                            End With
                            
                            For iRow = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
                                If isDateParameter Then
                                    csvFilteredData(iRow + 1, iColumn) = CLng(IIf(overwriteDatesForDebugging, DebuggingDate, dataToUpload(iRow, wantedColumn)))
                                    If minimumDate < dataToUpload(iRow, wantedColumn) Or iRow = LBound(dataToUpload, 1) Then minimumDate = dataToUpload(iRow, wantedColumn)
                                    If iRow = UBound(dataToUpload, 1) Then csvDateColumn = iColumn
                                Else
                                    csvFilteredData(iRow + 1, iColumn) = dataToUpload(iRow, wantedColumn)
                                End If
                            Next
                        #End If

                    Next wantedField
                End With
                              
                #If UseLoop Then
                
                    If overwriteDatesForDebugging Then
                        On Error GoTo Catch_ReportDateParameterMissing
                        Set cmdParameter = .Parameters(dateFieldKey)
                    End If

                    On Error GoTo Finally
                
                    .CommandText = "Insert Into " + tableToUpdateName + "(" + Join(tableFieldNames, ",") + ") Values (" + Join(fieldValues, ",") + ");"
                    
                    Erase fieldValues: Erase tableFieldNames
                
                    Dim dataCollectionsByDate As Object, selectedDateCollection As Collection, _
                    dataRowA() As Variant, collectionDate As Long, sortedAscCollectionDates As Variant, iDateKey&
    
                    ReDim dataRowA(LBound(dataToUpload, 2) To UBound(dataToUpload, 2))
                    
                    Set dataCollectionsByDate = GetDictionaryObject()
                    
                    ' Load each row of dataToUpload into a Dictionary keyed by date and store it in dataCollectionsByDate.
                    With dataCollectionsByDate
                        For iRow = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
    
                            For iColumn = LBound(dataToUpload, 2) To UBound(dataToUpload, 2)
                                dataRowA(iColumn) = dataToUpload(iRow, iColumn)
                            Next iColumn
                            ' Determine if selectedDateCollection should be swapped or if a new collection should be created.
                            If collectionDate <> dataRowA(dateColumnIndex) Then
                                collectionDate = dataRowA(dateColumnIndex)
                                If Not .Exists(collectionDate) Then .Add collectionDate, New Collection
                                Set selectedDateCollection = .item(collectionDate)
                            End If
                            
                            selectedDateCollection.Add dataRowA
                        Next iRow
                        sortedAscCollectionDates = .Keys
                    End With
                        
                    ' Sort keys to ensure data is uploaded in the right order (For error fixing purposes).
                    If LBound(sortedAscCollectionDates) <> UBound(sortedAscCollectionDates) Then
                        Quicksort sortedAscCollectionDates, LBound(sortedAscCollectionDates), UBound(sortedAscCollectionDates)
                    End If
                    
                    Application.StatusBar = "Uploading " & UBound(dataToUpload, 1) & " records to " & tableToUpdateName
                    
                    If enableProgressBar Then
                        With Progress_Bar
                            .InitializeValues maximumLoopCount:=UBound(dataToUpload, 1)
                            .Show
                        End With
                    End If
                    
                    Dim recordInsertionCount&, progressBarUpdateRate&
                    
                    Const percentUpdate As Double = 0.03
                    
                    progressBarUpdateRate& = Application.Max(UBound(dataToUpload, 1) * percentUpdate, 1)

                    For iDateKey = LBound(sortedAscCollectionDates) To UBound(sortedAscCollectionDates)
    
                        Set selectedDateCollection = dataCollectionsByDate.item(sortedAscCollectionDates(iDateKey))
                        
                        databaseConnection.BeginTrans
                        
                        For iRow = 1 To selectedDateCollection.Count
    
                            dataRowA = selectedDateCollection(iRow)
                            On Error GoTo Catch_ParameterValue_AssignmentFailure
    
                            For Each cmdParameter In .Parameters
                                With cmdParameter
                                    ' .Name property is standardized.
                                    wantedColumn = matchedDatabaseFieldNamesByStandardName(.name).ColumnIndex
    
                                    If Not (IsError(dataRowA(wantedColumn)) Or IsEmpty(dataRowA(wantedColumn)) Or IsNull(dataRowA(wantedColumn))) Then
                                        If VarType(dataRowA(wantedColumn)) <> vbString Then
                                            .value = dataRowA(wantedColumn)
                                        ElseIf dataRowA(wantedColumn) = "." Or LenB(Trim$(dataRowA(wantedColumn))) = 0 Then
                                            .value = Null
                                        Else
                                            .value = dataRowA(wantedColumn)
                                        End If
                                    Else
                                        .value = Null
                                    End If
                                End With
Next_Parameter:             Next cmdParameter
    
                            If overwriteDatesForDebugging Then .Parameters(dateFieldKey).value = DebuggingDate
                            
                            On Error GoTo Catch_UploadExecutionFailed
                            If uploadToDatabase Then .Execute Options:=adCmdText Or adExecuteNoRecords
                                                    
                            On Error Resume Next
                            recordInsertionCount = recordInsertionCount + 1
                            If enableProgressBar And recordInsertionCount Mod progressBarUpdateRate = 0 Then Progress_Bar.IncrementBar progressBarUpdateRate
                            On Error GoTo Finally
                        Next iRow
                    
                        Application.StatusBar = vbNullString
                        
                        On Error GoTo Catch_TransactionCommitError
                        
                        databaseConnection.CommitTrans
                        
                        On Error GoTo Finally
                        
                    Next iDateKey
                #Else
                    On Error GoTo Finally
                    BulkInsertIntoTable csvFilteredData, tableToUpdateName, tableFieldNames, .ActiveConnection, True, True, uploadToDatabase, columnIndexesForTextFormat:=Array(csvCodeColumn), primaryKeyByCsvName:=tableKeyByCsvHeader
                #End If
            End With
            
            On Error Resume Next
            If enableProgressBar Then Unload Progress_Bar
            On Error GoTo Finally
            
            Set uploadCommand = Nothing
            ' Update price data.
            If uploadToDatabase And Not (uploadingLegacyCombinedData Or overwriteDatesForDebugging Or connectingToSqlServer) Then

                If (Not doesPriceTableExist Or (doesPriceTableExist And versionToUpdate = FuturesAndOptions)) And TryGetDatabaseDetails(OpenInterestEnum.FuturesAndOptions, eLegacy, tableNameToReturn:=legacyCombinedTableName, databasePath:=legacyDatabasePath) Then

                    On Error GoTo CatchPriceUpdateFailed

                    With CreateObject("ADODB.Command")

                        If doesPriceTableExist And eReport <> eLegacy Then
                            ' Since price data is only held for Legacy_Combined retrieve data from table in different database.
                            .CommandText = "INSERT INTO " & PriceDataTableName & " (report_date_as_yyyy_mm_dd, cftc_contract_market_code, Price) " & _
                                            vbNewLine & "SELECT LegacyPrices.report_date_as_yyyy_mm_dd, LegacyPrices.cftc_contract_market_code, LegacyPrices.Price " & _
                                            "FROM ((" & tableToUpdateName & " as ReportTable " & _
                                            vbNewLine & "INNER JOIN [" & legacyDatabasePath & "]." & PriceDataTableName & " as LegacyPrices " & _
                                            vbNewLine & "ON ReportTable.report_date_as_yyyy_mm_dd = LegacyPrices.report_date_as_yyyy_mm_dd AND ReportTable.cftc_contract_market_code = LegacyPrices.cftc_contract_market_code)" & _
                                            vbNewLine & "LEFT JOIN " & PriceDataTableName & " as ReportPrices ON ReportPrices.report_date_as_yyyy_mm_dd = LegacyPrices.report_date_as_yyyy_mm_dd AND ReportPrices.cftc_contract_market_code = LegacyPrices.cftc_contract_market_code)" & _
                                            vbNewLine & "WHERE NOT ISNULL(LegacyPrices.Price) AND ISNULL(ReportPrices.cftc_contract_market_code) AND ReportTable.report_date_as_yyyy_mm_dd >= ?;"
                        ElseIf Not doesPriceTableExist Then

                            Const dateField$ = "[Report_Date_as_YYYY-MM-DD]", codeField$ = "[CFTC_Contract_Market_Code]"

                            .CommandText = "Update " & tableToUpdateName & " as ReportTable " & _
                                            vbNewLine & "INNER JOIN [" & legacyDatabasePath & "]." & legacyCombinedTableName & " as LegacyTable " & _
                                            vbNewLine & "ON LegacyTable." & dateField & "=ReportTable." & dateField & " AND LegacyTable." & codeField & "=ReportTable." & codeField & _
                                            vbNewLine & "SET ReportTable.[Price] = LegacyTable.[Price] " & _
                                            vbNewLine & "WHERE ReportTable." & dateField & ">=?;"
                        End If

                        If LenB(.CommandText) > 0 Then
                            .ActiveConnection = databaseConnection
                            .CommandType = adCmdText
                            
                            #If UseLoop Then
                                .Parameters.Append .CreateParameter("@FilterDate", IIf(connectingToSqlServer, adDBDate, adDate), value:=CDate(sortedAscCollectionDates(LBound(sortedAscCollectionDates))))
                            #Else
                                .Parameters.Append .CreateParameter("@FilterDate", IIf(connectingToSqlServer, adDBDate, adDate), value:=CDate(minimumDate))
                            #End If
                            
                            .Execute Options:=adCmdText Or adExecuteNoRecords
                        End If
                    End With

                    On Error GoTo Finally
                End If

            ElseIf uploadToDatabase And uploadingLegacyCombinedData And doesPriceTableExist And HasKey(suppliedFieldInfoByEditedName, "price") And Not overwriteDatesForDebugging Then
                Dim dateColumn&, contractCodeColumn&, priceColumn&

                With suppliedFieldInfoByEditedName
                    dateColumn = .item(dateFieldKey).ColumnIndex
                    contractCodeColumn = .item("cftc_contract_market_code").ColumnIndex
                    priceColumn = .item("price").ColumnIndex
                End With
                InsertIntoPriceTable dataToUpload, priceColumn, contractCodeColumn, dateColumn, databaseConnection
            End If

            If Not debugOnly Then
                With GetStoredReportDetails(eReport)
                    If .OpenInterestType.Value2 = versionToUpdate Then
                        'This will signal to worksheet activate events to update the currently visible data
                        .PendingUpdateInDatabase.Value2 = True
                    End If
                End With
            End If
        Else
            Err.Raise DbError.DatabaseConnectionFailed, Description:="Couldn't locate database: " & GetStoredReportDetails(eReport).FullReportName.Value2
        End If
        
Finally:
        #If DebugActive Then
            If Err.Number <> 0 Then
                DisplayErr Err, "Update_Database"
                Stop: Resume
            End If
        #End If

        If Not databaseFieldNamesRecord Is Nothing Then
            With databaseFieldNamesRecord
                If .State = adStateOpen Then .Close
            End With
            Set databaseFieldNamesRecord = Nothing
        End If

        If Not databaseConnection Is Nothing Then
            With databaseConnection
                If .State = adStateOpen Then
                    If Err.Number <> 0 Then
                        AppendErrorDescription Err, "Failed to update table " & tableToUpdateName & " in database " & .Properties("Data Source")
                        
                        Dim savedError As StoredError: Set savedError = HoldError(Err)
                        
                        On Error GoTo -1: On Error GoTo Catch_FailedRollBackTransaction
                        If errorDuringTransaction Then .RollbackTrans
                        On Error GoTo 0
                        Err = savedError.HeldError
                    End If
                End If
            End With
            Set databaseConnection = Nothing
        End If
        
        Application.StatusBar = vbNullString
        
        If Err.Number <> 0 Then
            If enableProgressBar Then Unload Progress_Bar
            PropagateError Err, "Update_Database"
        End If
        
        Exit Sub

Catch_FailedRollBackTransaction:
        savedError.pDescription = "Error while attempting rollback." & "    \  " & savedError.pDescription
        Resume Next
        
CatchPriceUpdateFailed:
        MsgBox "Non-Critical Error" & String$(2, vbNewLine) & "Failed to update " & tableToUpdateName & " price data using the Legacy_Combined table." & _
        vbNewLine & Err.Description
        Resume Next
        
Catch_ParamaterCreationFailure:
        If Not wantedField Is Nothing Then
            'Stop: Resume
            With wantedField
                AppendErrorDescription Err, "Failed to create a parameter for the " & .EditedName & " FieldInfo instance." & vbNewLine & _
                                            "DataType: " & .DataType
            End With
        Else
            AppendErrorDescription Err, "Failed to create parameter. [wantedField] is nothing."
        End If
        GoTo Finally
        
    #If UseLoop Then
       
Catch_TransactionCommitError:
        errorDuringTransaction = True
        AppendErrorDescription Err, "Attempted to commit transaction for date collection " & CDate(iDateKey) & "."
        GoTo Finally
        
Catch_ParameterValue_AssignmentFailure:
        If Err.Number = 9 Then
            'Subscript out of range error.
            AppendErrorDescription Err, "dataToUpload array isn't large enough for the current value of wantedColumn: " & wantedColumn
        ElseIf Not cmdParameter Is Nothing Then
            
            With cmdParameter
                ' The application uses an invalid type value for the current operation.
                If Err.Number = DbError.InvalidParameterAssignment Then
                    
                    #If DebugActive Then
                        Debug.Print "[cmdParameter] value assignment mismatch error. " & dataRowA(wantedColumn) & " should be of type " & .Type & vbNewLine & _
                                    Space$(4) & dataRowA(1) & " " & dataRowA(dateColumnIndex)
                    #End If

                    Select Case .Type
                        Case adVarWChar, adVarChar
                            If VarType(dataRowA(wantedColumn)) <> vbString Then
                                .value = CStr(dataRowA(wantedColumn))
                                Resume Next_Parameter
                            End If
                        Case Else
                             If IsNumeric(dataRowA(wantedColumn)) Then
                                Select Case .Type
                                    Case adNumeric, adDecimal
                                        .value = CDbl(dataRowA(wantedColumn))
                                    Case adInteger, adTinyInt, adBigInt, adDouble, adSmallInt
                                        .value = CLng(dataRowA(wantedColumn))
                                End Select
                                Resume Next_Parameter
                            End If
                    End Select
                    ' Failed to handle value.
                    .value = Null
                    Resume Next_Parameter
                End If

                AppendErrorDescription Err, "Value assignment for parameter '" & .name & "' failed." & vbNewLine & _
                                            "Parameter type: " & .Type & ", Array value: " & dataRowA(wantedColumn) & ", Value VarType: " & VarType(dataRowA(wantedColumn))
            End With
        ElseIf cmdParameter Is Nothing Then
            AppendErrorDescription Err, "Failed to assign value to parameter, [cmdParameter] is nothing."
        End If
        
        errorDuringTransaction = True
        GoTo Finally
        
Catch_UploadExecutionFailed:
        Select Case Err.Number
            ' Ignore if violation of primary key error.
            Case DbError.ProviderError, DbError.PrimaryKeyViolation_SQLSERVER
                Dim objError As ADODB.Error
                For Each objError In databaseConnection.Errors
                    With objError
                        If .Number = Err.Number And (Not connectingToSqlServer And .SqlState = 3022) Or (connectingToSqlServer And .NativeError = 2627) Then
                            databaseConnection.Errors.Clear
                            Resume Next
                        End If
                    End With
                Next objError
        End Select
        
        With HoldError(Err)
            On Error GoTo -1: On Error Resume Next
            With CreateObject("ADODB.Command")
                .CommndText = "DELETE FROM " & tableToUpdateName & " Where " & IIf(doesPriceTableExist, "report_date_as_yyyy_mm_dd", "[Report_Date_as_YYYY-MM-DD]") & " = ?;"
                .Parameters.Append .CreateParameter("@Date", IIf(connectingToSqlServer, adDBDate, adDate), value:=CDate(sortedAscCollectionDates(iDateKey)))
                .CommandType = adCmdText
                Set .ActiveConnection = databaseConnection
                .Execute Options:=adCmdText Or adExecuteNoRecords
            End With
            Err = .HeldError
        End With
        
        AppendErrorDescription Err, SummarizeAdoError(databaseConnection.Errors, Err.Number)
        errorDuringTransaction = True
        GoTo Finally
    #End If
        
Catch_ReportDateParameterMissing:
        AppendErrorDescription Err, dateFieldKey & " command parameter is missing."
        GoTo Finally
        
    End Sub
    Private Sub BulkInsertIntoTable(csvDataToUpload() As Variant, tableToUpdateName$, tableFieldNamesToInsertInto$(), _
                    conn As ADODB.Connection, headersAvailable As Boolean, ByVal eraseDataOnCompletion As Boolean, _
                    Optional executeSqlInsertion As Boolean = True, Optional csvFieldNameFilter As Variant, _
                    Optional columnIndexesForTextFormat As Variant, Optional primaryKeyByCsvName As Object)
    '===================================================================================================================
    'Summary: Bulk inserts data into a database table with a (.csv) file.
    'Inputs:
    '   csvDataToUpload - Array used to create a CSV file and then uploaded to a database.
    '   tableToUpdateName - Name of table to insert data into.
    '   tableFieldNamesToInsertInto - Field Names within [tableToUpdateName] to insert values into.
    '   conn                       - ADODB.Connection used to connect to a database that contains table [tableToUpdateName].
    '   headersAvailable           - Set to True if headers are in the first row of [csvDataToUpload].
    '   eraseDataOnCompletion      - Set to True if [csvDataToUpload] can be erased after creating a CSV.
    '   executeSqlInsertion        - Set to FALSE to skip insertion into database.
    '   csvFieldNameFilter         - Array of field names in csv used as a filter.
    '   columnIndexesForTextFormat - Array of column indexes to format as text.
    '   primaryKeyByCsvName        - Dictionary object used to JOIN csv table with primary key to avoid duplicate records.
    '===================================================================================================================
        On Error GoTo Propagate
        
        Const ProcedureName$ = "BulkInsertIntoTable", csvTableAliasName$ = "TMP"
        
        If Not IsMissing(csvFieldNameFilter) And Not IsArray(csvFieldNameFilter) Then
            Err.Raise vbObjectError + 10025, , "Variable 'csvFieldNameFilter' isn't an array."
        ElseIf Not IsMissing(columnIndexesForTextFormat) And Not IsArray(columnIndexesForTextFormat) Then
            Err.Raise vbObjectError + 10025, , "Variable 'columnIndexesForTextFormat' isn't an array."
        ElseIf Not IsMissing(columnIndexesForTextFormat) And Not headersAvailable Then
            Err.Raise vbObjectError + 10025, , "Please supply headers if you want to force certain columns to be exported as text."
        ElseIf Not primaryKeyByCsvName Is Nothing And TypeName(primaryKeyByCsvName) <> "Dictionary" Then
            Err.Raise vbObjectError + 10025, , "'primaryKeyByCsvName' musn't be a Dictionary."
        End If
        
        Dim sql$, tempFolder$, csvPath$, iColumn&, csvWorkbook As Workbook, columnWithinCsv&, _
        iRow&, wantedCsvFieldsByName As Object, csvHeader$
        
        Const csvName$ = "MoshiM_Database_Upload.csv"

        tempFolder = Environ("Temp") & Application.PathSeparator
        csvPath = tempFolder & csvName
        
        If headersAvailable And Not (IsMissing(csvFieldNameFilter) And IsMissing(columnIndexesForTextFormat)) Then
            Set wantedCsvFieldsByName = GetDictionaryObject()
            With wantedCsvFieldsByName
                For iColumn = LBound(csvDataToUpload, 2) To UBound(csvDataToUpload, 2)
                    '.Item(Header) = Field Selection via SQL.
                    csvHeader = csvDataToUpload(LBound(csvDataToUpload, 1), iColumn)
                    .item(csvHeader) = csvHeader
                Next iColumn
            End With
        End If
        
        Set csvWorkbook = Workbooks.Add
        
        With csvWorkbook
            .Windows(1).Visible = False
            With .Worksheets(1)
                If Not IsMissing(columnIndexesForTextFormat) Then
                    ' MS ACCESS might attempt to read values containing just numbers as INT rather than TEXT
                    ' and generate errors if non-numbers are found.
                    For iColumn = LBound(columnIndexesForTextFormat) To UBound(columnIndexesForTextFormat)
                        
                        columnWithinCsv = columnIndexesForTextFormat(iColumn)
                        .columns(columnWithinCsv).NumberFormat = "@"
                        'https://catolytics.com/posts/2018/0007-etl-solving-microsoft-access-import-errors/
                        
                        If Not wantedCsvFieldsByName Is Nothing Then
                            ' Attach an identifier to all columns that MUST be imported as text.
                            For iRow = LBound(csvDataToUpload, 1) + 1 To UBound(csvDataToUpload, 1)
                                csvDataToUpload(iRow, columnWithinCsv) = "@" & csvDataToUpload(iRow, columnWithinCsv)
                            Next iRow
                            
                            csvHeader = csvDataToUpload(LBound(csvDataToUpload, 1), columnWithinCsv)
                            ' Use SQL statement to import without added character.
                            With wantedCsvFieldsByName
                                .item(csvHeader) = "RIGHT(" & csvHeader & ",LEN(" & csvHeader & ")-1)"
                                If Not primaryKeyByCsvName Is Nothing Then
                                    If Not primaryKeyByCsvName.Exists(csvHeader) Then
                                        ' Field will be part of outer selection and needs to have the table alias added.
                                        .item(csvHeader) = Replace$(.item(csvHeader), csvHeader, csvTableAliasName & "." & csvHeader)
                                    End If
                                End If
                            End With
                        End If
                        
                    Next iColumn
                End If
                .Range("A1").Resize(UBound(csvDataToUpload, 1), UBound(csvDataToUpload, 2)).Value2 = csvDataToUpload
            End With

            If eraseDataOnCompletion Then Erase csvDataToUpload

            Application.DisplayAlerts = False
            .SaveAs csvPath, xlCSV
            Application.DisplayAlerts = True
            .Close
        End With
        Set csvWorkbook = Nothing
        
        Dim sqlInsertionConstraint$, sqlJoinConditions() As Variant, II As Variant, fieldAlias$, primaryKeyAliasDict As Object
        
        If Not primaryKeyByCsvName Is Nothing Then
            ' https://stackoverflow.com/questions/14393494/ms-access-insert-into-without-duplicate
            sqlJoinConditions = primaryKeyByCsvName.Keys
            Set primaryKeyAliasDict = GetDictionaryObject()

            For iColumn = LBound(sqlJoinConditions) To UBound(sqlJoinConditions)
                
                csvHeader = sqlJoinConditions(iColumn)
                fieldAlias = "PKEY" & iColumn
                ' Assign an alias for the csvHeader mapped to the primary key within [tableToUpdateName].
                With wantedCsvFieldsByName
                    .item(csvHeader) = .item(csvHeader) & " as " & fieldAlias
                End With

                ' Dictionary item will be used to select the correct field in the final SQL selection.
                primaryKeyAliasDict.item(csvHeader) = fieldAlias
                sqlJoinConditions(iColumn) = "Main." & primaryKeyByCsvName.item(csvHeader) & "=" & csvTableAliasName & "." & fieldAlias

            Next iColumn
            
            fieldAlias = vbNullString
            ' Get field name of primary key within database table.
            II = primaryKeyByCsvName.Items
            ' Only need to check for nulls on one key since it's a primary key.
            sqlInsertionConstraint = vbNewLine & "LEFT JOIN " & tableToUpdateName & " as Main ON " & Join(sqlJoinConditions, " AND ") & _
                                     vbNewLine & "Where IsNull(Main." & II(LBound(sqlJoinConditions)) & ")"
            Erase sqlJoinConditions: Erase II
        End If
        
        Dim finalSqlSelection$, innerSelectionCsvFields$, tempDict As Object
        
        innerSelectionCsvFields = "*"
        
        If Not wantedCsvFieldsByName Is Nothing Then

            With wantedCsvFieldsByName
                ' Just get wanted items with field name filter.
                If Not IsMissing(csvFieldNameFilter) Then
                    Set tempDict = GetDictionaryObject()
                    
                    For Each II In csvFieldNameFilter
                        If .Exists(II) Then
                            tempDict.item(II) = .item(II)
                        Else
                            Err.Raise 9, Description:="'wantedCsvFieldsByName' doesn't contain the csv header " & II & "."
                        End If
                    Next II
                    Set wantedCsvFieldsByName = tempDict
                End If
                
                If Not primaryKeyAliasDict Is Nothing Then
                    ' Build inner selection query to select all along with primary key aliases.
                    For Each II In primaryKeyAliasDict.Keys
                        If .Exists(II) Then
                            innerSelectionCsvFields = innerSelectionCsvFields & "," & .item(II)
                            .item(II) = primaryKeyAliasDict.item(II)
                        Else
                            Err.Raise 9, Description:="'wantedCsvFieldsByName' doesn't have the csv header(key) from 'primaryKeyAliasDict' " & II & " as a key."
                        End If
                    Next II
                    Set primaryKeyAliasDict = Nothing
                End If
                
                If .Count > 0 Then
                    ' Section is used to create a comma separated list of wanted fields in the outer selection.
                    II = .Items
                    For iColumn = LBound(II) To UBound(II)
                        ' Ignore any fields that are forced to be text and a primary key.
                        ' They should already be defined properly.
                        If Not II(iColumn) Like "RIGHT(*,LEN(*)-1)" Then
                            II(iColumn) = csvTableAliasName & "." & II(iColumn)
                        End If
                    Next iColumn
                    finalSqlSelection = Join(II, ",")
                    Erase II
                Else
                    Err.Raise vbObjectError + 9200, Description:="'wantedCsvFieldsByName' doesn't contain any key value pairs."
                End If
                
            End With
            Set wantedCsvFieldsByName = Nothing
            Set tempDict = Nothing
        End If
        
        If IsSqlServerConnection(conn) Then
            'https://stackoverflow.com/questions/12365992/sql-insert-multiple-rows
            'https://learn.microsoft.com/en-us/sql/t-sql/statements/bulk-insert-transact-sql?view=sql-server-ver16
            'https://www.codeproject.com/Articles/439843/Handling-BULK-Data-insert-from-CSV-to-SQL-Server
            Err.Raise DbError.NotImplemented, Description:="SQL Server bulk insert command hasn't been configured."
            
            sql = "BULK INSERT " & tableToUpdateName & " FROM '" & csvPath & "' WITH (FIRSTROW = 2,FORMAT = 'CSV',TABLOCK);"
            'Add filtration to avoid duplicate errors.
        Else
            sql = "Insert Into " & tableToUpdateName & "(" + Join(tableFieldNamesToInsertInto, ",") & ")" & vbNewLine & _
                  "Select " & finalSqlSelection & " FROM (Select " & innerSelectionCsvFields & vbNewLine & "From [Text;FMT=Delimited;IMEX=1;HDR=" & IIf(headersAvailable, "YES", "NO") & ";DATABASE=" & tempFolder & "].[" & csvName & "]) as " & csvTableAliasName _
                   & sqlInsertionConstraint & ";"
        End If
        
        Dim recordsAffectedCount&
        
        On Error GoTo Catch_Failed_BulkInsert
        If executeSqlInsertion Then conn.Execute sql, recordsAffectedCount, Options:=adExecuteNoRecords Or adCmdText

        On Error Resume Next
        Kill csvPath
        
        Exit Sub
Propagate:
        PropagateError Err, ProcedureName
Catch_Failed_BulkInsert:
        Debug.Print sql & String$(2, vbNewLine)
        PropagateError Err, ProcedureName, "SQL Insertion Failure"
    End Sub
    
    Private Function SummarizeAdoError(errorCollection As ADODB.Errors, wantedErrorNumber&) As String

        Dim adoError As ADODB.Error
        
        For Each adoError In errorCollection
            With adoError
                If .Number = wantedErrorNumber Then
                    SummarizeAdoError = "Error #" & .Number & _
                                    " " & .Description & vbCrLf & _
                                    "NativeError: " & .NativeError & vbCrLf & _
                                    "SQLState: " & .SqlState & vbCrLf & _
                                    "Reported by: " & .Source & vbCrLf & _
                                    "Help file: " & .HelpFile & vbCrLf & _
                                    "Help Context ID: " & .HelpContext
                    Exit For
                End If
            End With
                        
        Next adoError
        
    End Function
    
    Sub DeleteAllCFTCDataFromDatabaseByDate()
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_Description = "Asks the user for a minimum date and then deletes all data greater than or equal to that in all available databases."
Attribute DeleteAllCFTCDataFromDatabaseByDate.VB_ProcData.VB_Invoke_Func = " \n14"
    '===================================================================================================================
    'Summary: Asks the user for a minimum date and then deletes all data greater than or equal to that in all available databases.
    '===================================================================================================================
        Dim wantedDate As Date, eReport As Variant, combinedType As Variant
        
        On Error GoTo Display
        
        wantedDate = InputBox("Input date for which all data >= will be deleted in the format YYYY-MM-DD (year-month-day)." & vbNewLine & "Ex: 2024-05-10 for May 10, 2024")
        
        If MsgBox("Is this date correct? " & Format$(wantedDate, "mmmm dd, yyyy"), vbYesNo) = vbYes Then
            For Each eReport In Array(eLegacy, eDisaggregated, eTFF)
                For Each combinedType In Array(OpenInterestEnum.FuturesAndOptions, OpenInterestEnum.FuturesOnly)
                    DeleteCftcDataFromSpecificDatabase wantedDate, CLng(eReport), CLng(combinedType)
                Next
            Next
        End If
        Exit Sub
Display:
        DisplayErr Err, "DeleteAllCFTCDataFromDatabaseByDate"
    End Sub
    Private Sub DeleteCftcDataFromSpecificDatabase(smallest_date As Date, eReport As ReportEnum, versionToDelete As OpenInterestEnum)
    '===================================================================================================================
    'Summary: Deletes COT data from database that is as recent as smallest_date.
    'Inputs: smallest_date - all rows with a date value >= to this will be deleted.
    '        eReport - One of L,D,T to repersent which database to delete from.
    '        versionToDelete - true for futures+options and false for futures only.
    '===================================================================================================================

        Dim tableName$, databaseConnection As ADODB.Connection, connectedToSqlServer As Boolean, priceTableAvailable As Boolean
        
        On Error GoTo No_Table
        
        Set databaseConnection = GetStoredAdoClass(eReport).Connection
        
        If TryGetDatabaseDetails(versionToDelete, eReport, databaseConnection, tableName, isSqlServerDetail:=connectedToSqlServer, doesPriceTableExist:=priceTableAvailable) Then

            With databaseConnection
                If .State = adStateClosed Then .Open
                With CreateObject("ADODB.Command")
                    .ActiveConnection = databaseConnection
                    .CommandText = "DELETE FROM " & tableName & " WHERE " & IIf(priceTableAvailable, "report_date_as_yyyy_mm_dd", "[Report_Date_as_YYYY-MM-DD]") & " >= ?;"
                    .CommandType = adCmdText
                    .Parameters.Append .CreateParameter("@smallestDate", IIf(connectedToSqlServer, adDBDate, adDate), adParamInput, value:=smallest_date)
                    .Execute Options:=adCmdText Or adExecuteNoRecords
                End With
                '.Close
            End With

        End If

        Set databaseConnection = Nothing
        Exit Sub
No_Table:
        If Not databaseConnection Is Nothing Then
'            With databaseConnection
'                If .State = adStateOpen Then .Close
'            End With
            Set databaseConnection = Nothing
        End If
        PropagateError Err, "DeleteCftcDataFromSpecificDatabase"
    End Sub

    Public Function TryGetLatestDate(ByRef latestDate As Date, eReport As ReportEnum, ByVal versionToQuery As OpenInterestEnum, queryIceContracts As Boolean, Optional databaseConnection As ADODB.Connection) As Boolean
    '===================================================================================================================
    'Summary: Returns the date for the most recent data within a database.
    'Inputs:
    '   latestDate - ByRef param that will store the most recent date in the database.
    '   eReport - One of L,D,T to repersent which database to query.
    '   versionToQuery - OpenInterestEnum used to select a table within the database to query.
    '   queryIceContracts - True to filter for ICE contracts.
    'Returns: True if SQL query executes successfully; otherwise, False.
    '===================================================================================================================
        Dim tableName$, sql$, conn As ADODB.Connection, isSqlServerConn As Boolean, successfulConnection As Boolean

        On Error GoTo Connection_Unavailable

        If versionToQuery = OptionsOnly Then versionToQuery = FuturesAndOptions
        
        If queryIceContracts And eReport <> eDisaggregated Then
            Err.Raise DbError.UseDisaggregatedReport, Description:="You must query the Disaggregated report if querying ICE data."
        End If

        If databaseConnection Is Nothing Then
            Set conn = GetStoredAdoClass(eReport).Connection
            successfulConnection = TryGetDatabaseDetails(versionToQuery, eReport, conn, tableName, , True, isSqlServerDetail:=isSqlServerConn)
        Else
            Set conn = databaseConnection
            successfulConnection = True
            If IsSqlServerConnection(conn) Then
                isSqlServerConn = True
                tableName = GetSqlServerTableName(eReport, versionToQuery)
            Else
                ' Just get the table name.
                TryGetDatabaseDetails versionToQuery, eReport, , tableName, , True, True
            End If
        End If

        If successfulConnection Then
            With conn
                If .State = adStateClosed Then .Open

                With GetFieldInfoForDatabaseTable(conn, tableName)
                    sql = "SELECT MAX(" & .item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL & ") FROM " & tableName & _
                    vbNewLine & " WHERE " & IIf(queryIceContracts, vbNullString, "NOT ") & "LCASE(" & .item("market_and_exchange_names").DatabaseNameForSQL & ") LIKE 'ice%ice%';"
                    If isSqlServerConn Then sql = Replace$(sql, "LCASE", "LOWER")
                End With
                
                With .Execute(sql, Options:=adCmdText)
                    If Not (.EOF And .BOF) Then
                        latestDate = .Fields(0)
                    Else
                        latestDate = #1/1/1970#
                    End If
                    .Close
                End With
            End With
            TryGetLatestDate = True
        End If
Connection_Unavailable:
        Set conn = Nothing
        If Err.Number <> 0 Then PropagateError Err, "TryGetLatestDate"
    End Function
    Private Sub UpdateDatabasePricesWithArray(data() As Variant, eReport As ReportEnum, versionToUpdate As OpenInterestEnum, priceColumn As Long)
    '===================================================================================================================
    'Summary: Updates price data or inserts records where needed.
    'Inputs:
    '   data - Array that holds all necessary data for price updating.
    '   eReport - ReportEnum used to target a specific database or table.
    '   versionToUpdate - OpenInterestEnum used to select a table.
    '===================================================================================================================
        Dim sql$, tableToUpdateName$, iRow As Long, insertPriceCMD As ADODB.command, connectedToSqlServer As Boolean, _
        databaseConnection As ADODB.Connection, updatePriceCMD As ADODB.command, contractCodeColumn As Long, priceTableAvailable As Boolean, recordsAffectedCount&

        Const date_column As Long = 1

        contractCodeColumn = priceColumn - 1
        
        On Error GoTo Finally
        
        Set databaseConnection = GetStoredAdoClass(eReport).Connection

        If TryGetDatabaseDetails(versionToUpdate, eReport, databaseConnection, tableToUpdateName, doesPriceTableExist:=priceTableAvailable, isSqlServerDetail:=connectedToSqlServer) Then
            
            With databaseConnection
                If .State = adStateClosed Then .Open
                If priceTableAvailable Then tableToUpdateName = PriceDataTableName

                With GetFieldInfoForDatabaseTable(databaseConnection, tableToUpdateName)
                    sql = "UPDATE " & tableToUpdateName & _
                        vbNewLine & "SET [Price] = ? " & _
                        vbNewLine & "WHERE " & .item("cftc_contract_market_code").DatabaseNameForSQL & " = ? AND " & .item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL & "= ?;"
                End With
            End With

            Set updatePriceCMD = New ADODB.command 'CreateObject("ADODB.Command")

            With updatePriceCMD
                With .Parameters
                    .Append updatePriceCMD.CreateParameter("Price", adCurrency, adParamInput)
                    .Append updatePriceCMD.CreateParameter("Contract Code", IIf(connectedToSqlServer, adVarChar, adVarWChar), adParamInput, size:=10)
                    .Append updatePriceCMD.CreateParameter("Date", IIf(connectedToSqlServer, adDBDate, adDate), adParamInput)
                End With
                .ActiveConnection = databaseConnection
                .CommandType = adCmdText
                .CommandText = sql
                .Prepared = True
            End With

            If priceTableAvailable Then
            
                Set insertPriceCMD = New ADODB.command 'CreateObject("ADODB.Command")
                With insertPriceCMD
                    With .Parameters
                        Dim param As ADODB.Parameter
                        For Each param In updatePriceCMD.Parameters
                            .Append param
                        Next param
                    End With
                    .ActiveConnection = databaseConnection
                    .CommandType = adCmdText
                    .CommandText = "INSERT INTO " & PriceDataTableName & " (Price,cftc_contract_market_code,report_date_as_yyyy_mm_dd) Values(?,?,?);"
                    .Prepared = True
                End With
                
            End If
            
            Dim priceMissing As Boolean
            
            For iRow = LBound(data, 1) To UBound(data, 1)
                priceMissing = IsEmpty(data(iRow, priceColumn)) Or IsNull(data(iRow, priceColumn))
                                
                With updatePriceCMD
                
                    With .Parameters
                        .item("Price").value = IIf(priceMissing, Null, data(iRow, priceColumn))
                        .item("Contract Code").value = data(iRow, contractCodeColumn)
                        .item("Date").value = data(iRow, date_column)
                    End With
                    
                    On Error GoTo Catch_ExecutionError
                    .Execute recordsAffectedCount, Options:=adCmdText Or adExecuteNoRecords

                    If priceTableAvailable And recordsAffectedCount = 0 And Not priceMissing Then
                        insertPriceCMD.Execute Options:=adCmdText Or adExecuteNoRecords
                    End If
                    On Error GoTo Finally
                    
                End With

            Next iRow
        End If
Finally:
        If Not databaseConnection Is Nothing Then Set databaseConnection = Nothing
        Set updatePriceCMD = Nothing
        Set insertPriceCMD = Nothing
        If Err.Number <> 0 Then PropagateError Err, "UpdateDatabasePricesWithArray"
        Exit Sub
Catch_ExecutionError:
        If databaseConnection.Errors.Count > 0 Then
            AppendErrorDescription Err, SummarizeAdoError(databaseConnection.Errors, Err.Number)
        End If
        GoTo Finally
    End Sub
    Public Sub DownloadPriceDataForActiveContract()
Attribute DownloadPriceDataForActiveContract.VB_Description = "Retrieves dates from the currently active data table, relevant price data and uploads to available databases.\r\n"
Attribute DownloadPriceDataForActiveContract.VB_ProcData.VB_Invoke_Func = " \n14"
    '========================================================================================================================
    ' Summary - Retrieves dates from the currently active data table, relevant price data and uploads to available databases.
    '========================================================================================================================
        Dim Worksheet_Data() As Variant, ws As Variant, price_column As Long, _
        eReport As ReportEnum, availableContractInfo As Collection, contractCode$, _
        Source_Ws As Worksheet
        
        On Error GoTo Display
        
        For Each ws In Array(LC, DC, TC)
            If ThisWorkbook.ActiveSheet Is ws Then

                Set Source_Ws = ws
                eReport = ThisWorkbook.Worksheets(Source_Ws.name).WorksheetReportEnum()

                With GetStoredReportDetails(eReport)
                    contractCode = .CurrentContractCode.Value2
                    price_column = .RawDataCount.Value2 + 1
                End With

                With Get_CftcDataTable(eReport).DataBodyRange
                    Worksheet_Data = .Resize(.Rows.Count, price_column).value
                End With

                Set availableContractInfo = GetAvailableContractInfo(True)

                If HasKey(availableContractInfo, contractCode) Then
                    If TryGetPriceData(Worksheet_Data, price_column, availableContractInfo(contractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
                        'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
                        UpdateDatabasePricesWithArray Worksheet_Data, eLegacy, OpenInterestEnum.FuturesAndOptions, priceColumn:=price_column

                        'Overwrites all other database tables with price data from Legacy_Combined
                        If Not DoesUserPermit_SqlServer() Then HomogenizeWithLegacyCombinedPrices contractCode
                        RefreshTableData eReport, True
                    Else
                        MsgBox "Unable to retrieve price data."
                    End If
                Else
                    MsgBox "A symbol is unavailable for: [ " & contractCode & " ] on worksheet " & Symbols.name & "."
                End If

                Exit Sub
            End If
        Next ws
        Exit Sub
Display:
        DisplayErr Err, "DownloadPriceDataForActiveContract"
    End Sub
    Public Sub MirrorLegacyCombinedPriceTable()
Attribute MirrorLegacyCombinedPriceTable.VB_Description = "Used to mirror price data from the Legacy Database to all other databases and tables where needed."
Attribute MirrorLegacyCombinedPriceTable.VB_ProcData.VB_Invoke_Func = " \n14"
        Call HomogenizeWithLegacyCombinedPrices
    End Sub
    Private Sub HomogenizeWithLegacyCombinedPrices(Optional specificContractCode$ = vbNullString, Optional minimum_date As Date = #1/1/1970#)
    '===========================================================================================================================
    'Summary: Overwrites a given table found within a database with price data from the legacy combined table in the legacy database.
    'Inputs:  specificContractCode - If supplied then only data matching this con tract code will be edited.
    '         minimum_date - If supplied then only data after this date will be edited.
    '===========================================================================================================================
        Dim sql$, tableName$, databaseConnection As ADODB.Connection, legacy_database_path$, legacyConnection As ADODB.Connection, legacyTableName$

        Dim eReport As Variant, oiSelection As Variant, contractFilter$, connectedToSqlServer As Boolean, priceTableAvailable As Boolean

        On Error GoTo Close_Connections
        
        Set legacyConnection = GetStoredAdoClass(eLegacy).Connection
        
        If TryGetDatabaseDetails(OpenInterestEnum.FuturesAndOptions, eLegacy, legacyConnection, tableNameToReturn:=legacyTableName, databasePath:=legacy_database_path, isSqlServerDetail:=connectedToSqlServer, doesPriceTableExist:=priceTableAvailable) And Not connectedToSqlServer Then
            
            Dim legacyDateField$, legacyCodeField$, contractCodeFilter$, dateFilter$
            
            If priceTableAvailable Then legacyTableName = PriceDataTableName
            
            With GetFieldInfoForDatabaseTable(legacyConnection, legacyTableName)
                legacyDateField = .item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL
                legacyCodeField = .item("cftc_contract_market_code").DatabaseNameForSQL
            End With
            
            Dim cmd As New ADODB.command
            
            With cmd
                If LenB(specificContractCode) > 0 Then
                    .Parameters.Append .CreateParameter("@Code", adVarWChar, adParamInput, size:=Len(specificContractCode), value:=specificContractCode)
                End If
                If minimum_date <> #1/1/1970# Then
                    .Parameters.Append .CreateParameter("@Date", adDate, adParamInput, value:=minimum_date)
                End If
                .CommandType = adCmdText
            End With
            
            For Each eReport In Array(eLegacy, eDisaggregated, eTFF)
                For Each oiSelection In Array(OpenInterestEnum.FuturesAndOptions, OpenInterestEnum.FuturesOnly)

                    'If not Legacy Combined or Legacy database with price table available.
                    If Not ((eReport = eLegacy And oiSelection = OpenInterestEnum.FuturesAndOptions) Or (priceTableAvailable And eReport = eLegacy)) Then
                        
                        If databaseConnection Is Nothing Then Set databaseConnection = GetStoredAdoClass(CInt(eReport)).Connection
                        
                        If TryGetDatabaseDetails(CLng(oiSelection), CLng(eReport), databaseConnection, tableNameToReturn:=tableName, ignoreSqlServerDetails:=True, doesPriceTableExist:=priceTableAvailable) Then
                                                        
                            With databaseConnection
                                If .State = adStateClosed Then .Open
                            End With
                            
                            Dim dateField$, codeField$
                            
                            With GetFieldInfoForDatabaseTable(databaseConnection, IIf(priceTableAvailable, PriceDataTableName, tableName))
                                dateField = .item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL
                                codeField = .item("cftc_contract_market_code").DatabaseNameForSQL
                            End With
                            
                            If LenB(specificContractCode) > 0 Then contractCodeFilter = " AND L_Prices." & codeField & " = ?"
                            If minimum_date <> #1/1/1970# Then dateFilter = " AND L_Prices." & dateField & " >= ?"
                            
                            contractFilter = " WHERE NOT IsNull(L_Prices.Price)" & contractCodeFilter & dateFilter
                            
                            With cmd
                                If Not .ActiveConnection Is databaseConnection Then Set .ActiveConnection = databaseConnection
                                
                                If priceTableAvailable Then
                                    
                                    .CommandText = "UPDATE " & PriceDataTableName & " as ReportTable" & _
                                          vbNewLine & " INNER JOIN [" & legacy_database_path & "]." & PriceDataTableName & " as L_Prices" & _
                                          vbNewLine & " ON L_Prices." & legacyDateField & " = ReportTable." & dateField & " AND ReportTable." & codeField & " = L_Prices." & legacyCodeField & _
                                          vbNewLine & " SET ReportTable.Price = L_Prices.Price" & _
                                          vbNewLine & contractFilter & ";"
                                    
                                    .Execute Options:=adExecuteNoRecords Or adCmdText
    
                                    .CommandText = "INSERT INTO " & PriceDataTableName & " (" & Join(Array(dateField, codeField, "Price)"), ",") & _
                                        vbNewLine & "SELECT L_Prices." & legacyDateField & ", L_Prices." & codeField & ", L_Prices.Price FROM ((" & tableName & " as ReportTable" & _
                                        vbNewLine & "INNER JOIN [" & legacy_database_path & "]." & PriceDataTableName & " as L_Prices" & _
                                        vbNewLine & "ON ReportTable." & dateField & " = L_Prices." & legacyDateField & " AND ReportTable." & codeField & " = L_Prices." & legacyCodeField & ")" & _
                                        vbNewLine & "LEFT JOIN " & PriceDataTableName & " as ReportPrices" & _
                                        vbNewLine & "ON ReportPrices." & dateField & " = L_Prices." & legacyDateField & " AND ReportPrices." & codeField & " = L_Prices." & legacyCodeField & ")" & _
                                        vbNewLine & contractFilter & " AND ISNULL(ReportPrices." & codeField & ");"
                                Else
                                    .CommandText = "UPDATE " & tableName & _
                                        vbNewLine & "as ReportTable INNER JOIN [" & legacy_database_path & "]." & legacyTableName & " as L_Prices" & _
                                        vbNewLine & "ON (L_Prices." & legacyDateField & " = ReportTable." & dateField & " AND ReportTable." & codeField & " = L_Prices." & legacyCodeField & ")" & _
                                        vbNewLine & "SET ReportTable.[Price] = L_Prices.[Price]" & _
                                        vbNewLine & contractFilter & ";"
                                End If
                                
                                .Execute Options:=adExecuteNoRecords Or adCmdText
                            End With
                            If priceTableAvailable Then Exit For
                        End If
                    End If

                Next oiSelection
                
                Set databaseConnection = Nothing
                Set cmd.ActiveConnection = Nothing
            Next eReport
        ElseIf legacyConnection.State = adStateClosed And Not connectedToSqlServer Then
            MsgBox "Failed to connect to Legacy database."
        End If
        
Close_Connections:
        If Err.Number <> 0 Then
            'If Not cmd Is Nothing Then Debug.Print cmd.CommandText
            If databaseConnection.Errors.Count > 0 Then AppendErrorDescription Err, SummarizeAdoError(databaseConnection.Errors, Err.Number)
            PropagateError Err, "HomogenizeWithLegacyCombinedPrices"
        End If
        If Not databaseConnection Is Nothing Then Set databaseConnection = Nothing

    End Sub

    Sub Replace_All_Prices()
Attribute Replace_All_Prices.VB_Description = "For every contract code for which a price symbol is available, query new prices and upload to all available databases."
Attribute Replace_All_Prices.VB_ProcData.VB_Invoke_Func = " \n14"
    '================================================================================================================================
    'Summary: For every contract code for which a price symbol is available, query new prices and upload to all available databases.
    '================================================================================================================================
        Dim availableContractInfo As Collection, CO As ContractInfo, databaseConnection As ADODB.Connection, _
        tableName$, recordSet As ADODB.recordSet, priceRecords() As Variant, selectByContractCMD As ADODB.command, connectedToSqlServer As Boolean, doesPriceTableExist As Boolean

        Const PriceColumnIndex As Long = 3

        If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then Exit Sub

        Set databaseConnection = GetStoredAdoClass(eLegacy).Connection

        If TryGetDatabaseDetails(OpenInterestEnum.FuturesAndOptions, eLegacy, databaseConnection, tableName, isSqlServerDetail:=connectedToSqlServer, doesPriceTableExist:=doesPriceTableExist) Then

            Set availableContractInfo = GetAvailableContractInfo(True)
'            Set selectByContractCMD = CreateObject("ADODB.Command")
            Set selectByContractCMD = New ADODB.command

            On Error GoTo Close_Connection

            With databaseConnection
                If .State = adStateClosed Then .Open
'                If connectedToSqlServer Then
'                    .Execute "TRUNCATE TABLE " & PriceDataTableName & ";", Options:=adCmdText Or adExecuteNoRecords
'                ElseIf doesPriceTableExist Then
'                    .Execute "DELETE FROM " & PriceDataTableName & ";", Options:=adCmdText Or adExecuteNoRecords
'                End If
            End With
            
            With GetFieldInfoForDatabaseTable(databaseConnection, IIf(doesPriceTableExist, PriceDataTableName, tableName))
                selectByContractCMD.CommandText = "SELECT " & Join(Array(.item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL, .item("cftc_contract_market_code").DatabaseNameForSQL, "NULL as Price"), ",") & _
                " FROM " & tableName & _
                " WHERE " & .item("cftc_contract_market_code").DatabaseNameForSQL & "= ? ORDER BY " & .item("report_date_as_yyyy_mm_dd").DatabaseNameForSQL & " ASC;"
            End With
            
            With selectByContractCMD
                .Parameters.Append .CreateParameter("@ContractCode", IIf(connectedToSqlServer, adVarChar, adVarWChar), adParamInput, size:=10)
                .Prepared = True
                Set .ActiveConnection = databaseConnection
                .CommandType = adCmdText
            End With

            For Each CO In availableContractInfo
                With CO
                    If .HasSymbol Then
                        selectByContractCMD.Parameters("@ContractCode").value = .contractCode
                        With New ADODB.recordSet
                            .Open selectByContractCMD, , adOpenForwardOnly, adLockReadOnly
                            If Not .EOF And Not .BOF Then
                                priceRecords = TransposeData(.GetRows)
                                If TryGetPriceData(priceRecords, PriceColumnIndex, availableContractInfo(CO.contractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
                                    
                                    Call UpdateDatabasePricesWithArray(priceRecords, eLegacy, OpenInterestEnum.FuturesAndOptions, priceColumn:=PriceColumnIndex)
                                    If Not connectedToSqlServer Then HomogenizeWithLegacyCombinedPrices CO.contractCode
                                    
                                End If
                            End If
                            .Close
                        End With
                    End If
                End With
            Next CO
            Call RefreshAllDatabaseTables(True)
            MsgBox "Completed"
        End If
Close_Connection:
        Set selectByContractCMD = Nothing

        If Not recordSet Is Nothing Then
            With recordSet
                If .State = adStateOpen Then .Close
            End With
            Set recordSet = Nothing
        End If
        Set databaseConnection = Nothing
        If Err.Number <> 0 Then DisplayErr Err, "Replace_All_Prices"
    End Sub
    Public Sub ExchangeTableData(oiSelection As OpenInterestEnum, eReport As ReportEnum, contractCode$, maintainCurrentTableFilters As Boolean, _
    recalculateWorksheetFormulas As Boolean, _
    Optional combineSpreadWithGrossOptionsOnly As Boolean = False, Optional disableNegativeGrossOptionsOnly As Boolean = False)
    '===================================================================================================================
    'Summary: Retrieves data and updates a given listobject.
    'Inputs:
    '   destinationTable - Table to place queried data.
    '   eReport - ReportEnum used to target a database and table.
    '   oiSelection - OpenInterestEnum to query for.
    '   contractCode - Contract code to query for.
    '   maintainCurrentTableFilters = True to keep current tables found in [destinationTable].
    '   recalculateWorksheetFormulas - True to calculate formulas before exiting the subroutine.
    '===================================================================================================================
        Dim queriedData() As Variant, lastCalculatedColumn&, rawDataCountForReport&, newContractName$, _
        firstCalculatedColumn&, currentTableFilters() As Variant, currentTableDetails As LoadedData

        Dim queryDescription$, appProperties As Collection, savedState As Boolean, unitsColumnNumber&, _
        contractQuantities() As Variant, iRow&, allowPowerQuery As Boolean, destinationTable As ListObject

        Const contractNameColumnInAvailableContracts& = 2
            
        On Error GoTo Unhandled_Error_Discovered
        
        savedState = ThisWorkbook.Saved
        'allowPowerQuery = IsPowerQueryAvailable() And IsCreatorActiveUser() And DoesUserPermit_SqlServer() And oiSelection <> OptionsOnly

        #Const ProfilerEnabled = False
        '#Const ProfilerEnabled = True

        Set appProperties = DisableApplicationProperties(True, True, True)

        newContractName = WorksheetFunction.VLookup(contractCode, Available_Contracts.ListObjects(1).DataBodyRange, contractNameColumnInAvailableContracts, 0)

        queryDescription = "Query database for " & GetSqlServerTableName(eReport, oiSelection, permitOptionsOnly:=True) & " {" & contractCode & "}"

        Dim databaseQueryProfiler As TimedTask

        #If ProfilerEnabled Then
            Dim profiler As New TimedTask
            
            Const calculateFieldTask$ = "Calculations", outputToSheetTask$ = "Output to worksheet.", clearExtraCellsTask$ = "Clear extra cells beneath table"
            Const resizeTableTask$ = "Resize Table.", adjustQuantitiesTask$ = "Ensure quantity homogenity.", calculateTableTask$ = "Formula Calculation for Worksheet"
            Const GetQuantitiesTask$ = "Get quantities.", sortTask$ = "Re-Apply sort", RemoveFilterTask$ = "Remove Filters", RestoreFiltersTask$ = "Restore Filters"
            
            With profiler
                .Start "ExchangeTableData[" & newContractName & "]"
                Set databaseQueryProfiler = .StartSubTask(queryDescription)
            End With
        #End If

        With Application
            .StatusBar = "Querying database for > " & newContractName

            If allowPowerQuery Then
                queriedData = QueryForContractPQ(eReport, contractCode, oiSelection, databaseQueryProfiler)
            Else
                queriedData = QueryDatabaseForContract(eReport, oiSelection, contractCode, xlAscending, databaseQueryProfiler, combineSpreadWithGrossOptionsOnly, disableNegativeGrossOptionsOnly)
            End If

            If Not databaseQueryProfiler Is Nothing Then databaseQueryProfiler.EndTask
            .StatusBar = vbNullString
        End With
            
        If IsArrayAllocated(queriedData) Then
        
            Set currentTableDetails = GetStoredReportDetails(eReport)
            With currentTableDetails
                rawDataCountForReport = .RawDataCount.Value2
                firstCalculatedColumn = 3 + rawDataCountForReport 'Raw data coluumn count + (price) + (Empty) + (start)
                lastCalculatedColumn = .lastCalculatedColumn.Value2
            End With

            unitsColumnNumber = rawDataCountForReport - 1

            ReDim contractQuantities(LBound(queriedData, 1) To UBound(queriedData, 1), 1 To 1)
            
            ' Application.Index doesn't work because data may contain null values.
            For iRow = LBound(queriedData, 1) To UBound(queriedData, 1)
                contractQuantities(iRow, 1) = queriedData(iRow, unitsColumnNumber)
            Next iRow
            contractQuantities = GetNumbers(contractQuantities)

            ReDim Preserve queriedData(1 To UBound(queriedData, 1), 1 To lastCalculatedColumn)
            Select Case eReport
                Case eLegacy: queriedData = Legacy_Multi_Calculations(queriedData, UBound(queriedData, 1), firstCalculatedColumn, 156, 26)
                Case eDisaggregated: queriedData = Disaggregated_Multi_Calculations(queriedData, UBound(queriedData, 1), firstCalculatedColumn, 156, 26)
                Case eTFF: queriedData = TFF_Multi_Calculations(queriedData, UBound(queriedData, 1), firstCalculatedColumn, 156, 26, 52)
            End Select

            #If ProfilerEnabled Then
                profiler.StopSubTask calculateFieldTask
            #End If
            
            Set destinationTable = Get_CftcDataTable(eReport)
            
            With destinationTable
                ' This line is so that pagebreaks aren't re-calculated when removing filters.
                .Parent.DisplayPageBreaks = False
                ' You cannot write data to a filtered range so remove any currently applied filters.
                #If ProfilerEnabled Then
                    With profiler.StartSubTask(RemoveFilterTask)
                        ChangeFilters destinationTable, currentTableFilters
                        .EndTask
                    End With
                #Else
                    ChangeFilters destinationTable, currentTableFilters
                #End If

                ' Resize table first for efficiency before pasting data.
                If .ListRows.Count <> UBound(queriedData, 1) Then
                    #If ProfilerEnabled Then
                        profiler.StartSubTask resizeTableTask
                        .Resize .Range.Resize(UBound(queriedData, 1) + 1, .ListColumns.Count)
                        profiler.StopSubTask resizeTableTask
                    #Else
                        .Resize .Range.Resize(UBound(queriedData, 1) + 1, .ListColumns.Count)
                    #End If
                End If

                With .DataBodyRange
                    ' Write data to worksheet.
                    #If ProfilerEnabled Then
                        profiler.StartSubTask outputToSheetTask
                        .Resize(UBound(queriedData, 1), UBound(queriedData, 2)).Value2 = queriedData
                        profiler.StopSubTask outputToSheetTask
                    #Else
                        .Resize(UBound(queriedData, 1), UBound(queriedData, 2)).Value2 = queriedData
                    #End If

                    With .columns(1).Offset(0, -1)
                        ' Clear column that contains extracted quantities array formula.
                        If .Cells(1, 1).HasArray Then .ClearContents
                        ' Assign new quantities.
                        .Resize(UBound(queriedData, 1)).Value2 = contractQuantities
                    End With
                End With
                
                Erase queriedData: Erase contractQuantities
                
                With .Sort
                    #If ProfilerEnabled Then
                        profiler.StartSubTask sortTask
                        If .SortFields.Count > 0 Then .Apply
                        profiler.StopSubTask sortTask
                    #Else
                        If .SortFields.Count > 0 Then .Apply
                    #End If
                End With
            End With

            If maintainCurrentTableFilters And IsArrayAllocated(currentTableFilters) Then
                ' Re-Apply filters to worksheet.
                #If ProfilerEnabled Then
                    With profiler.StartSubTask(RestoreFiltersTask)
                        RestoreFilters destinationTable, currentTableFilters
                        .EndTask
                    End With
                #Else
                    RestoreFilters destinationTable, currentTableFilters
                #End If
                Erase currentTableFilters
            End If

            #If ProfilerEnabled Then
                With profiler
                    With .StartSubTask(clearExtraCellsTask)
                        ClearRegionBeneathTable destinationTable
                        .EndTask
                    End With

                    If recalculateWorksheetFormulas Then
                        With .StartSubTask(calculateTableTask)
                            destinationTable.Parent.Calculate
                            .EndTask
                        End With
                    End If
                End With
            #Else
                ClearRegionBeneathTable destinationTable
                If recalculateWorksheetFormulas Then destinationTable.Parent.Calculate
            #End If
            
            With currentTableDetails.CurrentContractName.Resize(, 4)
                ' Contract name, OI Type, Updates Pending, Contract Code
                .Value2 = Array(newContractName, oiSelection, False, contractCode)
            End With
            
            Set currentTableDetails = Nothing
            
            With MT
                'Caluclate basic dashboard worksheet if it reflects the currently active report.
                If eReport = .WorksheetReportEnum() Then
                    Dim eventState As Boolean
                                        
                    With Application
                        eventState = .EnableEvents: .EnableEvents = True
                    End With
                    .Calculate
                    Application.EnableEvents = eventState
                End If
            End With
            
            With ClientAvn
                If .Visible And eReport = eTFF Then
                    'ClientAvn defaults to using TFF data but uses Legacy data for some charts/tables.
                    If GetStoredReportDetails(eLegacy).CurrentContractCode.Value2 <> contractCode Then
                        Call ExchangeTableData(oiSelection, eLegacy, contractCode, maintainCurrentTableFilters, True)
                    End If
                    'ClientAvn.Shapes(reportAbbreviation & " Chart").ZOrder msoBringToFront
                    .Calculate
                End If
            End With
            
        End If
Finally:
        #If ProfilerEnabled Then
            profiler.DPrint
        #End If
        'Debug.Print "ExchangeTableData Re-Enabling properties : " & Now
        EnableApplicationProperties appProperties
        ThisWorkbook.Saved = savedState
        
        Exit Sub
Unhandled_Error_Discovered:
        ThisWorkbook.Saved = savedState
        With HoldError(Err)
            EnableApplicationProperties appProperties
            Application.StatusBar = vbNullString
            Call PropagateError(.HeldError, "ExchangeTableData")
        End With
    End Sub
    Public Sub RefreshAllDatabaseTables(forceUpdateTables As Boolean)
        
        Dim eReport As Variant
        On Error GoTo Display
        For Each eReport In ReportEnumArray()
            Call RefreshTableData(CLng(eReport), forceUpdateTables)
        Next
        Exit Sub
Display:
        DisplayErr Err, "RefreshAllDatabaseTables", "Failed to refresh table data for " & GetStoredReportDetails(CLng(eReport)).FullReportName.Value2
        Resume Next
    End Sub
    Public Sub RefreshTableData(ByVal eReport As ReportEnum, Optional forceUpdate As Boolean = False)
    '===================================================================================================================
    'Summary: Used to update the GUI after contracts have been updated upon activation of the calling worksheet.
    'Inputs:
    '   eReport - ReportEnum used to target a specific table.
    '===================================================================================================================
        Dim combineWithSpread As Boolean, disableNegativeGross As Boolean
        
        On Error GoTo Propagate
        With GetStoredReportDetails(eReport)
            If .PendingUpdateInDatabase.Value2 = True Or forceUpdate Then
            
                With Variable_Sheet
                    combineWithSpread = .Range("CombineSpread").Value2
                    disableNegativeGross = .Range("DisableNegativeGross").Value2
                End With
        
                Call ExchangeTableData(.OpenInterestType.Value2, eReport, .CurrentContractCode.Value2, True, True, combineWithSpread, disableNegativeGross)
            End If
        End With
        Exit Sub
Propagate:
        PropagateError Err, "RefreshTableData"
    End Sub
    Sub Latest_Contracts()
Attribute Latest_Contracts.VB_Description = "Queries the database for the latest contracts within the database loads them to the 'Available Contracts' worksheet."
Attribute Latest_Contracts.VB_ProcData.VB_Invoke_Func = " \n14"
    '======================================================================================================================
    'Summary: Queries the database for the latest contracts within the database loads them to the 'Available Contracts' worksheet.
    '======================================================================================================================
        Dim L_Table$, L_Path$, D_Path$, D_Table$, queryAvailable As Boolean, isSqlServerConn As Boolean

        Dim sqlQuery$, connectionString$, legacyAvailable As Boolean, disaggregatedAvailable As Boolean, priceTableAvailable As Boolean

        Const queryName$ = "Update Latest Contracts"

        On Error GoTo Propagate

        Dim legacyConnection As ADODB.Connection, recordSet As Object
        
        Set legacyConnection = GetStoredAdoClass(eLegacy).Connection

        legacyAvailable = TryGetDatabaseDetails(OpenInterestEnum.FuturesAndOptions, eLegacy, legacyConnection, L_Table, L_Path, isSqlServerDetail:=isSqlServerConn, doesPriceTableExist:=priceTableAvailable)
        disaggregatedAvailable = TryGetDatabaseDetails(OpenInterestEnum.FuturesOnly, eDisaggregated, , D_Table, D_Path)

        ' For why using % instead of * to match 0 or more characters see the below link.
        'https://stackoverflow.com/questions/48565908/adodb-connection-sql-not-like-query-not-working

        If legacyAvailable And disaggregatedAvailable Then

            'Query Description:
            '   Select the latest contracts in the Legacy database and join with the latest contracts in
            '   the Disaggregated database that aren't found in Legacy (ICE).
            '   Then Left join those records with disaggregated again to determine whether to assign L,T or D or L,D.
            sqlQuery = "Select contractNames.contractCode,contractNames.contractName,iif(" & IIf(isSqlServerConn, "LEN(ISNULL(Recent_Disaggregated.code,''))=0", "ISNULL(Recent_Disaggregated.code)") & ",'L,T', iif(LEFT(LCASE(Trim(contractNames.contractName)),3)= 'ice','D','L,D')) From" & _
                        "(" & _
                            "(" & _
                                "SELECT {nameField} as contractName,{codeField} as contractCode " & _
                                "From [{L_Path}].{L_Table} " & _
                                "WHERE {dateField} = {date_cutoff} " & _
                                "Union " & _
                                    "(SELECT D.{nameField} as contractName,D.{codeField} as contractCode " & _
                                    "FROM [{D_Path}].{D_Table} as D " & _
                                    "LEFT JOIN [{L_Path}].{L_Table} as L " & _
                                "ON L.{codeField}= D.{codeField} and D.{dateField}=L.{dateField} " & _
                                "WHERE " & IIf(isSqlServerConn, "LEN(ISNULL(L.{codeField},''))=0", "ISNULL(L.{codeField})") & " AND D.{dateField} = {date_cutoff})" & _
                            ") as contractNames " & _
                            "Left Join" & _
                            "(" & _
                                "Select {codeField} as code " & _
                                "From [{D_Path}].{D_Table} WHERE {dateField} = {date_cutoff}" & _
                            ") as Recent_Disaggregated " & _
                            "ON Recent_Disaggregated.code = contractNames.contractCode" & _
                        ")" & _
                        "Order by contractNames.contractName ASC;"

            Dim legacyDate As Date, disaggDate As Date

            If TryGetLatestDate(legacyDate, eLegacy, FuturesAndOptions, False, legacyConnection) And TryGetLatestDate(disaggDate, eDisaggregated, FuturesOnly, False, IIf(isSqlServerConn, legacyConnection, Nothing)) Then

                Dim dict As Object: Set dict = GetDictionaryObject()

                With dict

                    If priceTableAvailable Then
                        .item("nameField") = "market_and_exchange_names"
                        .item("dateField") = "report_date_as_yyyy_mm_dd"
                        .item("codeField") = "cftc_contract_market_code"
                    End If

                    If isSqlServerConn Then
                        sqlQuery = Replace$(sqlQuery, "LCASE", "LOWER")
                        sqlQuery = Replace$(sqlQuery, "[{L_Path}].", vbNullString)
                        sqlQuery = Replace$(sqlQuery, "[{D_Path}].", vbNullString)
                        .item("date_cutoff") = Format$(IIf(legacyDate < disaggDate, legacyDate, disaggDate), "'yyyy-mm-dd'")
                    Else
                        .item("L_Path") = L_Path
                        .item("D_Path") = D_Path
                        .item("date_cutoff") = "CDATE" & Format$(IIf(legacyDate < disaggDate, legacyDate, disaggDate), "('yyyy-mm-dd')")
                    End If

                    .item("D_Table") = D_Table
                    .item("L_Table") = L_Table
                End With

                Call Interpolator(sqlQuery, dict)

                With legacyConnection
                    If .State = adStateClosed Then .Open
                    With .Execute(sqlQuery, Options:=adCmdText)
                        On Error GoTo Close_Record
                        If Not (.BOF And .EOF) Then
                            Latest_Contracts_After_Refresh True, adodbData:=TransposeData(.GetRows)
                        End If
Close_Record:           If .State <> adStateClosed Then .Close
                    End With
                    '.Close
                End With
                If Err.Number <> 0 Then GoTo Propagate
            End If
'            Dim QT As QueryTable
'            connectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & L_Path & ";"
'
'            With Available_Contracts
'                For Each QT In .QueryTables
'                    If QT.Name Like queryName & "*" Then
'                        queryAvailable = True
'                        Exit For
'                    End If
'                Next QT
'
'                If Not queryAvailable Then
'                    Set QT = .QueryTables.Add(connectionString, .Range("K1"))
'                End If
'            End With
'
'            With QT
'                .CommandText = sqlQuery
'                .BackgroundQuery = True
'                If queryAvailable Then .Connection = connectionString
'                .CommandType = xlCmdSql
'                .MaintainConnection = False
'                .Name = queryName
'                .RefreshOnFileOpen = False
'                .RefreshStyle = xlOverwriteCells
'                .SaveData = False
'                .fieldNames = False
'
'                Set AfterEventHolder = New ClassQTE
'                AfterEventHolder.HookUpLatestContracts QT
'
'                .Refresh True
'            End With
'
        End If

        Exit Sub
Propagate:
        Set legacyConnection = Nothing
        PropagateError Err, "Latest_Contracts"

    End Sub
    Private Sub Latest_Contracts_After_Refresh(success As Boolean, Optional RefreshedQueryTable As QueryTable, Optional adodbData As Variant)
    '===================================================================================================================
    'Summary: Writes data to Available Contracts worksheet and queries API for additional contract info.
    'Inputs:
    '   success - True if COT database was successfully queried.
    '   RefreshedQueryTable - Option QueryTable that retrieved data from database.
    '   adodbData - Array used to hold data retrieved from database if a QueryTable isn't used.
    '===================================================================================================================
        Dim results() As Variant, iRow As Long, LO As ListObject, appProperties As Collection

        Const commodityColumn As Long = 4, subGroupColumn As Long = 5, codeColumn As Long = 1

        If success Then

             Set appProperties = DisableApplicationProperties(True, False, True)

             On Error GoTo Propagate

            If Not RefreshedQueryTable Is Nothing Then
                Set AfterEventHolder = Nothing
                With RefreshedQueryTable.ResultRange
                    results = .Value2
                    .ClearContents
                End With
            ElseIf IsArrayAllocated(adodbData) Then
                results = adodbData
            Else
                Exit Sub
            End If

            ReDim Preserve results(LBound(results, 1) To UBound(results, 1), 1 To UBound(results, 2) + 2)

            With CFTC_CommodityGroupings
                On Error GoTo NextResult
                For iRow = LBound(results, 1) To UBound(results, 1)
                    results(iRow, commodityColumn) = .item(results(iRow, codeColumn))(1)
                    results(iRow, subGroupColumn) = .item(results(iRow, codeColumn))(2)
NextResult:         If Err.Number <> 0 Then On Error GoTo -1
                Next iRow
            End With

            On Error GoTo Propagate
            Set LO = Available_Contracts.ListObjects("Contract_Availability")

            With LO
                With .DataBodyRange
                    .SpecialCells(xlCellTypeConstants).ClearContents
                    .Cells(1, 1).Resize(UBound(results, 1), UBound(results, 2)).Value2 = results
                End With
                .Resize .Range.Cells(1, 1).Resize(UBound(results, 1) + 1, .ListColumns.Count)
            End With

            ClearRegionBeneathTable LO
            EnableApplicationProperties appProperties
        End If

        Exit Sub
Propagate:
        With HoldError(Err)
            If success And Not appProperties Is Nothing Then EnableApplicationProperties appProperties
            PropagateError .HeldError, "Latest_Contracts_After_Refresh"
        End With
    End Sub
    Private Sub Interpolator(inputStr$, dict As Object)
    '=======================================================================================================
    ' Replaces text within {} with a value in the paramArray values.
    '=======================================================================================================
        Dim rightBrace As Long, leftSplit$(), Z As Long, D As Long, noEscapeCharacter As Boolean

        On Error GoTo Propagate

        leftSplit = Split(inputStr, "{")
        Const escapeCharacter$ = "\"

        For Z = LBound(leftSplit) To UBound(leftSplit)
            If Z > LBound(leftSplit) Then
                If Right$(leftSplit(Z), 1) = "\" Then
                    noEscapeCharacter = False
                Else
                    noEscapeCharacter = True
                End If

                If noEscapeCharacter Then
                    rightBrace = InStr(1, leftSplit(Z), "}")
                    leftSplit(Z) = dict.item(Left$(leftSplit(Z), rightBrace - 1)) & Right$(leftSplit(Z), Len(leftSplit(Z)) - rightBrace)
                    D = D + 1
                End If
            End If
        Next Z

        inputStr = Join(leftSplit, vbNullString)
        Exit Sub
Propagate:
        PropagateError Err, "Interpolator"
    End Sub
    Private Function GetFieldInfoForDatabaseTable(databaseConnection As ADODB.Connection, tableName$) As Collection
    '====================================================================================================================================
    '   Summary: Queries a database for its fields and generates a FieldInfo collection.
    '   Inputs:
    '       databaseConnection - ADODB.Connection object used to connect to the database.
    '       tableName - Name of table within database to query.
    '   Returns: A collection of FieldInfo instances.
    '====================================================================================================================================
        Dim tableField As Field, output As New Collection, standardName$, i As Long

        On Error GoTo Finally
        
        With databaseConnection
        
            If .State = adStateClosed Then .Open
            
            If DoesTableExist(databaseConnection, tableName) Then
                On Error GoTo Catch_Table404
                With .Execute(tableName, Options:=adCmdTable)
                    On Error GoTo Close_RecordSet
                    For Each tableField In .Fields
                        With tableField
                            standardName = StandardizedDatabaseFieldNames(.name)
                            i = i + 1
                            output.Add CreateFieldInfoInstance(standardName, i, .name, False, False, False, .Type), standardName
                        End With
                    Next tableField
Close_RecordSet:
                    If .State = adStateOpen Then .Close
                End With
            Else
                Err.Raise DbError.DatabaseTableMissing, Description:="Table '" & tableName & "' couldn't be found."
            End If
            
        End With
Finally:
        If Err.Number <> 0 Then PropagateError Err, "GetFieldInfoForDatabaseTable"
        Set GetFieldInfoForDatabaseTable = output
        Exit Function
Catch_Table404:
        AppendErrorDescription Err, "Failed to execute table query. Table '" & tableName & "' may not exist."
        GoTo Finally
    End Function
    Function GetDataForMultipleContractsFromDatabase(eReport As ReportEnum, versionToQuery As OpenInterestEnum, dateSortOrder As XlSortOrder, _
                            Optional maxWeeksInPast As Long = -1, Optional alternateCodes As Variant, _
                            Optional includePriceColumn As Boolean = False) As Collection
    '====================================================================================================================================
    '   Summary: Retrieves data for all favorites or select contracts from the database and stores an array for each contract keyed to its contract code.
    '   Inputs:
    '       eReport: One of L,D or T to select which database to target.
    '       versionToQuery: true if Futures + Options data is wanted; otherwise, false.
    '       sortDataAscending: true to sort data in ascending order by date otherwise false for descending.
    '       maxWeeksInPast: Number of weeks in the past in addition to the current week to query for. Use -1 to return all data
    '       alternateCodes: Specific contract codes to filter for from the database.
    '       includePriceColumn: true if you want to return prices as well.
    '   Returns: A collection of arrays keyed to that contracts contract code.
    '====================================================================================================================================
        Dim sql$, tableName$, databaseConnection As ADODB.Connection, SQL2$, availableField As FieldInfo, _
        favoritedContractCodes$, queryResult() As Variant, fieldNames$(), isPriceTableAvailable As Boolean, wantedFieldInfo As Collection, _
        contractClctn As Collection, allContracts As New Collection, oldestWantedDate As Date, mostRecentDate As Date, connectedToSqlServer As Boolean

        Dim dateColumn As Long, codeColumn As Long, nameColumn As Long, iRow As Long, iColumn As Long, queryRow() As Variant, output As New Collection

        Const dateField$ = "report_date_as_yyyy_mm_dd", codeField$ = "cftc_contract_market_code"

        On Error GoTo Finally

        If IsMissing(alternateCodes) Then
            ' Get a list of all contract codes that have been favorited.
            queryResult = WorksheetFunction.Transpose(Variable_Sheet.ListObjects("Current_Favorites").DataBodyRange.columns(1).Value2)
        Else
            queryResult = alternateCodes
        End If

        favoritedContractCodes = Join(QuotedForm(queryResult, "'"), ",")

        Set databaseConnection = GetStoredAdoClass(eReport).Connection

        If TryGetDatabaseDetails(versionToQuery, eReport, databaseConnection, tableName, , , , connectedToSqlServer, isPriceTableAvailable) Then

            With databaseConnection
                If .State = adStateClosed Then .Open
            End With

            Set wantedFieldInfo = FilterCollectionOnFieldInfoKey(GetFieldInfoForDatabaseTable(databaseConnection, tableName), GetExpectedLocalFieldInfo(eReport, True, True, includePriceColumn, True))

            ReDim fieldNames(wantedFieldInfo.Count - 1)
            iRow = LBound(fieldNames)

            For Each availableField In wantedFieldInfo
                fieldNames(iRow) = "T." & availableField.DatabaseNameForSQL
                iRow = iRow + 1
            Next

            If TryGetLatestDate(mostRecentDate, eReport, versionToQuery, False) Then

                oldestWantedDate = IIf(maxWeeksInPast > 0, DateAdd("ww", -maxWeeksInPast, mostRecentDate), DateSerial(1970, 1, 1))
                Erase queryResult

                With wantedFieldInfo
                    SQL2 = "SELECT " & .item(codeField).DatabaseNameForSQL & " FROM " & tableName & " WHERE " & .item(dateField).DatabaseNameForSQL & " = CDATE('" & Format$(mostRecentDate, "yyyy-mm-dd") & "') AND " & .item(codeField).DatabaseNameForSQL & " in (" & favoritedContractCodes & ")"

                    sql = "SELECT " & Join(fieldNames, ",") & " FROM " & tableName & " as T " & _
                    IIf(isPriceTableAvailable And includePriceColumn, " LEFT JOIN " & PriceDataTableName & " as P on P.report_date_as_yyyy_mm_dd=T." & .item(dateField).DatabaseNameForSQL & " AND P.cftc_contract_market_code=T." & .item(codeField).DatabaseNameForSQL, vbNullString) & _
                    " WHERE T." & .item(codeField).DatabaseNameForSQL & " in (" & SQL2 & ") AND T." & .item(dateField).DatabaseNameForSQL & " >=CDATE('" & oldestWantedDate & "')" & _
                    " Order BY T." & .item(codeField).DatabaseNameForSQL & " ASC,T." & .item(dateField).DatabaseNameForSQL & " " & IIf(dateSortOrder = xlAscending, "ASC;", "DESC;")
                    Erase fieldNames
                    If connectedToSqlServer Then sql = Replace$(sql, "CDATE", vbNullString)

                    codeColumn = .item(codeField).ColumnIndex
                    nameColumn = .item("market_and_exchange_names").ColumnIndex
                    dateColumn = .item(dateField).ColumnIndex
                End With

                With CreateObject("ADODB.RecordSet")
                    .cursorLocation = adUseClient
                    .Open sql, databaseConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                    queryResult = TransposeData(.GetRows())
                    .Close
                End With

                Set wantedFieldInfo = Nothing

                ReDim queryRow(LBound(queryResult, 2) To UBound(queryResult, 2))

                With allContracts
                    'Group contracts into separate collections for further processing
                    For iRow = LBound(queryResult, 1) To UBound(queryResult, 1)
                        For iColumn = LBound(queryResult, 2) To UBound(queryResult, 2)
                            queryRow(iColumn) = IIf(IsNull(queryResult(iRow, iColumn)), Empty, queryResult(iRow, iColumn))
                        Next iColumn

                        On Error GoTo Catch_CollectionMissing
                        Set contractClctn = .item(queryRow(codeColumn))
                        On Error GoTo Catch_DuplicateKeyAttempt
                        ' Use dates as a key
                        contractClctn.Add queryRow, CStr(queryRow(dateColumn))
                        On Error GoTo Finally
Next_QueryRow_Iterator:
                    Next iRow
                    Erase queryResult
                End With

                On Error GoTo Finally

                With output
                    For iRow = 1 To allContracts.Count
                        .Add CombineArraysInCollection(allContracts(iRow), Append_Type.Multiple_1d), allContracts(iRow)(1)(codeColumn)
                    Next iRow
                End With

                Set GetDataForMultipleContractsFromDatabase = output
            End If
        End If
Finally:
        Set databaseConnection = Nothing
        If Err.Number <> 0 Then PropagateError Err, "GetDataForMultipleContractsFromDatabase"
        Exit Function
        
Catch_CollectionMissing:
        Set contractClctn = New Collection
        allContracts.Add contractClctn, queryRow(codeColumn)
        Resume Next
Catch_DuplicateKeyAttempt:
        Debug.Print "Duplicate found " & queryRow(1) & " " & queryRow(nameColumn) & "   " & queryRow(codeColumn)
        Resume Next_QueryRow_Iterator
    End Function

    Public Sub Generate_Database_Dashboard(callingWorksheet As Worksheet, eReport As ReportEnum)

        Dim contractDataByCode As Collection, tempData As Variant, worksheetOutput() As Variant, totalStoch() As Variant, _
        outputRow As Long, tempRow As Long, tempCol As Long, commercialNetColumn As Long, _
        indexWeekCount As Long, Z As Long, targetColumn As Long, versionToQuery As OpenInterestEnum, sourceDates() As Date

        Dim dealerNetColumn As Long, assetNetColumn As Long, levFundNet As Long, otherNet As Long, _
        nonCommercialNetColumn As Long, totalNetColumns As Long, _
        iColumn As Variant, traderGroupDashNetColumns() As Variant, producerNet As Long, swapNet As Long, managedNet As Long, latestDate As Date

        Const threeYearsInWeeks As Long = 156, sixMonthsInWeeks As Long = 26, oneYearInWeeks As Long = 52, _
        previousWeeksToCalculate As Long = 1

        On Error GoTo No_Data

        If callingWorksheet.Shapes("FUT Only").OLEFormat.Object.value = xlOn Then
            versionToQuery = OpenInterestEnum.FuturesOnly
        Else
            versionToQuery = OpenInterestEnum.FuturesAndOptions
        End If

        Set contractDataByCode = GetDataForMultipleContractsFromDatabase(eReport, versionToQuery, xlAscending, threeYearsInWeeks + previousWeeksToCalculate + 2)

        With contractDataByCode
            If .Count = 0 Then Exit Sub
            ReDim worksheetOutput(1 To .Count, 1 To callingWorksheet.ListObjects("Dashboard_Results" & ConvertReportTypeEnum(eReport)).ListColumns.Count)
        End With

        On Error GoTo 0

        Select Case eReport
            Case eLegacy
                totalNetColumns = 2
                commercialNetColumn = UBound(contractDataByCode(1), 2) + 1
                nonCommercialNetColumn = commercialNetColumn + 1
                totalStoch = Array(3, 7, 8, commercialNetColumn, 4, 5, nonCommercialNetColumn)
                traderGroupDashNetColumns = Array(commercialNetColumn, nonCommercialNetColumn)
            Case eDisaggregated
                totalNetColumns = 4
                producerNet = UBound(contractDataByCode(1), 2) + 1
                swapNet = producerNet + 1
                managedNet = swapNet + 1
                otherNet = managedNet + 1
                totalStoch = Array(3, 4, 5, producerNet, 6, 7, swapNet, 9, 10, managedNet, 12, 13, otherNet)
                traderGroupDashNetColumns = Array(producerNet, swapNet, managedNet, otherNet)
            Case eTFF
                totalNetColumns = 4
                dealerNetColumn = UBound(contractDataByCode(1), 2) + 1
                assetNetColumn = dealerNetColumn + 1
                levFundNet = assetNetColumn + 1
                otherNet = levFundNet + 1
                totalStoch = Array(3, 4, 5, dealerNetColumn, 7, 8, assetNetColumn, 10, 11, levFundNet, 13, 14, otherNet)
                traderGroupDashNetColumns = Array(dealerNetColumn, assetNetColumn, levFundNet, otherNet)
        End Select

        For Each tempData In contractDataByCode

            contractDataByCode.Remove tempData(1, UBound(tempData, 2))

            outputRow = outputRow + 1
            'Contract name without exchange name.
            worksheetOutput(outputRow, 1) = Left$(tempData(UBound(tempData, 1), 2), InStrRev(tempData(UBound(tempData, 1), 2), "-") - 2)

            ReDim Preserve tempData(1 To UBound(tempData, 1), 1 To UBound(tempData, 2) + totalNetColumns)

            ReDim sourceDates(LBound(tempData, 1) To UBound(tempData, 1))
            'Net Position calculations.
            For tempRow = LBound(tempData, 1) To UBound(tempData, 1)

                Select Case eReport
                    Case eLegacy
                        tempData(tempRow, commercialNetColumn) = tempData(tempRow, 7) - tempData(tempRow, 8)
                        tempData(tempRow, nonCommercialNetColumn) = tempData(tempRow, 4) - tempData(tempRow, 5)
                    Case eDisaggregated
                        tempData(tempRow, producerNet) = tempData(tempRow, 4) - tempData(tempRow, 5)
                        tempData(tempRow, swapNet) = tempData(tempRow, 6) - tempData(tempRow, 7)
                        tempData(tempRow, managedNet) = tempData(tempRow, 9) - tempData(tempRow, 10)
                        tempData(tempRow, otherNet) = tempData(tempRow, 12) - tempData(tempRow, 13)
                    Case eTFF
                        tempData(tempRow, dealerNetColumn) = tempData(tempRow, 4) - tempData(tempRow, 5)
                        tempData(tempRow, assetNetColumn) = tempData(tempRow, 7) - tempData(tempRow, 8)
                        tempData(tempRow, levFundNet) = tempData(tempRow, 10) - tempData(tempRow, 11)
                        tempData(tempRow, otherNet) = tempData(tempRow, 13) - tempData(tempRow, 14)
                End Select
                sourceDates(tempRow) = tempData(tempRow, 1)
            Next tempRow
            'Index calculations using all available data.
            For Z = LBound(totalStoch) To UBound(totalStoch)
                targetColumn = totalStoch(Z)
                worksheetOutput(outputRow, 2 + Z) = Stochastic_Calculations(targetColumn, UBound(tempData, 1), tempData, previousWeeksToCalculate, sourceDates, True)(1)
            Next Z
            'Variable Index calculations.
            'tempCol is used to track the column that correlates with the given calculation.
            tempCol = 3 + UBound(totalStoch)
            For Each iColumn In traderGroupDashNetColumns
                For Z = 0 To 2
                    indexWeekCount = Array(threeYearsInWeeks, oneYearInWeeks, sixMonthsInWeeks)(Z)
                    worksheetOutput(outputRow, tempCol) = Stochastic_Calculations(CInt(iColumn), indexWeekCount, tempData, previousWeeksToCalculate, sourceDates, True)(1)
                    tempCol = tempCol + 1
                Next Z
            Next iColumn

            If tempData(UBound(tempData, 1), 1) > latestDate Then latestDate = tempData(UBound(tempData, 1), 1)

        Next tempData

        On Error GoTo 0

        With Application
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
        End With

        Dim LO As ListObject

        With callingWorksheet
            .Range("A1").Value2 = latestDate
            Set LO = .ListObjects("Dashboard_Results" & ConvertReportTypeEnum(eReport))

            With LO
                With .DataBodyRange
                    .ClearContents
                    With .Resize(UBound(worksheetOutput, 1), UBound(worksheetOutput, 2))
                        .Value2 = worksheetOutput
                        .Sort key1:=.columns(1), Orientation:=xlSortColumns, ORder1:=xlAscending, header:=xlNo, MatchCase:=False
                    End With
                End With

                If UBound(worksheetOutput, 1) <> .ListRows.Count Then
                    .Resize .Range.Resize(UBound(worksheetOutput, 1) + 1, .ListColumns.Count)
                End If
            End With
            ClearRegionBeneathTable LO
        End With

        Re_Enable

        Exit Sub
No_Data:
        DisplayErr Err, "Generate_Database_Dashboard"
    End Sub

    Public Function GetCftcWorksheet(eReport As ReportEnum, returnDataWorksheet As Boolean, getCharts As Boolean) As Worksheet

        Dim T As Long, WSA() As Variant

        If returnDataWorksheet Then
            WSA = Array(LC, DC, TC)
        ElseIf getCharts Then
            WSA = Array(L_Charts, D_Charts, T_Charts)
        Else
            Err.Raise 5, "GetCftcWorksheet", "Neither returnDataWorksheet nor getCharts is TRUE."
        End If

        On Error GoTo Catch_ReportType_Not_Found
        T = Application.Match(eReport, Array(eLegacy, eDisaggregated, eTFF), 0) - 1

        Set GetCftcWorksheet = WSA(T)

        Exit Function
Catch_ReportType_Not_Found:
        PropagateError Err, "GetCftcWorksheet", eReport & " isn't 1 of 'L,D,T'."
    End Function

    Public Function Get_CftcDataTable(eReport As ReportEnum) As ListObject
    '==================================================================================================
    '   Returns the ListObject used to store data for all data associated with the eReport paramater.
    '   Parameters: eReport - ReportEnum used to select a table.
    '==================================================================================================
        Dim LO As ListObject, tableName$
        
        tableName = ConvertReportTypeEnum(eReport) & "_Data*"
        
        With GetCftcWorksheet(eReport, True, False)
            For Each LO In .ListObjects
                If LO.name Like tableName Then
                    Set Get_CftcDataTable = LO
                    Exit Function
                End If
            Next LO
        End With
        Err.Raise DbError.ExcelTableMissing, "Get_CftcDataTable", tableName & " table not found."
    End Function

    Public Sub Save_For_Github()
Attribute Save_For_Github.VB_Description = "Marks workbook for GitHub if conditions are met."
Attribute Save_For_Github.VB_ProcData.VB_Invoke_Func = " \n14"
    '=======================================================================================================
    ' Marks workbook for GitHub if conditions are met.
    '=======================================================================================================
        On Error GoTo Display
        
        If IsCreatorActiveUser() Then
            Variable_Sheet.Range("Github_Version").Value2 = True
            Custom_SaveAS Environ("USERPROFILE") & "\Desktop\COT-GIT.xlsb"
        End If
        Exit Sub
Display:
        DisplayErr Err, "Save_For_Github"
    End Sub
    Private Sub Launch_Database_Path_Selector_Userform()
         Database_Path_Selector.Show
    End Sub
    Private Sub Adjust_Contract_Selection_Shapes()
    '=======================================================================================================
    ' Loops through worksheets containing COT data and moves the Contract_Selection launching shape to a point
    ' on the worksheet.
    '=======================================================================================================
        Dim gg As Range, ws As Variant

        For Each ws In Array(LC, DC, TC)
            Set gg = ws.Range("A1")
            With ws.Shapes("Launch Selection")
                .Top = gg.Top
                .Left = gg.Left
                .Width = gg.Width
                .Height = gg.Height
            End With
        Next ws

    End Sub
    Sub OverwritePricesAfterDate()
Attribute OverwritePricesAfterDate.VB_Description = "Overwrites all price data >= a user selected date in all available databases."
Attribute OverwritePricesAfterDate.VB_ProcData.VB_Invoke_Func = " \n14"
    '======================================================================================================
    'Summary: Overwrites all price data >= a user selected date in all available databases.
    '======================================================================================================
        Dim availableContractInfo As Collection, sql$, databaseConnection As ADODB.Connection, tableName$, queryResult() As Variant, iCount As Long, wantedCodes$

        Const dateField$ = "report_date_as_yyyy_mm_dd", _
              codeField$ = "cftc_contract_market_code", _
              nameField$ = "market_and_exchange_names"

        Dim rowIndex As Long, ColumnIndex As Long, recordsWithSameContractCode As Collection, isPriceTableAvailable As Boolean, _
        queryRow() As Variant, recordsByDateByCode As New Collection, userSelectedDate As Date, dbFields As Collection, connectedToSqlServer As Boolean

        On Error GoTo Catch_InvalidDate
        userSelectedDate = CDate(InputBox("Input date in form YYYY-MM-DD"))
        On Error GoTo DisplayError

        If MsgBox("Is this the date you want? " & vbNewLine & Format$(userSelectedDate, "[$-F800]dddd, mmmm dd, yyyy"), vbYesNo) <> vbYes Then Exit Sub

        Set databaseConnection = GetStoredAdoClass(eLegacy).Connection

        If TryGetDatabaseDetails(OpenInterestEnum.FuturesAndOptions, eLegacy, databaseConnection, tableName, isSqlServerDetail:=connectedToSqlServer, doesPriceTableExist:=isPriceTableAvailable) Then

            wantedCodes = "('" & Join(Application.Transpose(Symbols.ListObjects("Symbols_TBL").DataBodyRange.columns(1).Value2), "','") & "')"

            Const codeColumn& = 2, priceColumn& = 3

            With databaseConnection
                If .State = adStateClosed Then .Open
                
                Dim cmd As New ADODB.command, cmdRecordSet As New ADODB.recordSet
                ' Generate a command to retrieve all rows that need to be replaced.
                With GetFieldInfoForDatabaseTable(databaseConnection, tableName)
                    cmd.CommandText = "SELECT " & Join(Array(.item(dateField).DatabaseNameForSQL, .item(codeField).DatabaseNameForSQL, "null as Price"), ",") & " FROM " & tableName & _
                        vbNewLine & "WHERE " & .item(codeField).DatabaseNameForSQL & " IN " & wantedCodes & " AND " & .item(dateField).DatabaseNameForSQL & " >=?" & _
                        vbNewLine & "ORDER BY " & .item(dateField).DatabaseNameForSQL & " ASC;"
                End With
                
                With cmd
                    .CommandType = adCmdText
                    .Parameters.Append .CreateParameter("@Date", IIf(connectedToSqlServer, adDBDate, adDate), adParamInput, value:=userSelectedDate)
                    Set .ActiveConnection = databaseConnection
                End With
                                                               
                With cmdRecordSet
                    On Error GoTo Catch_FailedRetrievalFromDatabase
                    .Open cmd, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly
                    On Error GoTo DisplayError
                    If Not .EOF Then queryResult = TransposeData(.GetRows())
                    .Close
                End With
                
            End With
            
            Set databaseConnection = Nothing
            
            If IsArrayAllocated(queryResult) Then
                
                Dim contractCode$
                ReDim queryRow(LBound(queryResult, 2) To UBound(queryResult, 2))

                With recordsByDateByCode
                    'Group contracts into separate collections for further processing
                    For rowIndex = LBound(queryResult, 1) To UBound(queryResult, 1)
                    
                        For ColumnIndex = LBound(queryResult, 2) To UBound(queryResult, 2)
                            queryRow(ColumnIndex) = queryResult(rowIndex, ColumnIndex)
                        Next ColumnIndex
                        
                        contractCode = queryRow(codeColumn)
                        
                        On Error GoTo Create_Contract_Collection
                        Set recordsWithSameContractCode = .item(contractCode)
                                                
                        On Error Resume Next ' Key using dates.
                        recordsWithSameContractCode.Add queryRow, CStr(queryRow(LBound(queryRow)))
                        On Error GoTo DisplayError
                        
                    Next rowIndex
                    
                End With

                Erase queryResult
                Erase queryRow
                Set availableContractInfo = GetAvailableContractInfo(True)
                
                With recordsByDateByCode
                    For iCount = .Count To 1 Step -1
                        Set recordsWithSameContractCode = .item(iCount)
                        queryResult = CombineArraysInCollection(recordsWithSameContractCode, Append_Type.Multiple_1d)
                        contractCode = queryResult(LBound(queryResult, 1), codeColumn)
                        
                        .Remove contractCode
                        
                        If HasKey(availableContractInfo, contractCode) Then
                            ' If price data can be retrieved then re-add to collection.
                            If TryGetPriceData(queryResult, priceColumn, availableContractInfo(contractCode), True, True) Then
                                .Add queryResult, contractCode
                            Else
                                Debug.Print "Couldn't retrieve price data for " & contractCode
                            End If
                        End If
                    Next iCount
                End With

                If recordsByDateByCode.Count > 0 Then
                    queryResult = CombineArraysInCollection(recordsByDateByCode, Append_Type.Multiple_2d)
                    
                    UpdateDatabasePricesWithArray queryResult, eLegacy, True, priceColumn
                    
                    If Not connectedToSqlServer Then HomogenizeWithLegacyCombinedPrices minimum_date:=userSelectedDate
                    
                    Call RefreshAllDatabaseTables(True)
                End If
            Else
                MsgBox "No records could be retrieved from the Legacy database matching the current parameters could be retrieved."
            End If
        End If

        Exit Sub
Catch_InvalidDate:
        Exit Sub
Create_Contract_Collection:
        Set recordsWithSameContractCode = New Collection
        recordsByDateByCode.Add recordsWithSameContractCode, contractCode
        Resume Next
DisplayError:
        'Stop: Resume
        DisplayErr Err, "OverwritePricesAfterDate"
        Exit Sub
Catch_FailedRetrievalFromDatabase:
        AppendErrorDescription Err, "Failed to execute commnad."
        GoTo DisplayError
    End Sub

    Public Sub FindDatabasePathInSameFolder(Optional hideFromPublic As Boolean)
    '===========================================================================================================
    ' Looks for MS Access Database files that haven't been renamed within the same folder as the Excel workbook.
    '===========================================================================================================
        Dim legacy As New LoadedData, TFF As New LoadedData, DGG As New LoadedData, _
        databaseFileName$, databasePathRange As Range
        
        #If Not Mac Then
            If Not DoesUserPermit_SqlServer() Then
            
                On Error GoTo Propagate
                ' Initializing these classes will wipe database paths if they cannot be found.
                With legacy
                    .InitializeClass eLegacy, True
                    With .CurrentDatabasePath.ListObject.DataBodyRange
                        Set databasePathRange = .columns(legacy.CurrentDatabasePath.Column - .Column + 1)
                    End With
                End With
                DGG.InitializeClass eDisaggregated, True
                TFF.InitializeClass eTFF, True
                
                If WorksheetFunction.CountA(databasePathRange) <> databasePathRange.Rows.Count Then
                        
                    Dim FSO As New Scripting.FileSystemObject, myFolder As Folder, mySubFolder As Folder, _
                    myFile As file, queue As New Collection, foundSubFoldersLevelOne As Boolean, checksize As Boolean
                
                    On Error GoTo ManualCalibrationRequired
                    queue.Add FSO.GetFolder(ThisWorkbook.path)
                    
                    ' FIFO check folder containing this file and 1 level of sub-folders.
                    Do While queue.Count > 0
                    
                        Set myFolder = queue(1)
                        queue.Remove 1
                        
                        If Not foundSubFoldersLevelOne Then
                            For Each mySubFolder In myFolder.SubFolders
                                queue.Add mySubFolder
                            Next mySubFolder
                            foundSubFoldersLevelOne = True
                        End If
                        
                        For Each myFile In myFolder.Files
                            With myFile
                                databaseFileName = LCase$(.name)
                                If databaseFileName Like "*.accdb" And Not (.path Like "*[[]*" Or .path Like "*[]]*") Then
                                
                                    If databaseFileName Like "*disaggregated.accdb" Then
                                        With DGG.CurrentDatabasePath
                                            If LenB(.Value2) = 0 Then .Value2 = .path: checksize = True
                                        End With
                                    ElseIf databaseFileName Like "*legacy.accdb" Then
                                        With legacy.CurrentDatabasePath
                                            If LenB(.Value2) = 0 Then .Value2 = .path: checksize = True
                                        End With
                                    ElseIf databaseFileName Like "*tff.accdb" Then
                                        With TFF.CurrentDatabasePath
                                            If LenB(.Value2) = 0 Then .Value2 = .path: checksize = True
                                        End With
                                    End If
                                    If checksize Then CheckMsAccessSize .path: checksize = False
                                End If
                            End With
                        Next myFile
                        
                    Loop
ManualCalibrationRequired:
                    If Err.Number <> 0 Then On Error GoTo -1
                    On Error GoTo Propagate
        
                    If WorksheetFunction.CountA(databasePathRange) <> databasePathRange.Rows.Count Then
                        ' If any database path is still empty then display a message.
                        MsgBox "Database paths couldn't be auto-retrieved." & String$(2, vbNewLine) & _
                        "Please use the Database Paths Userform on the [ " & HUB.name & " ] worksheet to fill in the needed data."
                        
                        Err.Raise 17, Description:="1 or more missing databases."
                    End If
                    
                End If
            End If
        #End If
        
        Exit Sub
Propagate:
        PropagateError Err, "FindDatabasePathInSameFolder"
    End Sub
    Public Function GetStoredReportDetails(reportType As ReportEnum) As LoadedData
    '===================================================================================================================
    'Summary: Loads relevant details for the report indicated by [reportType] into a class
    'Parameters:
    '   reportType - An enum used to determine which report to gather data for.
    'Returns:
    '   A LoadedData class object is returned.
    '===================================================================================================================
        Dim storedData As New LoadedData
        storedData.InitializeClass reportType
        Set GetStoredReportDetails = storedData

    End Function

    Public Function GetContractInfo_DbVersion(Optional includeAllContractsWithSymbol As Boolean = False) As Collection
    '==============================================================================================
    ' Creates a collection of Contract objects keyed to their contract code for each
    ' available contract within the database.
    '==============================================================================================

        Dim ContractList() As Variant, CD As ContractInfo, iRow As Long, _
        pAllContracts As New Collection, PriceSymbol$, usingYahoo As Boolean, availableSymbols() As Variant
        
        #Const EngageTimer = False
        
        #If EngageTimer Then
            Dim profiler As New TimedTask
            profiler.Start "GetContractInfo_DbVersion"
        #End If
        
        On Error GoTo Propagate
        ' Get array of latest contracts and supplemental info.
        ContractList = Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.Value2

        Const codeColumn As Long = 1, nameColumn As Long = 2, availabileColumn As Long = 3, _
        commodityGroupColumn As Long = 4, subGroupColumn As Long = 5, hasSymbolColumn As Long = 6, isFavoriteColumn As Long = 7

        availableSymbols = Symbols.ListObjects("Symbols_TBL").DataBodyRange.Value2

        For iRow = LBound(ContractList) To UBound(ContractList)
            PriceSymbol = vbNullString
            usingYahoo = False

            If ContractList(iRow, hasSymbolColumn) = True Then
                On Error GoTo Catch_SymbolNotFound
                PriceSymbol = WorksheetFunction.VLookup(ContractList(iRow, codeColumn), availableSymbols, 3, False)
                On Error GoTo Propagate
                usingYahoo = LenB(PriceSymbol) <> 0
            End If

            Set CD = New ContractInfo

            CD.InitializeBasicVersion CStr(ContractList(iRow, codeColumn)), CStr(ContractList(iRow, nameColumn)), CStr(ContractList(iRow, availabileColumn)), CBool(ContractList(iRow, isFavoriteColumn)), PriceSymbol, usingYahoo
            On Error GoTo Possible_Duplicate_Key
            pAllContracts.Add CD, ContractList(iRow, codeColumn)
            On Error GoTo Propagate
        Next iRow

        If includeAllContractsWithSymbol Then
            With pAllContracts
                For iRow = LBound(availableSymbols, 1) To UBound(availableSymbols, 1)
                    If Not HasKey(pAllContracts, CStr(availableSymbols(iRow, codeColumn))) And LenB(availableSymbols(iRow, 3)) <> 0 Then
                        Set CD = New ContractInfo
                        CD.InitializeBasicVersion CStr(availableSymbols(iRow, codeColumn)), "Na", "L", False, CStr(availableSymbols(iRow, 3)), True
                        .Add CD, availableSymbols(iRow, codeColumn)
                    End If
                Next iRow
            End With
        End If

        Set GetContractInfo_DbVersion = pAllContracts
        
        #If EngageTimer Then
            With profiler
                .EndTask
                .DPrint
            End With
        #End If
        
        Exit Function

Possible_Duplicate_Key:
        Resume Next
Catch_SymbolNotFound:
        'priceSymbol = Right$(String$(6, "0") & contractList(iRow, codeColumn), 6)
        Resume Next
Propagate:
        Call PropagateError(Err, "GetContractInfo_DbVersion")
    End Function

    Public Sub DeactivateContractSelection(Optional hideFromUser As Boolean = True)
        If IsLoadedUserform("Contract_Selection") Then
           Unload Contract_Selection
        End If
    End Sub

    Public Sub Open_Contract_Selection()
Attribute Open_Contract_Selection.VB_Description = "Opens the contract selection userform if conditions are met."
Attribute Open_Contract_Selection.VB_ProcData.VB_Invoke_Func = " \n14"
        Dim reportToLoad As ReportEnum

        On Error GoTo Failed_To_Get_Type
            With ThisWorkbook
                reportToLoad = .Worksheets(.ActiveSheet.name).WorksheetReportEnum
            End With
        On Error GoTo 0

        With Contract_Selection
            .SetReport reportToLoad
            .Show
        End With
Finally:
        Exit Sub

Failed_To_Get_Type:
        MsgBox ThisWorkbook.ActiveSheet.name & " does not have a publicly available WorksheetReportEnum Function."
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

        Dim contractData As Variant, outputA() As Variant, contractDataByCode As Collection, iRow As Long, mostRecentContractCodes As Variant

        Dim availableContracts As Collection, currentWeekNet As Long, previousWeekNet As Long

        Const maxWeeksToReturn As Long = 52, weekCountOfShift As Long = 1
        On Error GoTo ShowError
        mostRecentContractCodes = Application.Transpose(Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.columns(1).Value2)

        Set contractDataByCode = GetDataForMultipleContractsFromDatabase(eLegacy, OpenInterestEnum.FuturesOnly, xlAscending, maxWeeksToReturn - 1, mostRecentContractCodes)

        If Not contractDataByCode Is Nothing Then

            Dim commLong As Long, commShort As Long, nonCommLong As Long, nonCommShort As Long, codeColumn As Long, _
            iColumn As Long, columnLong As Long, columnShort As Long, oiColumn As Long ', clusteringAndConcentration()

            With GetExpectedLocalFieldInfo(eLegacy, True, True, False, True)
                commLong = .item("comm_positions_long_all").ColumnIndex
                commShort = .item("comm_positions_short_all").ColumnIndex
                nonCommLong = .item("noncomm_positions_long_all").ColumnIndex
                nonCommShort = .item("noncomm_positions_short_all").ColumnIndex
                codeColumn = .item("cftc_contract_market_code").ColumnIndex
                oiColumn = .item("oi_all").ColumnIndex

                'ReDim clusteringAndConcentration(1 To UBound(outputA, 1), 1 To 5)

                Dim recentDate As Date, commConcLong As Long, commConcShort As Long, nonCommConcShort As Long, nonCommConcLong As Long, traderCount As Long
                Dim longTraders As Long, shortTraders As Long, clustering() As Double, iCountCluster As Long, dateColumn As Long

                commConcLong = .item("pct_of_oi_comm_long_all").ColumnIndex
                commConcShort = .item("pct_of_oi_comm_short_all").ColumnIndex
                nonCommConcLong = .item("pct_of_oi_noncomm_long_all").ColumnIndex
                nonCommConcShort = .item("pct_of_oi_noncomm_short_all").ColumnIndex
                traderCount = .item("traders_tot_all").ColumnIndex

                longTraders = .item("traders_noncomm_long_all").ColumnIndex
                shortTraders = .item("traders_noncomm_short_all").ColumnIndex
                dateColumn = .item("report_date_as_yyyy_mm_DD").ColumnIndex
            End With

            recentDate = Variable_Sheet.Range("Last_Updated_CFTC").Value2
            ReDim outputA(1 To contractDataByCode.Count, 1 To 12)

            Set availableContracts = GetAvailableContractInfo
            Dim currentWeek As Long, comparisonWeek As Long, sourceDates() As Date

            For Each contractData In contractDataByCode

                currentWeek = UBound(contractData, 1):

                On Error GoTo Next_ContracData

                If UBound(contractData, 1) >= 2 And contractData(currentWeek, dateColumn) = recentDate Then

                    comparisonWeek = currentWeek - (weekCountOfShift)
                    iRow = iRow + 1
                    outputA(iRow, 1) = contractData(currentWeek, codeColumn)

                    On Error GoTo Catch_CodeMissing
                        outputA(iRow, 2) = availableContracts(contractData(currentWeek, codeColumn)).ContractNameWithoutMarket
                    On Error GoTo ShowError

                    ReDim clustering(LBound(contractData, 1) To UBound(contractData, 1), 1 To 2)
                    ReDim sourceDates(LBound(contractData, 1) To UBound(contractData, 1))

                    For iCountCluster = LBound(contractData, 1) To UBound(contractData, 1)
                        'Longs
                        clustering(iCountCluster, 1) = contractData(iCountCluster, longTraders) / contractData(iCountCluster, traderCount)
                        'Shorts
                        clustering(iCountCluster, 2) = contractData(iCountCluster, shortTraders) / contractData(iCountCluster, traderCount)
                        sourceDates(iCountCluster) = contractData(iCountCluster, dateColumn)
                    Next iCountCluster

                    outputA(iRow, 7) = Stochastic_Calculations(CLng(nonCommConcLong), UBound(clustering, 1), contractData, 1, sourceDates, True, dateColumn)(1)
                    'Long clustering
                    outputA(iRow, 8) = Stochastic_Calculations(1, UBound(clustering, 1), clustering, 1, sourceDates, True, dateColumn)(1)
                    outputA(iRow, 9) = Stochastic_Calculations(CLng(nonCommConcShort), UBound(clustering, 1), contractData, 1, sourceDates, True, dateColumn)(1)
                    'clustering
                    outputA(iRow, 10) = Stochastic_Calculations(2, UBound(clustering, 1), clustering, 1, sourceDates, True, dateColumn)(1)

                    ' Loop positions for both groups.
                    For iColumn = 0 To 3

                        Dim columnTarget As Long

                        columnTarget = Array(nonCommLong, nonCommShort, commLong, commShort)(iColumn)

                        If comparisonWeek >= LBound(contractData, 1) Then
                            If contractData(comparisonWeek, columnTarget) <> 0 Then
                                ' Calculate % change for longs and shorts.
                                outputA(iRow, 3 + iColumn) = 100 * ((contractData(currentWeek, columnTarget) - contractData(comparisonWeek, columnTarget)) / contractData(comparisonWeek, columnTarget))
                            End If
                        End If

                        If iColumn Mod 2 = 0 Then

                            currentWeekNet = contractData(currentWeek, columnTarget) - contractData(currentWeek, columnTarget + 1)

                            If comparisonWeek >= LBound(contractData, 1) Then
                                previousWeekNet = contractData(comparisonWeek, columnTarget) - contractData(comparisonWeek, columnTarget + 1)
                            End If

                            Dim calc As Long: calc = 0

                            If columnTarget = nonCommLong Then
                                calc = 11
                            ElseIf columnTarget = commLong Then
                                calc = 12
                            End If

                            If calc <> 0 Then
                                On Error Resume Next
                                'Commercial net change from previouse / (previous longs+shorts)
                                outputA(iRow, calc) = (currentWeekNet - previousWeekNet) / (contractData(comparisonWeek, columnTarget) + contractData(comparisonWeek, columnTarget + 1))

                                '% difference in net position.
                                'outputA(iRow, calc) = (currentWeekNet - previousWeekNet) / contractData(comparisonWeek, 3)

                                If Err.Number <> 0 Then
                                    outputA(iRow, calc) = 0
                                    Err.Clear
                                End If
                                On Error GoTo ShowError
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
Next_ContracData:
                On Error GoTo -1
            Next contractData

            On Error GoTo 0

            Dim tableDataRng As Range, LO As ListObject, currentFilters As Variant, appProperties As Collection

            Set LO = WeeklyChanges.ListObjects("PctNetChange")
            Set tableDataRng = LO.DataBodyRange

            With tableDataRng

                Set appProperties = DisableApplicationProperties(True, False, True)

                ChangeFilters LO, currentFilters
                On Error Resume Next
                    .SpecialCells(xlCellTypeConstants).ClearContents
                On Error GoTo ShowError

                .columns(4).Resize(UBound(outputA, 1), UBound(outputA, 2)).Value2 = outputA

                ResizeTableBasedOnColumn LO, .columns(4)

                ClearRegionBeneathTable LO
                With LO.Sort
                    'With .SortFields
                        '.Clear
                        'Group
                        '.Add tableDataRng.columns(2), xlSortOnValues, xlAscending
                        'SubGroup
                        '.Add tableDataRng.columns(3), xlSortOnValues, xlAscending
                        'Name
                        '.Add tableDataRng.columns(5), xlSortOnValues, xlAscending
                        'Rank
                        '.Add tableDataRng.columns(12), xlSortOnValues, xlAscending
                    'End With
                    If .SortFields.Count > 0 Then .Apply
                End With
                RestoreFilters LO, currentFilters

                WeeklyChanges.Range("reflectedDate").Value2 = Variable_Sheet.Range("Last_Updated_CFTC").Value2
                EnableApplicationProperties appProperties

                '=SUM(IF(SUBTOTAL(103,OFFSET([Commercial Net change/Total Position],ROW([Commercial Net change/Total Position])-ROW($A$3),0,1))>0,IF(K10<[Commercial Net change/Total Position],1)))+1
            End With
        Else
            MsgBox "Database Unavailable"
        End If

        Exit Sub
ShowError:
        DisplayErr Err, "CFTC_CalculateWeeklyChanges"
        Exit Sub
Catch_CodeMissing:
        Resume Next_ContracData
    End Sub

    Private Sub AttemptCross()

        Dim dataFieldInfoByEditedName As Collection, notionalValue() As Double, iRow As Long, _
        dataFromDatabase As Collection, reportType As ReportEnum, notionalValuesByCode As New Collection, _
        code As Variant, contractUnits As Variant, prices As Variant

        On Error GoTo Finally

        ' Setting equal to -1 will allow all data to be retrieved.
        Const maxWeeksInPast As Long = -1, versionToQuery As Long = OpenInterestEnum.FuturesOnly

        IncreasePerformance

        Dim codeToLong$, codeToShort$, wantedContractCodes As Variant

        With ForexCross
            .Visible = True

            wantedContractCodes = .ListObjects("Long_Short").DataBodyRange.Value2

            With .ListObjects("ForexTickers")
                codeToLong = WorksheetFunction.VLookup(wantedContractCodes(1, 1), .DataBodyRange, 2, False)
                codeToShort = WorksheetFunction.VLookup(wantedContractCodes(1, 2), .DataBodyRange, 2, False)
            End With

            If LenB(codeToLong) = 0 Or LenB(codeToShort) = 0 Or codeToLong = codeToShort Then
                MsgBox "Invalid input paramaters."
                Exit Sub
            End If

            wantedContractCodes = Array(codeToLong, codeToShort)
        End With

        reportType = eLegacy

        Set dataFromDatabase = GetDataForMultipleContractsFromDatabase(reportType, versionToQuery, xlAscending, maxWeeksInPast, wantedContractCodes, True)

        Set dataFieldInfoByEditedName = GetExpectedLocalFieldInfo(reportType, True, True, False)

'        For Each Code In wantedContractCodes
'
'            contractUnits = WorksheetFunction.index(dataFromDatabase(Code), 0, dataFieldInfoByEditedName("contract_units").ColumnIndex)  '
'            contractUnits = GetNumbers(contractUnits)
'
'            With notionalValuesByCode
'                .Add New Collection, Code
'                With .Item(Code)
'                    prices = Application.index(dataFromDatabase(Code), 0, dataFieldInfoByEditedName("price").ColumnIndex)
'
'                    ReDim notionalValue(LBound(contractUnits, 1) To UBound(contractUnits, 1))
'
'                    For iRow = LBound(contractUnits, 1) To UBound(contractUnits, 1)
'                        If Not IsEmpty(prices(iRow, 1)) Then
'                            notionalValue(iRow) = prices(iRow, 1) * contractUnits(iRow, 1)
'                        End If
'                    Next
'
'                    .Add notionalValue, "Notional"
'                End With
'            End With
'        Next

        'Calculate hedge ratio and combine.

        Dim contractToLong() As Variant, contractToShort() As Variant, iColumn As Long, _
        hedgeRatio As Double, output() As Variant, nonCommLong As Long, commLong As Long, commShort As Long, _
        nonCommShort As Long, iShortRow As Long, iReduction As Long

        contractToLong = dataFromDatabase(codeToLong)
        contractToShort = dataFromDatabase(codeToShort)

    '    commLong = dataFieldInfoByEditedName("comm_positions_long_all").ColumnIndex
    '    commShort = dataFieldInfoByEditedName("comm_positions_short_all").ColumnIndex
    '    nonCommLong = dataFieldInfoByEditedName("noncomm_positions_long_all").ColumnIndex
    '    nonCommShort = dataFieldInfoByEditedName("noncomm_positions_short_all").ColumnIndex

        commLong = dataFieldInfoByEditedName("pct_of_oi_comm_long_all").ColumnIndex
        commShort = dataFieldInfoByEditedName("pct_of_oi_comm_short_all").ColumnIndex

        nonCommLong = dataFieldInfoByEditedName("pct_of_oi_noncomm_long_all").ColumnIndex
        nonCommShort = dataFieldInfoByEditedName("pct_of_oi_noncomm_short_all").ColumnIndex

        ReDim output(1 To UBound(contractToLong, 1), 1 To 5)

        iShortRow = UBound(contractToShort, 1)

        On Error GoTo Exit_Loop

        For iRow = UBound(contractToLong, 1) To LBound(contractToLong, 1) Step -1

            If contractToLong(iRow, 1) = contractToShort(iShortRow, 1) Then
'--------------------------------------------------------------------------------------------------------------------------------------------
                'hedgeRatio = notionalValuesByCode(codeToLong)("Notional")(iRow) / notionalValuesByCode(codeToShort)("Notional")(iShortRow)
                'hedgeRatio = notionalValuesByCode(codeToShort)("Notional")(iShortRow) / notionalValuesByCode(codeToLong)("Notional")(iRow)

    '            For iColumn = LBound(output, 2) + 2 To UBound(output, 2)
    '
    '                If InStrB(1, dataFieldInfoByEditedName(iColumn).EditedName, "spread") = 0 Then
    '
    '                    If InStrB(1, dataFieldInfoByEditedName(iColumn).EditedName, "comm") = 1 Then
    '                        nonCommLong = dataFieldInfoByEditedName("comm_positions_long_all").ColumnIndex
    '                        nonCommShort = dataFieldInfoByEditedName("comm_positions_short_all").ColumnIndex
    '                    Else
    '                        nonCommLong = dataFieldInfoByEditedName("noncomm_positions_long_all").ColumnIndex
    '                        nonCommShort = dataFieldInfoByEditedName("noncomm_positions_short_all").ColumnIndex
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
'-----------------------------------------------------------------------------------------------------------------------------------------
                ' NET OI of contract to long NonComm
                output(iRow, 2) = contractToLong(iRow, nonCommLong) - contractToLong(iRow, nonCommShort)
                ' NET OI of  contract to short NonComm
                output(iRow, 3) = contractToShort(iShortRow, nonCommLong) - contractToShort(iShortRow, nonCommShort)
                ' NET OI of  contract to long Comm
                output(iRow, 4) = contractToLong(iRow, commLong) - contractToLong(iRow, commShort)
                ' NET OI of  contract to Short Comm
                output(iRow, 5) = contractToShort(iShortRow, commLong) - contractToShort(iShortRow, commShort)
'------------------------------------------------------------------------------------------------------------------------------------------

                'output(iRow, 2) = (contractToLong(iRow, nonCommLong) - contractToShort(iRow, nonCommLong))
                'Dates
                output(iRow, 1) = contractToLong(iRow, 1)

            End If

            iShortRow = iShortRow - 1

        Next iRow

PlaceOnSheet:
        On Error GoTo 0

        Dim tableFilters() As Variant, outputTable As ListObject

        Set outputTable = ForexCross.ListObjects("CrossTable")

        With outputTable.DataBodyRange
            ChangeFilters outputTable, tableFilters
            .Range(.Cells(1, 1), .Cells(.Rows.Count, UBound(output, 2))).ClearContents
            .Resize(UBound(output, 1), UBound(output, 2)).Value2 = Reverse_2D_Array(output)
            ResizeTableBasedOnColumn outputTable, .columns(1)
            ClearRegionBeneathTable outputTable
            RestoreFilters outputTable, tableFilters
        End With
        
        With ForexCross
            .Range("ActiveContract").Value2 = Join(Application.Transpose(Application.Transpose(.ListObjects("Long_Short").DataBodyRange.Value2)), "/")
        End With
Finally:
        DisplayErr Err, "AttemptCross"
        Re_Enable
        Exit Sub
Exit_Loop:
        Resume PlaceOnSheet
    End Sub
    '
    '
    '
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Function DoesUserPermit_SqlServer() As Boolean
    '===================================================================================================================================================
    'Summary: Returns a boolean to determine if SQL Server connections should be allowed.
    'Returns: True if a server name is provided and SQL Server has been permitted.
    '===================================================================================================================================================
        On Error GoTo Propagate
        Const FormulaToEvaluate$ = "=AND(NOT(ISBLANK(SqlServerName)),SQL_Server_Allowed)"
        DoesUserPermit_SqlServer = Evaluate(FormulaToEvaluate)
        Exit Function
Propagate:
        PropagateError Err, "DoesUserPermit_SqlServer", "Failed to evaluate " & FormulaToEvaluate
    End Function

    Private Function GetSqlServerConnectionString(databaseExistanceKnown As Boolean) As String

        Dim connectionString$, serverName$

        serverName = Variable_Sheet.Range("SqlServerName").Value2

        If LenB(serverName) > 0 Then
            connectionString = "Provider=MSOLEDBSQL;Data Source=" & serverName & ";Trusted_Connection=Yes;TrustServerCertificate=True;MultipleActiveResultSets=True;Connection Timeout=5"
            If databaseExistanceKnown Then
                connectionString = connectionString & ";Database=" & SqlServerDatabaseName
            End If
            GetSqlServerConnectionString = connectionString
        Else
            Err.Raise DbError.ServerNameMissing, "GetSqlServerConnectionString", "SQL Server name cannot be determined."
        End If

    End Function

    Private Function TryConnectingToSqlServer(closeConnectionIfSuccess As Boolean, Optional databaseConnection As ADODB.Connection, Optional connectToDatabaseCOT As Boolean = False, Optional eReport As ReportEnum, Optional oiType As OpenInterestEnum, Optional tableNameToReturn$) As Boolean
    '===================================================================================================================================================
    'Summary: Attempts to connecto to a SQL Server database and generates a database and report table if needed.
    'Parameters:
    '   connectToDatabaseCOT - If set to true, then if possible then the current catalog will be changed to point to the COT Database.
    '   eReport - A ReportEnum used to select for Legacy,Disaggregated or TFF tables.
    '   oiType - Used to select for a Futures Combined or Futures Only table.
    '   closeConnectionIfSuccess - If true and [databaseConnection] has been supplied and the connection succeds then the conection will be closd on exit.
    '   databaseConnection - ADODB.Connection object that will connect to the server if available
    'Returns: True if connection to the server suceeds.
    '===================================================================================================================================================
        Dim conn As ADODB.Connection, succesfullyCompleted As Boolean
        
        If databaseConnection Is Nothing Then
            Set conn = New ADODB.Connection
        Else
            Set conn = databaseConnection
        End If

        On Error GoTo Finally

        With conn

            .Open GetSqlServerConnectionString(databaseExistanceKnown:=False)
            ' No error generated ergo SQL Server exists.
            If Not SQLSERVER_DatabaseExists_COT Then
                Ensure_SQLSERVER_DatabaseExists conn
                SQLSERVER_DatabaseExists_COT = True
            End If

            If connectToDatabaseCOT Then
                .Properties("Current Catalog").value = SqlServerDatabaseName
                tableNameToReturn = GetSqlServerTableName(eReport, oiType)
                
                If oiType <> OpenInterestEnum.OptionsOnly And Not IsReportTableExistanceConfirmed(tableNameToReturn) Then
                    Ensure_SQLSERVER_ReportTableExists eReport, oiType, conn, tableNameToReturn
                    SQL_Server_TableExistance.item(tableNameToReturn) = Null
                End If
            End If

            succesfullyCompleted = True
Finally:
            If databaseConnection Is Nothing Or (closeConnectionIfSuccess And succesfullyCompleted) Or Err.Number <> 0 Then
                If .State <> adStateClosed Then .Close
            End If

            Select Case Err.Number
                Case 0
                Case DbError.InvalidConnectionStringSqlServer
                    PropagateError Err, "TryConnectingToSqlServer", "SQL Server may be shutdown or unavailable."
                Case Else
                    PropagateError Err, "TryConnectingToSqlServer"
            End Select

        End With

        TryConnectingToSqlServer = succesfullyCompleted

    End Function
    Private Function IsReportTableExistanceConfirmed(tableName$) As Boolean
    
        If SQL_Server_TableExistance Is Nothing Then
            Set SQL_Server_TableExistance = GetDictionaryObject()
        Else
            IsReportTableExistanceConfirmed = SQL_Server_TableExistance.Exists(tableName)
        End If
    
    End Function
    Private Function GetSqlServerTableName(eReport As ReportEnum, oiType As OpenInterestEnum, Optional appendDatabaseName As Boolean = False, Optional permitOptionsOnly As Boolean = False) As String

        Dim tableName$
        
        On Error GoTo Propagate
        
        Select Case oiType
            Case FuturesOnly
                tableName = ConvertReportTypeEnum(eReport) & "_Futures_Only"
            Case FuturesAndOptions
                tableName = ConvertReportTypeEnum(eReport) & "_Combined"
            Case Else
                If Not permitOptionsOnly Then Err.Raise DbError.VersionUnacceptable, "Invalid open interest selection."
                tableName = ConvertReportTypeEnum(eReport) & "_Options_only"
        End Select

        If appendDatabaseName Then
            GetSqlServerTableName = SqlServerDatabaseName & "." & tableName
        Else
            GetSqlServerTableName = tableName
        End If
        Exit Function
Propagate:
        PropagateError Err, "GetSqlServerTableName"
    End Function
    Private Function IsSqlServerConnection(conn As ADODB.Connection) As Boolean
    '===================================================================================================================
    'Summary: Determines if a given ADODB.Connection [conn] is configured for SQL Server access.
    'Parameters:
    '   conn - ADODB.Connection to test the properties of.
    'Returns: True if [conn] is configured for SQL Server; otherwise, False.
    'Remarks: Testing configuration on an unconfigured Connection oject will lead to errors in its further usage.
    '===================================================================================================================
        On Error GoTo ExitSUB
        If Not conn Is Nothing Then
            With conn
                If LenB(.connectionString) > 0 Then
                    IsSqlServerConnection = .Properties("DBMS Name").value = "Microsoft SQL Server"
                End If
            End With
        End If
ExitSUB:
    End Function
    Private Sub Ensure_SQLSERVER_DatabaseExists(databaseConnection As ADODB.Connection)
    '===================================================================================================================
    'Summary: Creates a COT database if it doesn't exist for the connected database via [databaseConnection]
    'Parameters:
    '   eReport - A ReportEnum used to select for Legacy,Disaggregated or TFF tables.
    '   oiType - Used to select for a Futures Combined or Futures Only table.
    '   fieldInfoByEditedName - Collection of FieldInfo instances from which field names and types are determined.
    '===================================================================================================================
        Dim sql$

        With databaseConnection
            On Error GoTo Finally
            If .State = adStateClosed Then .Open

            If .Properties("Current Catalog").value = "master" Then
                sql = "IF NOT EXISTS(SELECT name from sys.databases WHERE name='" & SqlServerDatabaseName & "') BEGIN CREATE Database " & SqlServerDatabaseName & "; END;"
                .Execute sql, Options:=adCmdText Or adExecuteNoRecords

                sql = "IF NOT EXISTS(SELECT name from " & SqlServerDatabaseName & ".sys.tables WHERE name='" & PriceDataTableName & "') BEGIN CREATE TABLE " & SqlServerDatabaseName & "." & PriceDataTableName & " (report_date_as_yyyy_mm_dd Date NOT NULL, cftc_contract_market_code VARCHAR(10) NOT NULL, price smallmoney NOT NULL, Primary Key (report_date_as_yyyy_mm_dd, cftc_contract_market_code)); END;"
                .Execute sql, Options:=adCmdText Or adExecuteNoRecords
            Else
                Err.Raise DbError.UseMasterCatalog, "Ensure_SQLSERVER_DatabaseExists", "Use the master catalog if checking for server existance."
            End If
Finally:
            If Err.Number <> 0 Then PropagateError Err, "Ensure_SQLSERVER_DatabaseExists"
        End With

    End Sub
    Private Sub CreateCommitmentsOfTradersTable_SqlServer(eReport As ReportEnum, oiType As OpenInterestEnum, fieldInfoByEditedName As Collection, createForSqlServer As Boolean, databaseConnection As ADODB.Connection)
    '===================================================================================================================
    'Summary: Creates a table within the COT database based on the given parameters.
    'Parameters:
    '   eReport - A ReportEnum used to select for Legacy,Disaggregated or TFF tables.
    '   oiType - Used to select for a Futures Combined or Futures Only table.
    '   fieldInfoByEditedName - Collection of FieldInfo instances from which field names and types are determined.
    '===================================================================================================================
        Dim wantedField As FieldInfo, i&, fieldDeclarations$(), standardName$, sql$, openedConnection As Boolean

        ReDim fieldDeclarations(fieldInfoByEditedName.Count)

        i = LBound(fieldDeclarations)
        For Each wantedField In fieldInfoByEditedName
            With wantedField
                standardName = .EditedName
                Select Case standardName
                    Case "cftc_subgroup_code", "as_of_date_in_form_yymmdd", "cftc_region_code", "cftc_market_code", "cftc_commodity_code", "futonly_or_combined"
                    Case Else
                        If InStrB(standardName, "quotes") = 0 Then
                            Select Case .DataType
                                Case NumericField
                                    If standardName Like "pct*" And Len(standardName) <= 15 Then
                                        fieldDeclarations(i) = standardName & " TINYINT"
                                    Else
                                        fieldDeclarations(i) = standardName & " DECIMAL(5,2)"
                                    End If
                                Case IntegerField
                                    If InStrB(standardName, "trader") <> 0 Then
                                        fieldDeclarations(i) = standardName & " SMALLINT"
                                    ElseIf InStrB(standardName, "pct_of_oi") <> 0 Then
                                        fieldDeclarations(i) = standardName & " TINYINT"
                                    Else
                                        fieldDeclarations(i) = standardName & " INT"
                                    End If
                                Case adDate, adDBDate
                                    fieldDeclarations(i) = standardName & " DATE"
                                Case adVarChar, adVarWChar
                                    If InStrB(standardName, "name") <> 0 Then
                                        fieldDeclarations(i) = standardName & " VARCHAR(90)"
                                    ElseIf InStrB(standardName, "cftc_contract_market_code") <> 0 Then
                                        fieldDeclarations(i) = standardName & " VARCHAR(10)"
                                    Else
                                        fieldDeclarations(i) = standardName & " VARCHAR(120)"
                                    End If
                            End Select
                            standardName = vbNullString
                            i = i + 1
                        End If
                    End Select
            End With

NEXT_FIELD: Next
        If i <> UBound(fieldDeclarations) Then ReDim Preserve fieldDeclarations(i)
        fieldDeclarations(UBound(fieldDeclarations)) = "PRIMARY KEY (report_date_as_yyyy_mm_dd,cftc_contract_market_code));"

        sql = "CREATE TABLE " & GetSqlServerTableName(eReport, oiType) & " (" & Join(fieldDeclarations, ",")

        With databaseConnection
            On Error GoTo Finally
            If .State = adStateClosed Then
                .Open
                openedConnection = True
            End If
            .Execute sql, Options:=adCmdText Or adExecuteNoRecords
Finally:    If openedConnection Then .Close
            If Err.Number <> 0 Then
                PropagateError Err, "CreateCommitmentsOfTradersTable_SqlServer", "Unable to create COT table in SQL Server."
            End If
        End With
    End Sub

    Private Sub Ensure_SQLSERVER_ReportTableExists(eReport As ReportEnum, oiType As OpenInterestEnum, databaseConnection As ADODB.Connection, Optional ByVal tableName As String)
    '===================================================================================================================
    'Summary: Checks if a table name generated by the given parameters exists within the COT database.
    'Parameters:
    '   eReport - A ReportEnum used to select for Legacy,Disaggregated or TFF tables.
    '   oiType - Used to select for a Futures Combined or Futures Only table.
    'Returns:
    '   True if the table exists; otherwise, False.
    '===================================================================================================================
        If LenB(tableName) = 0 Then tableName = GetSqlServerTableName(eReport, oiType, False)

        If Not DoesTableExist(databaseConnection, tableName) Then
            CreateCommitmentsOfTradersTable_SqlServer eReport, oiType, GetExpectedLocalFieldInfo(eReport, False, False, False, False), True, databaseConnection
        End If

    End Sub
    Private Function DoesTableExist(conn As ADODB.Connection, tableName$) As Boolean
    '===================================================================================================================
    'Summary: Checks if a table with the name [tableName] is present within the database connected with [conn]
    'Parameters:
    '   tableName: Table name to check for.
    '   conn: An ADODB.Connection connected to a database to search within.
    'Returns:
    '   True if the table exists; otherwise, false.
    '===================================================================================================================
        Dim sql$
        On Error GoTo Propagate
        If Not conn Is Nothing Then
            With conn
                If .State = adStateClosed Then .Open

                If IsSqlServerConnection(conn) Then
                    With .Properties("Current Catalog")
                        If .value <> SqlServerDatabaseName Then .value = SqlServerDatabaseName
                    End With

                    sql = "SELECT name from sys.tables WHERE name='" & tableName & "';"

                    With .Execute(sql, Options:=adCmdText)
                        DoesTableExist = Not .EOF
                        .Close
                    End With
                Else
                    On Error GoTo ExitFunction
                    .Execute tableName, Options:=adCmdTable
                    DoesTableExist = True
                End If

            End With
        Else
            Err.Raise 424, Description:="Variable |conn| is set to nothing."
        End If
ExitFunction:
    Exit Function
Propagate:
        PropagateError Err, "DoesTableExist"
    End Function
'    Private Sub MigrateTableToSqlServer(eReport As ReportEnum, oiSelection As OpenInterestEnum)
'
'        Dim msAccessRecordSet As Object, sqlServerRecordSet As Object, _
'        sqlServerFieldNames$(), i&, profiler As New TimedTask, transactionStarted As Boolean ',values() As Variant
'
'        Dim msAccessConn As New ADODB.Connection, sqlServerConn As New ADODB.Connection, msAccessTableName$, tableToUpdateName$, createdTable As Boolean, sqlConfirmation As Boolean
'
'        On Error GoTo Finally
'
'        Select Case oiSelection
'            Case FuturesOnly, FuturesAndOptions
'            Case Else
'                Err.Raise DbError.VersionUnacceptable, Description:="variable oiSelection is invalid."
'        End Select
'
''        Set msAccessConn = CreateObject("ADODB.Connection")
''        Set sqlServerConn = CreateObject("ADODB.Connection")
'
'        'Get connection details for both SQL Server and MS Access.
'        If TryGetDatabaseDetails(oiSelection, eReport, sqlServerConn, tableToUpdateName, isSqlServerDetail:=sqlConfirmation, ignoreSqlServerDetails:=False) And TryGetDatabaseDetails(oiSelection, eReport, msAccessConn, msAccessTableName, ignoreSqlServerDetails:=True) Then
'
'            With sqlServerConn
'                '.connectionString = GetSqlServerConnectionString(True)
'                '.cursorLocation = adUseServer
'                If .State = adStateClosed Then .Open
'                If Not sqlConfirmation Then GoTo Finally
'
'                On Error GoTo Catch_SQLSERVER_TableMissing
'                Set sqlServerRecordSet = .Execute(tableToUpdateName, Options:=adCmdTable)
'
'Recieved_MsAccessFields:
'                On Error GoTo Finally
'
'                If Not createdTable Then
'                    With sqlServerRecordSet
'                        If .Fields.Count > 0 Then
'                            If MsgBox("Table already exists are you sure you want to continue?", vbYesNo) <> vbYes Then
'                                .Close
'                                GoTo Finally
'                            End If
'                        End If
'                    End With
'                End If
'            End With
'
'            Set msAccessRecordSet = GetTableFieldsRecordSet(msAccessConn, msAccessTableName)
'
'            Dim msAccessFieldInfo As Collection, msAccessNames$()
'
'            Set msAccessFieldInfo = FilterDatabaseFieldsWithLocalInfo(msAccessRecordSet, GetExpectedLocalFieldInfo(eReport, False, False, False, False))
'
'            sqlServerRecordSet.Close
'
'            Dim uploadCommand As Object, dbField As Object, cmdParameter As Object, recordsProcessedCount&
'
'            sqlServerFieldNames = GetFieldNamesFromRecord(sqlServerRecordSet, False)
'            Set uploadCommand = CreateObject("ADODB.Command")
'
'            With uploadCommand
'                .Prepared = True
'                .ActiveConnection = sqlServerConn
'                .CommandType = adCmdText
'
'                With .Parameters
'                    i = LBound(sqlServerFieldNames)
'                    ReDim fieldValues(UBound(sqlServerFieldNames))
'                    ReDim msAccessNames(UBound(sqlServerFieldNames))
'
'                    For Each dbField In sqlServerRecordSet.Fields
'                        With dbField
'                            Set cmdParameter = uploadCommand.CreateParameter(.name, .Type, adParamInput, value:=Null)
'                            Select Case .Type
'                                Case adNumeric, adDecimal
'                                    With cmdParameter
'                                        .NumericScale = dbField.NumericScale
'                                        .Precision = dbField.Precision
'                                    End With
'                                Case adVarChar, adVarWChar
'                                    cmdParameter.size = .DefinedSize
'                            End Select
'                            msAccessNames(i) = msAccessFieldInfo(.name).DatabaseName
'                        End With
'                        fieldValues(i) = "?"
'                        .Append cmdParameter
'                        i = i + 1
'                    Next dbField
'                End With
'
'                .CommandText = "Insert Into " + tableToUpdateName + "(" + Join(sqlServerFieldNames, ",") + ") Values (" + Join(fieldValues, ",") + ");"
'                'sqlServerConn.BeginTrans: transactionStarted = True
'
'                With profiler
'                    .Start tableToUpdateName
'                    With msAccessRecordSet
'
'                        Do While Not .EOF
'                            On Error GoTo Catch_ParameterValueError
'                            For i = LBound(sqlServerFieldNames) To UBound(sqlServerFieldNames)
'                                uploadCommand.Parameters(sqlServerFieldNames(i)).value = .Fields(msAccessNames(i)).value
'                            Next i
'                            On Error GoTo Catch_ExecutionError
'                            uploadCommand.Execute Options:=adCmdText Or adExecuteNoRecords
'                            On Error GoTo Finally
'                            recordsProcessedCount = recordsProcessedCount + 1
'                            If recordsProcessedCount Mod 5000 = 0 Then
'                                Application.StatusBar = recordsProcessedCount & " " & Round(profiler.ElapsedTime, 3) & "(s)"
'                                DoEvents
'                            End If
'                            .MoveNext
'                        Loop
'                    End With
'                    .EndTask
'                End With
'            End With
'
'            'sqlServerConn.CommitTrans
'Finally:
'            If Err.Number <> 0 Then
'                profiler.EndTask
'                DisplayErr Err, "MigrateTableToSqlServer", "Record number " & recordsProcessedCount
'                'Stop: Resume
'            End If
'
'            Application.StatusBar = vbNullString
'
'            With msAccessConn
'                If .State = adStateOpen Then .Close
'            End With
'
'            With sqlServerConn
'                If .State = adStateOpen Then
'                    'If Err.Number <> 0 And transactionStarted Then .RollbackTrans
'                    .Close
'                End If
'            End With
'            profiler.DPrint
'        End If
'
'        Exit Sub
'Catch_SQLSERVER_TableMissing:
'        If Err.Number = -2147217865 Then
'            On Error GoTo -1: On Error GoTo Finally
'
'            Ensure_SQLSERVER_ReportTableExists eReport, oiSelection, sqlServerConn
'            Set sqlServerRecordSet = GetTableFieldsRecordSet(sqlServerConn, tableToUpdateName)
'            createdTable = True
'            GoTo Recieved_MsAccessFields
'        Else
'            GoTo Finally
'        End If
'        Resume
'Catch_ExecutionError:
'        Select Case Err.Number
'            Case DbError.InvalidCast
'                Resume Next
'            Case DbError.PrimaryKeyViolation_SQLSERVER, DbError.ProviderError
'                ' Violation of primary key
'                Resume Next
'            Case Else
'                GoTo Finally
'        End Select
'Catch_ParameterValueError:
'        If Err.Number = DbError.InvalidParameterAssignment And msAccessRecordSet.Fields(msAccessNames(i)).value = "." Then
'            uploadCommand.Parameters(sqlServerFieldNames(i)).value = Null
'            Resume Next
'        Else
'            GoTo Finally
'        End If
'
'    End Sub
    Private Sub InsertIntoPriceTable(dataToUpload() As Variant, priceColumn&, contractCodeColumn&, dateColumn&, databaseConnection As ADODB.Connection)
    '===================================================================================================================
    'Summary: Inserts price data from an array into a new record within the database.
    'Parameters:
    '   dataToUpload - Array that contains data to upload.
    '   priceColumn -Column within [dataToUpload] that contains price data for the uploaded record.
    '   contractCodeColumn - Column within [dataToUpload] that contains the contract code for the uploaded record.
    '   dateColumn - Column within [dataToUpload] that contains dates for the uploaded record.
    '   databaseConnection - An open connection the SQL Server database.
    '===================================================================================================================
        Dim iRow&, databaseCommand As New ADODB.command, connectedToSqlServer As Boolean

        On Error GoTo Catch_Error
        
        connectedToSqlServer = IsSqlServerConnection(databaseConnection)
        
        Const InsertText$ = "Insert INTO " & PriceDataTableName & " (Price, cftc_contract_market_code, report_date_as_yyyy_mm_dd) Values (?,?,?);", _
              UpdateText$ = "Update " & PriceDataTableName & " Set Price = ? WHERE cftc_contract_market_code = ? AND report_date_as_yyyy_mm_dd= ?;"
        
        With databaseCommand
        
            .ActiveConnection = databaseConnection
            .CommandType = adCmdText
            .Prepared = True
            
            .CommandText = InsertText
            ' Parameter addition order matters.
            .Parameters.Append .CreateParameter("@Price", adCurrency)
            .Parameters.Append .CreateParameter("@Code", IIf(connectedToSqlServer, adVarChar, adVarWChar), size:=10)
            .Parameters.Append .CreateParameter("@Date", IIf(connectedToSqlServer, adDBDate, adDate))

            For iRow = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
                If Not IsEmpty(dataToUpload(iRow, priceColumn)) Then
                    .Parameters("@Date").value = dataToUpload(iRow, dateColumn)
                    .Parameters("@Code").value = dataToUpload(iRow, contractCodeColumn)
                    .Parameters("@Price").value = dataToUpload(iRow, priceColumn)
                    .Execute Options:=adCmdText Or adExecuteNoRecords
                End If
            Next iRow
            
        End With
        
        Exit Sub
Catch_Error:

        Dim objError As ADODB.Error
        
        Select Case Err.Number
            Case DbError.ProviderError, DbError.PrimaryKeyViolation_SQLSERVER
                For Each objError In databaseConnection.Errors
                    With objError
                        If .Number = Err.Number And (Not connectedToSqlServer And .SqlState = 3022) Or (connectedToSqlServer And .NativeError = 2627) Then
                            databaseConnection.Errors.Clear
                            Resume Next
                        End If
                    End With
                Next objError
        End Select
        
        If databaseConnection.Errors.Count > 0 Then AppendErrorDescription Err, SummarizeAdoError(databaseConnection.Errors, Err.Number)
        
        PropagateError Err, "InsertIntoPriceTable"
        
    End Sub
    Private Function QueryForContractPQ(eReport As ReportEnum, contractCode As String, oiType As OpenInterestEnum, Optional profiler As TimedTask) As Variant()
    '===================================================================================================================
    'Summary: Retrieves data from SQL Server via Power Query.
    'Parameters:
    '   eReport - A ReportEnum used to select for Legacy,Disaggregated or TFF tables.
    '   oiType - Used to select for a Futures Combined or Futures Only table.
    '   contractCode - Contract code to filter for.
    'Returns:
    '   A variant array of data retrieved from the server.
    '===================================================================================================================
        Dim i&, str$(), initial$, wb As Workbook
        Const profilerDescription$ = "Power Query Execution", extractionText$ = "Extract From Worksheet"

        On Error GoTo Propagate

        initial = ConvertReportTypeEnum(eReport)
        
        Set wb = ThisWorkbook
        
        With wb.Queries("ContractCode")
            .Formula = Chr(34) & contractCode & Chr(34) & " " & Split(.Formula, " ", 2)(1)
        End With

        With wb.Queries(initial) '("SqlServer_DataSelector")

            str = Split(.Formula, vbNewLine)
            str(1) = vbTab & "details = [reportInitial = """ & initial & """, contractCode = """ & contractCode & """, oiType = " & oiType & "],"
            .Formula = Join(str, vbNewLine)

            If Not profiler Is Nothing Then profiler.StartSubTask profilerDescription
            .Refresh
            If Not profiler Is Nothing Then
                With profiler
                    .StopSubTask profilerDescription
                    .StartSubTask extractionText
                End With
            End If
            
            Err.Raise DbError.NotImplemented, Description:="Target worksheet or table doesn't exist."
            'QueryForContractPQ = PowerQuery_Server.ListObjects(initial).DataBodyRange.Value2

            If Not profiler Is Nothing Then profiler.StopSubTask extractionText
        End With

        Exit Function
Propagate:
        PropagateError Err, "QueryForContractPQ"
    End Function
    Public Sub CheckMsAccessSize(filePath$)
    
        Dim fs As New Scripting.FileSystemObject, databaseFile As file
        
        Const warningSize As Long = 1.8 * 10 ^ 9
        
        On Error GoTo Propagate
        
        If fs.FileExists(filePath) Then
            Set databaseFile = fs.GetFile(filePath)
            
            If databaseFile.size >= warningSize Then
                MsgBox "Your database is nearing its' max size. MS Access files have a max size of 2GB." & vbNewLine & _
                "Please consider migrating data to SQL Server." & vbNewLine & vbNewLine & _
                "Contact MoshiM_UC@outlook.com for assistance."
            End If
        End If
        Exit Sub
Propagate:
        PropagateError Err, "CheckMsAccessSize"
    End Sub
    
#End If
