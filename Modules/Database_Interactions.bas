Attribute VB_Name = "Database_Interactions"
#If DatabaseFile And Not Mac Then

    Private AfterEventHolder As ClassQTE
    Public Const ERROR_REQUESTED_VERSION_NOT_ACCEPTABLE As Long = vbObjectError + 515
    Public Const ERROR_NO_WANTED_FIELDS As Long = vbObjectError + 516
    #Const TimersEnabled = False
    
    Option Explicit
    Function TryGetDatabaseDetails(wantedVersion As OpenInterestType, reportType$, Optional ByRef adodbConnection As Object, _
                        Optional ByRef tableNameToReturn$, Optional ByRef databasePath$, Optional ByRef suppressMsgBoxIfUnavailable As Boolean = False) As Boolean
    '===================================================================================================================
        'Purpose: Determines if database exists. If it does the appropriate variables or properties are assigned values if needed.
        'Inputs:
        '        reportType - One of L,D,T to repersent which database to delete from.
        '        getFuturesAndOptions - true for futures+options and false for futures only.
        'Outputs:
    '===================================================================================================================
        Dim Report_Name$, isDatabaseAvailable As Boolean
        
        If Not wantedVersion = OpenInterestType.OptionsOnly Then
        
            With GetStoredReportDetails(reportType)
                If reportType = "T" Then
                    Report_Name = "TFF"
                Else
                    Report_Name = .FullReportName.Value2
                End If
                databasePath = .CurrentDatabasePath.Value2
            End With
            
            isDatabaseAvailable = FileOrFolderExists(databasePath) And LenB(databasePath) > 0
            
            If Not isDatabaseAvailable And Not suppressMsgBoxIfUnavailable Then
                MsgBox Report_Name & " database not found."
            ElseIf Not adodbConnection Is Nothing And isDatabaseAvailable Then
                adodbConnection.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databasePath & ";"
            End If
            
            tableNameToReturn = Report_Name & IIf(wantedVersion = OpenInterestType.FuturesAndOptions, "_Combined", "_Futures_Only")
            TryGetDatabaseDetails = isDatabaseAvailable
            
        End If
        
    End Function
    Private Function FilterColumnsAndDelimit(fieldsInDatabase$(), reportType$, includePriceColumn As Boolean) As String
    '===================================================================================================================
    'Loops table found on Variables Worksheet that contains True/False values for wanted columns
    'An array of wanted columns with some re-ordering is returned
    '===================================================================================================================
        Dim wantedColumns() As Variant
        
        wantedColumns = Filter_Market_Columns(False, True, convert_skip_col_to_general:=False, reportType:=reportType, Create_Filter:=True, inputA:=fieldsInDatabase)
        
        If includePriceColumn Then
            ReDim Preserve wantedColumns(LBound(wantedColumns) To UBound(wantedColumns) + IIf(includePriceColumn = True, 1, 0))
            wantedColumns(UBound(wantedColumns)) = "Price"
        End If
        
        FilterColumnsAndDelimit = WorksheetFunction.TextJoin(",", True, wantedColumns)
        
    End Function
    Function FilteredFieldsFromRecordSet(record As Object, fieldInfoByEditedName As Collection) As Collection
            
        Dim Item As Variant, EditedName$, output As New Collection, FI As FieldInfo
        
        On Error GoTo Catch_MissingKey
        
        For Each Item In record.Fields
        
            EditedName = EditDatabaseNames(Item.Name)
            
            Set FI = fieldInfoByEditedName(EditedName)
            
            With FI
                If .isMissing = False And Not .EditedName = "id" Then
                    Call .EditDatabaseName(Item.Name)
                    output.Add FI, .EditedName
                End If
            End With
            
AttemptNextField:
        Next Item
        
        Set FilteredFieldsFromRecordSet = output
        Exit Function
        
Catch_MissingKey:
        Resume AttemptNextField
    End Function
    Function RecordSetFieldNames(record As Object, encloseFieldsInBrackets As Boolean) As String()
    '===================================================================================================================
    'record is a RecordSET object containing a single row of data from which field names are retrieved,formatted and output as an array
    '===================================================================================================================
        Dim x As Long, Z As Byte, fieldNamesInRecord$(), currentFieldName$
    
        ReDim fieldNamesInRecord(1 To record.Fields.count - 1)
        
        For x = 0 To record.Fields.count - 1
            
            currentFieldName = record(x).Name
            
            If Not currentFieldName = "ID" Then
                Z = Z + 1
                If encloseFieldsInBrackets Then
                    fieldNamesInRecord(Z) = "[" & currentFieldName & "]"
                Else
                    fieldNamesInRecord(Z) = currentFieldName
                End If
            End If
                
        Next x
        
        RecordSetFieldNames = fieldNamesInRecord
    
    End Function
    Public Function FilterCollectionOnFieldInfoKey(databaseFields As Collection, localFieldInfo As Collection) As Collection
        
        Dim CC As New Collection, FI As FieldInfo, editedFieldName$
        
        On Error Resume Next
        
        With CC
            For Each FI In localFieldInfo
                editedFieldName = FI.EditedName
                .Add databaseFields(editedFieldName), editedFieldName
            Next FI
        End With
        
        With Err
            If .Number <> 0 Then .Clear
        End With
        Set FilterCollectionOnFieldInfoKey = CC
    
    End Function
    
    Public Function QueryDatabaseForContract(reportType$, ByVal wantedVersion As OpenInterestType, wantedContractCode$, Optional sortOrder As XlSortOrder = xlAscending) As Variant()
    '===================================================================================================================
    'Retrieves filtered data from database and returns as an array
    '===================================================================================================================
        Dim record As Object, adodbConnection As Object, tableNameWithinDatabase$, wantedFieldInfo As Collection
    
        Dim SQL$, delimitedWantedColumns$, allFieldNames$(), secondaryTable$, _
        databaseFields As Collection, optionsOnlyFields$(), iCount As Long, sqlFieldName$, _
        detailedEditNeeded As Boolean, result() As Variant
        
        Dim oiSelectionIndex As Byte, currentFieldEdited$, groupedTraderData As Collection, _
        traderGroup$, dateFieldName$, wantedField As FieldInfo, swappedToFuturesAndOptions As Boolean
        
        Const FutOnly$ = "FutOnly", FutOpt$ = "FutOpt"
        
        On Error GoTo Finally
        
        Set adodbConnection = CreateObject("ADODB.Connection")
        Set wantedFieldInfo = GetExpectedLocalFieldInfo(reportType, True, True, True, True)
        
        If wantedFieldInfo.count = 0 Then
            On Error GoTo 0
            Err.Raise ERROR_NO_WANTED_FIELDS, "QueryDatabaseForContract", "No wanted fields have been selected."
        End If
        
        If TryGetDatabaseDetails(IIf(wantedVersion = OpenInterestType.OptionsOnly, OpenInterestType.FuturesAndOptions, wantedVersion), reportType, adodbConnection, tableNameWithinDatabase) Then
    
            With adodbConnection
                .Open
                Set record = .Execute(tableNameWithinDatabase, , adCmdTable)
            End With
            
            allFieldNames = RecordSetFieldNames(record, encloseFieldsInBrackets:=False)
            record.Close
            
            Set databaseFields = New Collection
            
            With databaseFields
                For iCount = LBound(allFieldNames) To UBound(allFieldNames)
                    .Add "[" & allFieldNames(iCount) & "]", EditDatabaseNames(allFieldNames(iCount))
                Next iCount
            End With
            
            Erase allFieldNames
Create_SQL_Statement:
            If Not wantedVersion = OpenInterestType.OptionsOnly Then
            
                delimitedWantedColumns = Join(ConvertCollectionToArray(FilterCollectionOnFieldInfoKey(databaseFields, wantedFieldInfo)), ",")
                SQL = "SELECT " & delimitedWantedColumns & " FROM " & tableNameWithinDatabase & " WHERE " & databaseFields("cftc_contract_market_code") & "='" & wantedContractCode & "' ORDER BY " & databaseFields("report_date_as_yyyy_mm_dd") & " " & IIf(sortOrder = xlAscending, "ASC", "DESC") & ";"
                
            ElseIf TryGetDatabaseDetails(OpenInterestType.FuturesOnly, reportType, tableNameToReturn:=secondaryTable) Then
                            
                'Spread column will now be the number of options offsseting an equivalent futures or option position.
                
                Dim futOptField$, isTotalColumn As Boolean, isTraderColumn As Boolean, _
                isLongColumn As Boolean, isSpreadColumn As Boolean, tempRef$, reDistributeSpread As Boolean
                
                ReDim optionsOnlyFields(1 To wantedFieldInfo.count)
                
                Set groupedTraderData = New Collection
                ' If TRUE then the spread will be removed and +1 will be added to longs and shorts.
                ' Else if FUT+OPT - FUT < 0 then -1 from spread and +1 to opposite side.
                reDistributeSpread = False
                
                iCount = 0
                For Each wantedField In wantedFieldInfo
                
                    iCount = iCount + 1
                    
                    With wantedField
                    
                        currentFieldEdited = .EditedName
                        'This is effectively an inner join.
                        On Error GoTo Catch_WantedFieldMissing_OptionsOnly
                            sqlFieldName = databaseFields(currentFieldEdited)
                        On Error GoTo Finally
                        
                        futOptField = FutOpt & "." & sqlFieldName
                        
                        isLongColumn = InStrB(1, currentFieldEdited, "long") > 0
                        isSpreadColumn = InStrB(1, currentFieldEdited, "spread") > 0
                        
                        Select Case .DataType
                        
                            Case adInteger
                                                                            
                                optionsOnlyFields(iCount) = futOptField & "-" & FutOnly & "." & sqlFieldName
                                isTraderColumn = InStrB(1, currentFieldEdited, "trader") > 0
                                detailedEditNeeded = Not (currentFieldEdited Like "*oi*" Or isTraderColumn)
                                    
                                If detailedEditNeeded Then
                                
                                    isTotalColumn = InStrB(1, currentFieldEdited, "tot") > 0
                                    ' If not change column.
                                    If InStrB(1, currentFieldEdited, "change") = 0 Then
                                        ' Calculate difference with a minimum value of 0. Exclude spread columns.
                                        ' Store column name in relevant collection.
                                        If Not (isSpreadColumn Or isTotalColumn) Then
                                        
                                            Select Case Split(currentFieldEdited, "_", 2)(0)
                                                Case "prod", "comm", "nonrept"
                                                    'These groups can't spread.
                                                Case Else
                                                    'if FutOpt-FutOnly < 0 then the trader added positions that ended up in spread. Add the abs(of negative number) to opposite group.
                                                    'If above condition then ensure contract is removed from spread.
                                                    If isLongColumn Then
                                                        If reDistributeSpread Then
                                                            tempRef = "." & databaseFields(Replace$(currentFieldEdited, "long", "spread"))
                                                        Else
                                                            tempRef = "." & databaseFields(Replace$(currentFieldEdited, "long", "short"))
                                                        End If
                                                    Else
                                                        If reDistributeSpread Then
                                                            tempRef = "." & databaseFields(Replace$(currentFieldEdited, "short", "spread"))
                                                        Else
                                                            tempRef = "." & databaseFields(Replace$(currentFieldEdited, "short", "long"))
                                                        End If
                                                    End If
    
                                                    If reDistributeSpread Then
                                                        'Column + Spread
                                                        optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & "+(" & FutOpt & tempRef & "-" & FutOnly & tempRef & ")"
                                                    Else
                                                        'Column with min value of 0 + IIF(Opposing Side Options Only count< 0,ABS(Options Only opposing side),0)
                                                        optionsOnlyFields(iCount) = "(IIF(" & optionsOnlyFields(iCount) & ">=0," & optionsOnlyFields(iCount) & ",0)+ IIF((" & FutOpt & tempRef & "-" & FutOnly & tempRef & ")<0,ABS(" & FutOpt & tempRef & "-" & FutOnly & tempRef & "),0))"
                                                    End If
                                            End Select
                                            
                                        ElseIf isSpreadColumn Then
                                            
                                            If reDistributeSpread Then
                                                optionsOnlyFields(iCount) = "NULL"
                                            Else
                                                ' If Long < 0 or short<0 add subtract abs(value) from spread column for current trader group.
                                                For oiSelectionIndex = 0 To 1
                                                    tempRef = "." & databaseFields(Replace$(currentFieldEdited, "spread", Array("long", "short")(oiSelectionIndex)))
                                                    optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & " - IIF(" & FutOpt & tempRef & "-" & FutOnly & tempRef & "<0,ABS(" & FutOpt & tempRef & "-" & FutOnly & tempRef & "),0)"
                                                Next oiSelectionIndex
                                            End If
                                        End If
                                        ' Store column with raw positions
                                        If Not (isSpreadColumn And reDistributeSpread) Then
                                            traderGroup = Split(currentFieldEdited, "_", 2)(0)
                                            On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                groupedTraderData(traderGroup).Add currentFieldEdited, IIf(isLongColumn, "long", IIf(isSpreadColumn, "spread", "short"))
                                            On Error GoTo Finally
                                        End If
                                        
                                    ElseIf Not isTotalColumn Then
                                        ' If not change in total or spread.
                                        ' Store change column name in relevant collection.
                                        traderGroup = Split(currentFieldEdited, "_", 4)(2)
                                        
                                        Select Case traderGroup
                                            Case "comm", "prod", "nonrept"
                                                'These groups don't have spreading to effect changes.
                                            Case Else
                                                If Not (isSpreadColumn And reDistributeSpread) Then
                                                    On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                    groupedTraderData(traderGroup).Add currentFieldEdited, IIf(isLongColumn, "longChange", IIf(isSpreadColumn, "spreadChange", "shortChange"))
                                                    On Error GoTo Finally
                                                End If
                                                optionsOnlyFields(iCount) = "NULL"
                                        End Select
                                        
                                    End If
                                    
                                ElseIf isTraderColumn Then
                                    optionsOnlyFields(iCount) = "NULL"
                                End If
                                
                            Case adNumeric
                                Select Case Split(currentFieldEdited, "_", 2)(0)
                                    Case "pct"
                                        Select Case currentFieldEdited
                                            Case "pct_of_oi_all", "pct_of_oi_old", "pct_of_oi_other"
                                                optionsOnlyFields(iCount) = 100
                                            Case Else
                                                If Not (isSpreadColumn And reDistributeSpread) Then
                                                    traderGroup = Split(currentFieldEdited, "_", 5)(3)
                                                    On Error GoTo Catch_OptionsOnly_TraderGroup_Missing
                                                        groupedTraderData(traderGroup).Add currentFieldEdited, IIf(isLongColumn, "longPct", IIf(isSpreadColumn, "spreadPct", "shortPct"))
                                                    On Error GoTo Finally
                                                End If
                                                optionsOnlyFields(iCount) = "NULL"
                                        End Select
                                    Case "conc"
                                        'Concentration
                                        optionsOnlyFields(iCount) = "NULL"
                                End Select
                                
                            Case Else
                                optionsOnlyFields(iCount) = futOptField
                        End Select
                        
                    End With
OptionsOnly_AssignAlias:
                    optionsOnlyFields(iCount) = optionsOnlyFields(iCount) & " as " & currentFieldEdited
                Next wantedField
                
                'optionsOnlyFields(iCount + 1) = FutOpt & ".[Price]"
                dateFieldName = databaseFields("report_date_as_yyyy_mm_dd")
                
                SQL = " SELECT " & Join(optionsOnlyFields, ",") & " FROM " & tableNameWithinDatabase & " as " & FutOpt & _
                " INNER JOIN " & secondaryTable & " as " & FutOnly & _
                " ON ((" & FutOpt & "." & dateFieldName & "=" & FutOnly & "." & dateFieldName & ") AND (" & FutOpt & ".[CFTC_Contract_Market_Code]=" & FutOnly & ".[CFTC_Contract_Market_Code]))" & _
                " WHERE " & FutOpt & ".[CFTC_Contract_Market_Code]='" & wantedContractCode & "' ORDER BY " & FutOpt & ".[Report_Date_as_YYYY-MM-DD] " & IIf(sortOrder = xlAscending, "ASC", "DESC") & ";"
                
            Else
                Err.Raise ERROR_REQUESTED_VERSION_NOT_ACCEPTABLE, "QueryDatabaseForContract", "ERROR_REQUESTED_VERSION_NOT_ACCEPTABLE"
            End If
            
            delimitedWantedColumns = vbNullString
            On Error GoTo Finally
            
            With record
                .Open SQL, adodbConnection
                On Error GoTo Data_Unavailable
                result = TransposeData(.GetRows)
                On Error GoTo Finally
                .Close
            End With
            
            Set databaseFields = Nothing
            adodbConnection.Close
            Set record = Nothing: Set adodbConnection = Nothing
            
            ' Calculate Changes and percent of OI.
            If wantedVersion = OpenInterestType.OptionsOnly Then
                 
                Dim Item As Collection, columnTarget As Byte, pctOiColumn As Byte, offsetN As Long, _
                calculatePctOI As Boolean, calculateChange As Boolean, oiSelectionForGroup$, positionColumn As Byte
                ' Loop trader groups
                
                Const oiColumn As Byte = 3
                offsetN = IIf(sortOrder = xlAscending, -1, 1)
                For Each Item In groupedTraderData
                    
                    On Error GoTo OptionsOnly_Next_GroupOiSelection
                    For oiSelectionIndex = 0 To 2
                        
                        oiSelectionForGroup = Array("long", "short", "spread")(oiSelectionIndex)
                                            
                        calculatePctOI = HasKey(Item, oiSelectionForGroup & "Pct")
                        calculateChange = HasKey(Item, oiSelectionForGroup & "Change")
    
                        If calculatePctOI Then pctOiColumn = wantedFieldInfo(Item(oiSelectionForGroup & "Pct")).ColumnIndex
                        If calculateChange Then columnTarget = wantedFieldInfo(Item(oiSelectionForGroup & "Change")).ColumnIndex
                                                                                  
                        If calculatePctOI Or calculateChange Then
                            
                            positionColumn = wantedFieldInfo(Item(oiSelectionForGroup)).ColumnIndex
                            
                            On Error GoTo Catch_OptionsOnlyCalculationError
                            For iCount = UBound(result, 1) To LBound(result, 1) Step -1
                                If calculatePctOI Then
                                    If result(iCount, oiColumn) <> 0 Then
                                        result(iCount, pctOiColumn) = Round(100 * (result(iCount, positionColumn) / result(iCount, oiColumn)), 1)
                                    Else
                                        result(iCount, pctOiColumn) = 0
                                    End If
                                End If
                                'This line won't generate an error unless missing data in database.
                                If calculateChange Then result(iCount, columnTarget) = result(iCount, positionColumn) - result(iCount + offsetN, positionColumn)
                            Next iCount
                            On Error GoTo OptionsOnly_Next_GroupOiSelection
                            
                        End If
OptionsOnly_Next_GroupOiSelection:
                        On Error GoTo -1
                    Next oiSelectionIndex
                Next Item
                
                On Error GoTo 0
            End If
            QueryDatabaseForContract = result
        End If
        
Finally:
        If Not record Is Nothing Then
            If record.State = adStateOpen Then record.Close
            Set record = Nothing
        End If
        
        If Not adodbConnection Is Nothing Then
            If adodbConnection.State = adStateOpen Then adodbConnection.Close
            Set adodbConnection = Nothing
        End If
        
        If Err.Number <> 0 Then Call PropagateError(Err, "QueryDatabaseForContract")
        
        Exit Function
        
Data_Unavailable:
        
        With Err
            If .Number = 3021 Then
                If wantedVersion <> OpenInterestType.OptionsOnly Then
                    .Description = "No data available for " & reportType & "_" & wantedContractCode & "_{ " & ConvertOpenInterestTypeToName(IIf(swappedToFuturesAndOptions, OpenInterestType.OptionsOnly, wantedVersion)) & " }. " & vbNewLine & .Description
                Else
                    ' It's likely that the wanted contract doesn't exist in Futures Only so SQL statement fails.
                    wantedVersion = OpenInterestType.FuturesAndOptions
                    swappedToFuturesAndOptions = True
                    With record
                        If .State = adStateOpen Then .Close
                    End With
                    
                    Resume Create_SQL_Statement
                End If
            End If
        End With
        
        GoTo Finally
Catch_OptionsOnly_TraderGroup_Missing:
        On Error GoTo Finally
        groupedTraderData.Add New Collection, traderGroup
        Resume
Catch_OptionsOnlyCalculationError:
        
        Select Case Err.Number
            Case 9
                'Subscript out of range when calculating change.
                Resume Next
            Case 6
                'Overflow error: Division by 0.
                result(iCount, pctOiColumn) = 0
                Resume Next
            Case Else
                Resume OptionsOnly_Next_GroupOiSelection
        End Select
Catch_WantedFieldMissing_OptionsOnly:
        optionsOnlyFields(iCount) = "NULL"
        Resume OptionsOnly_AssignAlias
    End Function
    
    Public Sub Update_Database(dataToUpload() As Variant, versionToUpdate As OpenInterestType, reportType$, debugOnly As Boolean, suppliedFieldInfoByEditedName As Collection)
    '===================================================================================================================
        'Purpose: Uploads rows contained within dataToUpload to a database determined by other Parameters.
        'Inputs:
        '       dataToUpload  - 2D array of rows to be uploaded.
        '       versionToUpdate - True if data being uploaded is Futures + Options combined.
        '       reportType - One of L,D,T to repersent which database to upload to.
        '       suppliedFieldInfoByEditedName - A Collection of FieldInfo instances used to describe columns contained within dataToUpload.
    '===================================================================================================================
       
        Dim tableToUpdateName$, wantedDatabaseFields As Collection, row As Long, _
        legacyCombinedTableName$, legacyDatabasePath$, useYear3000 As Boolean, _
        uploadingLegacyCombinedData As Boolean, oldestDateInUpload As Date, uploadToDatabase As Boolean
        
        If debugOnly Then
            If MsgBox("Debug Active: Do you want to upload data to databse?", vbYesNo) = vbYes Then uploadToDatabase = True
            If MsgBox("Replace dates with year 3000?", vbYesNo) = vbYes Then useYear3000 = True
            Dim year3000 As Date: year3000 = DateSerial(3000, 1, 1)
        Else
            uploadToDatabase = True
        End If
        
        Dim databaseFieldNamesRecord As Object, adodbConnection As Object, SQL$
        
        On Error GoTo Close_Connection
        
        Const legacy_abbreviation$ = "L"
    
        Set adodbConnection = CreateObject("ADODB.Connection")
        Set databaseFieldNamesRecord = CreateObject("ADODB.RecordSet")
        
        uploadingLegacyCombinedData = (reportType = legacy_abbreviation And versionToUpdate = OpenInterestType.FuturesAndOptions)
        
        If TryGetDatabaseDetails(versionToUpdate, reportType, adodbConnection, tableToUpdateName) Then
    
            With adodbConnection
                '.CursorLocation = adUseServer                                   'Batch update won't work otherwise
                .Open
                'Get a record of all field names within the database.
                Set databaseFieldNamesRecord = .Execute(CommandText:=tableToUpdateName, Options:=adCmdTable)
            End With
            ' Get a ccollection of FieldInfo instances with matching fields for input and target.
            Set wantedDatabaseFields = FilteredFieldsFromRecordSet(databaseFieldNamesRecord, suppliedFieldInfoByEditedName)
            
            databaseFieldNamesRecord.Close
            
            Dim uploadCommand As Object, FI As FieldInfo, Item As Object, obj As Object
            Dim fieldNames As New Collection, fieldValues As New Collection, startedTransaction As Boolean
            
            Set uploadCommand = CreateObject("ADODB.Command")
            
            adodbConnection.BeginTrans
            startedTransaction = True
            
'            Const dateField$ = "[Report_Date_as_YYYY-MM-DD]", _
'              codeField$ = "[CFTC_Contract_Market_Code]", _
'              nameField$ = "[Market_and_Exchange_Names]"
            ' SELECT codeField ,MAX(dateField) From TABLE GROUP BY codeField
            
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
                
                Dim wantedColumn As Byte, attemptedValueAllocation As Boolean
                
                attemptedValueAllocation = True
                
                For row = LBound(dataToUpload, 1) To UBound(dataToUpload, 1)
                
                    For Each Item In .Parameters
                        
                        With Item
                            wantedColumn = wantedDatabaseFields(.Name).ColumnIndex
            
                            If Not (IsError(dataToUpload(row, wantedColumn)) Or IsEmpty(dataToUpload(row, wantedColumn)) Or IsNull(dataToUpload(row, wantedColumn))) Then
                                
                                If IsNumeric(dataToUpload(row, wantedColumn)) Then
                                    .value = dataToUpload(row, wantedColumn)
                                ElseIf dataToUpload(row, wantedColumn) = "." Or LenB(Trim$(dataToUpload(row, wantedColumn))) = 0 Then
                                    .value = Null
                                Else
                                    .value = dataToUpload(row, wantedColumn)
                                End If
                                
                                If .Type = adDate Then
                                    If dataToUpload(row, wantedColumn) > oldestDateInUpload Then oldestDateInUpload = dataToUpload(row, wantedColumn)
                                End If
                                
                            Else
                                .value = Null
                            End If
                        End With
                        
                    Next Item
                    
                    If useYear3000 Then
                        .Parameters("report_date_as_yyyy_mm_dd").value = year3000
                    End If
                    
                    If uploadToDatabase Then .Execute
                    
                Next row
                
            End With
            
            adodbConnection.CommitTrans
            startedTransaction = False
            Set uploadCommand = Nothing
            
            If uploadToDatabase And Not uploadingLegacyCombinedData Then 'retrieve price data from the legacy combined table
                'Legacy COmbined Data should be the first data retrieved
                If TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, legacy_abbreviation, tableNameToReturn:=legacyCombinedTableName, databasePath:=legacyDatabasePath) Then
            
                    'T alias is for table that is being updated
                    SQL = "Update " & tableToUpdateName & " as T INNER JOIN [" & legacyDatabasePath & "]." & legacyCombinedTableName & " as Source_TBL ON Source_TBL.[Report_Date_as_YYYY-MM-DD]=T.[Report_Date_as_YYYY-MM-DD] AND Source_TBL.[CFTC_Contract_Market_Code]=T.[CFTC_Contract_Market_Code]" & _
                        " SET T.[Price] = Source_TBL.[Price] WHERE T.[Report_Date_as_YYYY-MM-DD]>=CDate('" & Format(oldestDateInUpload, "YYYY-MM-DD") & "');"
                
                    adodbConnection.Execute CommandText:=SQL, Options:=adCmdText + adExecuteNoRecords
                    
                End If
                
            End If
            
            If Not debugOnly Then
                With GetStoredReportDetails(reportType)
                    If .OpenInterestType.Value2 = versionToUpdate Then
                        'This will signal to worksheet activate events to update the currently visible data
                        .PendingUpdateInDatabase.Value2 = True
                    End If
                End With
            End If
            
        End If
Close_Connection:
        
        With Err
            If .Number <> 0 Then
            
                .Description = "An error occurred while attempting to update table [ " & tableToUpdateName & " ] in database " & adodbConnection.Properties("Data Source") & _
                vbNewLine & vbNewLine & _
                "Error description: " & .Description
                                                
                If startedTransaction Then adodbConnection.RollbackTrans
                
            End If
        End With
    
        If Not databaseFieldNamesRecord Is Nothing Then
            If databaseFieldNamesRecord.State = adStateOpen Then databaseFieldNamesRecord.Close
            Set databaseFieldNamesRecord = Nothing
        End If
        
        If Not adodbConnection Is Nothing Then
            If adodbConnection.State = adStateOpen Then adodbConnection.Close
            Set adodbConnection = Nothing
        End If
        
        If Err.Number <> 0 Then PropagateError Err, "Update_Database"
        
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
    Sub DeleteCftcDataFromSpecificDatabase(smallest_date As Date, reportType$, versionToDelete As OpenInterestType)
    '===================================================================================================================
        'Purpose: Deletes COT data from database that is as recent as smallest_date.
        'Inputs: smallest_date - all rows with a date value >= to this will be deleted.
        '        reportType - One of L,D,T to repersent which database to delete from.
        '        versionToDelete - true for futures+options and false for futures only.
        'Outputs:
    '===================================================================================================================
    
        Dim SQL$, tableName$, adodbConnection As Object
        
        Set adodbConnection = CreateObject("ADODB.Connection")
        
        If TryGetDatabaseDetails(versionToDelete, reportType, adodbConnection, tableName) Then
            
            On Error GoTo No_Table
            SQL = "DELETE FROM " & tableName & " WHERE [Report_Date_as_YYYY-MM-DD] >= ?;"
            
            With adodbConnection
                .Open
                With CreateObject("ADODB.Command")
                    .ActiveConnection = adodbConnection
                    .CommandText = SQL
                    .commandType = adCmdText
                    .Parameters.Append .CreateParameter("@smallestDate", adDate, adParamInput, value:=smallest_date)
                    .Execute
                End With
                .Close
            End With
            
        End If
        
        Set adodbConnection = Nothing
        Exit Sub
        
No_Table:
        MsgBox "TableL " & tableName & " not found within database."
        
        If Not adodbConnection Is Nothing Then
            If adodbConnection.State = adStateOpen Then adodbConnection.Close
            Set adodbConnection = Nothing
        End If
        
    End Sub
    
    Public Function TryGetLatestDate(ByRef latestDate As Date, reportType$, versionToQuery As OpenInterestType, queryIceContracts As Boolean) As Boolean
    '===================================================================================================================
    'Returns the date for the most recent data within a database
    '===================================================================================================================
    
        Dim tableName$, SQL$, adodbConnection As Object, record As Object
        
        Const iceFilter$ = "('Cocoa','B','RC','G','Wheat','W');"
        
        On Error GoTo Connection_Unavailable
    
        Set adodbConnection = CreateObject("ADODB.Connection")
        
        If TryGetDatabaseDetails(versionToQuery, reportType, adodbConnection, tableName, , True) Then
            
            SQL = "SELECT MAX([Report_Date_as_YYYY-MM-DD]) FROM " & tableName & _
            " WHERE " & IIf(queryIceContracts, vbNullString, "NOT ") & "[CFTC_Contract_Market_Code] IN " & iceFilter
        
            With adodbConnection
                '.CursorLocation = adUseServer
                .Open
                Set record = .Execute(SQL, , adCmdText)
            End With
            
            If Not IsNull(record(0)) Then
                latestDate = record(0)
            Else
                latestDate = 0
            End If
            
            TryGetLatestDate = True
            
        End If
        
Connection_Unavailable:
    
        If Err.Number <> 0 Then TryGetLatestDate = False
    
        If Not record Is Nothing Then
            If record.State = adStateOpen Then record.Close
            Set record = Nothing
        End If
    
        If Not adodbConnection Is Nothing Then
            If adodbConnection.State = adStateOpen Then adodbConnection.Close
            Set adodbConnection = Nothing
        End If
        
    End Function
    Private Sub UpdateDatabasePrices(data As Variant, reportType$, versionToUpdate As OpenInterestType, priceColumn As Byte)
    '===================================================================================================================
    'Updates database with price data from a given array. Array should come from a worksheet
    '===================================================================================================================
        Dim SQL$, tableName$, iRow As Long, adodbConnection As Object, price_update_command As Object, contractCodeColumn As Byte
        
        Const date_column As Byte = 1
        
        contractCodeColumn = priceColumn - 1
    
        Set adodbConnection = CreateObject("ADODB.Connection")
    
        If TryGetDatabaseDetails(versionToUpdate, reportType, adodbConnection, tableName) Then
    
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
        
            For iRow = LBound(data, 1) To UBound(data, 1)
        
                On Error GoTo Exit_Code
                
                With price_update_command
        
                    With .Parameters
                    
                        If Not IsEmpty(data(iRow, priceColumn)) Then
                            .Item("Price").value = data(iRow, priceColumn)
                        Else
                            .Item("Price").value = Null
                        End If
                        
                        .Item("Contract Code").value = data(iRow, contractCodeColumn)
                        .Item("Date").value = data(iRow, date_column)
                        
                    End With
                    
                    .Execute
                    
                End With
                
            Next iRow
            
        End If
        
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
        reportType$, availableContractInfo As Collection, contractCode$, _
        Source_Ws As Worksheet, D As Byte, current_Filters() As Variant, lo As ListObject
        
        Const legacy_initial$ = "L"
        
        For Each WS In Array(LC, DC, TC)
            
            If ThisWorkbook.ActiveSheet Is WS Then
                reportType = Array("L", "D", "T")(D)
                Set Source_Ws = WS
                Exit For
            End If
            D = D + 1
        Next WS
        
        If Source_Ws Is Nothing Then Exit Sub
        
        Set lo = Get_CftcDataTable(reportType)
        
        With GetStoredReportDetails(reportType)
            contractCode = .CurrentContractCode.Value2
            price_column = .RawDataCount.Value2 + 1
        End With
        
        With lo.DataBodyRange
            Worksheet_Data = .Resize(.Rows.count, price_column).value
        End With
              
        Set availableContractInfo = GetAvailableContractInfo
        
        If HasKey(availableContractInfo, contractCode) Then
        
            If TryGetPriceData(Worksheet_Data, price_column, availableContractInfo(contractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
                
                'Scripts are set up in a way that only price data for Legacy Combined databases are retrieved from the internet
                UpdateDatabasePrices Worksheet_Data, legacy_initial, OpenInterestType.FuturesAndOptions, priceColumn:=price_column
                
                'Overwrites all other database tables with price data from Legacy_Combined
                
                HomogenizeWithLegacyCombinedPrices contractCode
                
                ChangeFilters lo, current_Filters
                    
                lo.DataBodyRange.columns(price_column).Value2 = WorksheetFunction.index(Worksheet_Data, 0, price_column)
                
                RestoreFilters lo, current_Filters
            Else
                MsgBox "Unable to retrieve data."
            End If
            
        Else
            MsgBox "A symbol is unavailable for: [ " & contractCode & " ] on worksheet " & Symbols.Name & "."
        End If
        
    End Sub
    
    Private Sub HomogenizeWithLegacyCombinedPrices(Optional specificContractCode$ = vbNullString, Optional minimum_date As Date)
    '===========================================================================================================
    ' Overwrites a given table found within a database with price data from the legacy combined table in the legacy database
    '===========================================================================================================
        Dim SQL$, tableName$, adodbConnection As Object, legacy_database_path$
          
        Dim reportType As Variant, overwritingFuturesAndOptions As Variant, contract_filter$
            
        Const legacy_initial$ = "L"
        
        On Error GoTo Close_Connections
    
        If TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, legacy_initial, databasePath:=legacy_database_path) Then
        
            contract_filter = " WHERE NOT IsNull(F.[Price])"
            
            If LenB(specificContractCode) > 0 Then
                contract_filter = contract_filter & " AND F.[CFTC_Contract_Market_Code] = '" & specificContractCode & "'"
            End If
            
            If Not minimum_date = TimeSerial(0, 0, 0) Then
                If IsDate(minimum_date) Then
                    contract_filter = contract_filter & " AND T.[Report_Date_as_YYYY-MM-DD] >= Cdate('" & Format(minimum_date, "YYYY-MM-DD") & "')"
                End If
            End If
            
            contract_filter = contract_filter
            
            For Each reportType In Array(legacy_initial, "D", "T")
                
                For Each overwritingFuturesAndOptions In Array(OpenInterestType.FuturesAndOptions, OpenInterestType.FuturesOnly)
                    
                    If overwritingFuturesAndOptions = True Then
                        'Related Report tables currently share the same database so only 1 connecton is needed between the 2
                        Set adodbConnection = CreateObject("ADODB.Connection")
                        
                        If TryGetDatabaseDetails(CBool(overwritingFuturesAndOptions), CStr(reportType), adodbConnection) Then
                            adodbConnection.Open
                        End If
                        
                    End If
                    
                    If adodbConnection.State = adStateOpen And Not (reportType = legacy_initial And overwritingFuturesAndOptions = True) Then
                        
                        If TryGetDatabaseDetails(CBool(overwritingFuturesAndOptions), CStr(reportType), tableNameToReturn:=tableName) Then
                    
                            SQL = "UPDATE " & tableName & _
                                " as T INNER JOIN [" & legacy_database_path & "].Legacy_Combined as F ON (F.[Report_Date_as_YYYY-MM-DD] = T.[Report_Date_as_YYYY-MM-DD] AND T.[CFTC_Contract_Market_Code] = F.[CFTC_Contract_Market_Code])" & _
                                " SET T.[Price] = F.[Price]" & contract_filter & ";"
                        
                            adodbConnection.Execute SQL, , adExecuteNoRecords
                            
                        End If
                    End If

                Next overwritingFuturesAndOptions
                
                If Not adodbConnection Is Nothing Then
                    With adodbConnection
                        If .State = adStateOpen Then .Close
                    End With
                    Set adodbConnection = Nothing
                End If
                
            Next reportType
            
        End If
Close_Connections:

        If Err.Number <> 0 Then
            DisplayErr Err, "HomogenizeWithLegacyCombinedPrices"
        End If
        
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
        Dim availableContractInfo As Collection, CO As ContractInfo, SQL$, adodbConnection As Object, _
        tableName$, recordSet As Object, data() As Variant, cmd As Object
    
        Const legacy_initial$ = "L"
        'Const combined_Bool As Boolean = True
        Const price_column As Byte = 3
        
        If Not MsgBox("Are you sure you want to replace all prices?", vbYesNo) = vbYes Then
            Exit Sub
        End If
    
        Set adodbConnection = CreateObject("ADODB.Connection")
            
        If TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, legacy_initial, adodbConnection, tableName) Then
            
            Set availableContractInfo = GetAvailableContractInfo
            Set cmd = CreateObject("ADODB.Command")
            
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
                        .Parameters("@ContractCode").value = CO.contractCode
                        Set recordSet = .Execute
                    End With
                    
                    With recordSet
                        If Not .EOF And Not .BOF Then
                            data = TransposeData(.GetRows)
                            
                            If TryGetPriceData(data, price_column, availableContractInfo(CO.contractCode), overwriteAllPrices:=True, datesAreInColumnOne:=True) Then
                                Call UpdateDatabasePrices(data, legacy_initial, OpenInterestType.FuturesAndOptions, priceColumn:=price_column)
                                HomogenizeWithLegacyCombinedPrices CO.contractCode
                            End If
                        End If
                         .Close
                    End With
        
                End If
                
            Next CO
    
        End If
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
    Public Sub ExchangeTableData(destinationTable As ListObject, versionToQuery As OpenInterestType, reportType$, contractCode$, maintainCurrentTableFilters As Boolean, recalculateWorksheetFormulas As Boolean)
    '===================================================================================================================
    'Retrieves data and updates a given listobject
    '===================================================================================================================
        Dim data() As Variant, Last_Calculated_Column As Byte, rawDataCountForReport As Byte, newContractName$, _
        First_Calculated_Column As Byte, currentTableFilters() As Variant, reportDetails As LoadedData, worksheetForTable As Worksheet
        
        Dim DebugTasks As New TimedTask, queryDescription$, appProperties As Collection, adjustQuantities As Boolean
        Dim unitsColumnNumber As Byte, contractQuantities() As Variant, iRow As Long
        
        Const contractNameColumnInAvailableContracts As Byte = 2
        adjustQuantities = False
        
        #If TimersEnabled Then
            Const calculateFieldTask$ = "Calculations", outputToSheetTask$ = "Output to worksheet.", clearExtraCellsTask$ = "Clear extra cells beneath table"
            Const resizeTableTask$ = "Resize Table.", adjustQuantitiesTask$ = "Ensure quantity homogenity.", calculateTableTask$ = "Formula Calculation for Worksheet"
        #End If
        
        Set appProperties = DisableApplicationProperties(True, True, True)
        
        newContractName = WorksheetFunction.VLookup(contractCode, Available_Contracts.ListObjects(1).DataBodyRange, contractNameColumnInAvailableContracts, 0)
        
        queryDescription = "Query database for [" & reportType & "-" & contractCode & "]" & IIf(versionToQuery = OpenInterestType.FuturesAndOptions, "Futures+Options", "Futures Only")
        
        On Error GoTo Unhandled_Error_Discovered
        
        With DebugTasks
            
            #If TimersEnabled Then
                .Start "ExchangeTableData[" & newContractName & "]"
                .StartSubTask queryDescription
            #End If
            
            With Application
                .StatusBar = "Querying database for > " & newContractName
                data = QueryDatabaseForContract(reportType, versionToQuery, contractCode, xlAscending)
                .StatusBar = vbNullString
            End With
            
            #If TimersEnabled Then
                .StopSubTask queryDescription
            #End If
            
            Set reportDetails = GetStoredReportDetails(reportType)
            
            With reportDetails
                rawDataCountForReport = .RawDataCount.Value2
                First_Calculated_Column = 3 + rawDataCountForReport 'Raw data coluumn count + (price) + (Empty) + (start)
                Last_Calculated_Column = .LastCalculatedColumn.Value2
            End With
            '==========================================================================================================
            ' Determine if any rows need to be adjusted to match the most recent contract size.
            #If TimersEnabled Then
                If adjustQuantities Then .StartSubTask adjustQuantitiesTask
            #End If
            
            unitsColumnNumber = rawDataCountForReport - 1
            ReDim contractQuantities(LBound(data, 1) To UBound(data, 1), 1 To 1)
                                               
            ' Application.Index doesn't work because data may contain null values.
            For iRow = LBound(data, 1) To UBound(data, 1)
                contractQuantities(iRow, 1) = data(iRow, unitsColumnNumber)
            Next iRow
            contractQuantities = GetNumbers(contractQuantities)
            
            If adjustQuantities Then AdjustForQuantityDifference contractQuantities, data, unitsColumnNumber, reportType
            
            #If TimersEnabled Then
                If adjustQuantities Then .StopSubTask adjustQuantitiesTask
                .StartSubTask calculateFieldTask
            #End If
            
            ReDim Preserve data(1 To UBound(data, 1), 1 To Last_Calculated_Column)
            Select Case reportType
                Case "L": data = Legacy_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
                Case "D": data = Disaggregated_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26)
                Case "T": data = TFF_Multi_Calculations(data, UBound(data, 1), First_Calculated_Column, 156, 26, 52)
            End Select
            
            #If TimersEnabled Then
                .StopSubTask calculateFieldTask
            #End If
            
        End With
            
        With destinationTable
            
            Set worksheetForTable = .parent
            worksheetForTable.DisplayPageBreaks = False
            ChangeFilters destinationTable, currentTableFilters
            
            With .DataBodyRange
                            
                #If TimersEnabled Then
                    DebugTasks.StartSubTask outputToSheetTask
                        .Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
                    DebugTasks.StopSubTask outputToSheetTask
                #Else
                    .Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
                #End If
                ' Clear column that contains extracted quantities array formula.
                With .columns(1).offset(0, -1)
                    If .Cells(1, 1).HasArray Then .ClearContents
                End With
                
            End With
            
            #If TimersEnabled Then
                DebugTasks.StartSubTask resizeTableTask
                    .Resize .Range.Resize(UBound(data, 1) + 1, .Range.columns.count)
                DebugTasks.StopSubTask resizeTableTask
            #Else
                .Resize .Range.Resize(UBound(data, 1) + 1, .Range.columns.count)
            #End If

            .DataBodyRange.columns(1).offset(0, -1).Value2 = contractQuantities
  
            With destinationTable.Sort
                If .SortFields.count > 0 Then .Apply
            End With
            
        End With
        
        On Error GoTo Unhandled_Error_Discovered
        
        If maintainCurrentTableFilters Then
            RestoreFilters destinationTable, currentTableFilters
        End If
        
        With reportDetails
            .OpenInterestType.Resize(, 2).Value2 = Array(versionToQuery, False)
            .RowWithinTable.Calculate
        End With
        
        #If TimersEnabled Then
        
            With DebugTasks
            
                With .StartSubTask(clearExtraCellsTask)
                    ClearRegionBeneathTable destinationTable
                    .EndTask
                End With
                
                If recalculateWorksheetFormulas Then
                    With .StartSubTask(calculateTableTask)
                        destinationTable.DataBodyRange.Calculate
                        .EndTask
                    End With
                End If
                
            End With
            
        #Else
            ClearRegionBeneathTable destinationTable
            If recalculateWorksheetFormulas Then
                destinationTable.DataBodyRange.Calculate
            End If
        #End If

Finally:
        #If TimersEnabled Then
            DebugTasks.DPrint
        #End If
        
        EnableApplicationProperties appProperties
        Exit Sub
Unhandled_Error_Discovered:
        
        With HoldError(Err)
            EnableApplicationProperties appProperties
            Call PropagateError(.HeldError, "ExchangeTableData")
        End With
        
    End Sub
    Private Sub AdjustForQuantityDifference(contractQuantities() As Variant, data() As Variant, unitsColumnNumber As Byte, reportType$)
    
        Dim wantedColumnsTableRange  As Range, lastColumnToEdit As Byte, quantityToMatch As Double, _
        ratio As Double, iRow As Long, iColumn As Byte, lastIntegerFieldIndex As Byte
        Const oiColumn As Byte = 3
        
        Set wantedColumnsTableRange = GetAvailableFieldsTable(reportType).DataBodyRange
        ' Get the column previous to the first column with a % in the name
        On Error GoTo Catch_Percentage_Not_Found
        lastIntegerFieldIndex = -1 + Evaluate("=MATCH( ""*%*""," & wantedColumnsTableRange.columns(1).Address(external:=True) & ",0)")
        On Error GoTo Unhandled_Error_Discovered
    
        ' columnNUmber ToEnd is the last column that needs to be edited in the event of a quantity mismatch.
        ' Subtract 1 since contract codes are moved to the end of the data but would otherwise appear in column 4.
        lastColumnToEdit = -1 + Evaluate("=COUNTIF(" & wantedColumnsTableRange.columns(1).offset(, 1).Resize(lastIntegerFieldIndex).Address(external:=True) & ",TRUE)")
    
        quantityToMatch = contractQuantities(UBound(contractQuantities, 1), 1)
    
        For iRow = LBound(contractQuantities, 1) To UBound(contractQuantities, 1) - 1
    
            If contractQuantities(iRow, 1) <> quantityToMatch Then
                ratio = contractQuantities(iRow, 1) / quantityToMatch
                For iColumn = oiColumn To lastColumnToEdit
                    data(iRow, iColumn) = data(iRow, iColumn) * ratio
                Next iColumn
                data(iRow, unitsColumnNumber) = data(UBound(contractQuantities, 1), unitsColumnNumber)
                contractQuantities(iRow, 1) = quantityToMatch
            End If
    
        Next iRow
        
        Exit Sub
        
Catch_Percentage_Not_Found:
        On Error GoTo Unhandled_Error_Discovered
        lastColumnToEdit = Evaluate("=MATCH( ""*Pct*""," & wantedColumnsTableRange.columns(1).Address(external:=True) & ",0)") - 1
        Resume Next
Unhandled_Error_Discovered:
        Call PropagateError(Err, "AdjustForQuantityDifference")
    End Sub
    Public Sub RefreshTableData(reportType$)
    '==================================================================================================
    'This sub is used to update the GUI after contracts have been updated upon activation of the calling worksheet
    '==================================================================================================
        Dim tableToRefresh As ListObject
        
        With GetStoredReportDetails(reportType)
            If .PendingUpdateInDatabase.Value2 = True Then
                Set tableToRefresh = Get_CftcDataTable(reportType)
                Call ExchangeTableData(tableToRefresh, .OpenInterestType.Value2, reportType, .CurrentContractCode.Value2, True, True)
            End If
        End With
        
    End Sub
    Sub Latest_Contracts()
Attribute Latest_Contracts.VB_Description = "Queries available databases for the latest contracts in a specified timeframe."
Attribute Latest_Contracts.VB_ProcData.VB_Invoke_Func = " \n14"
    '=======================================================================================================
    ' Queries the database for the latest contracts within the database.
    '=======================================================================================================
        Dim L_Table$, L_Path$, D_Path$, D_Table$, queryAvailable As Boolean
         
        Dim sqlQuery$, connectionString$, QT As QueryTable, legacyAvailable As Boolean, disaggregatedAvailable As Boolean
    
        Const queryName$ = "Update Latest Contracts"
            
        On Error GoTo Propagate
    
        legacyAvailable = TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, "L", , L_Table, L_Path)
        disaggregatedAvailable = TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, "D", , D_Table, D_Path)
        
        ' For why using % instead of * to match 0 or more characters see the below link.
        'https://stackoverflow.com/questions/48565908/adodb-connection-sql-not-like-query-not-working
        
        If legacyAvailable And disaggregatedAvailable Then
            
            'Query Description:
            '   Select the latest contracts in the Legacy database and join with the latest contracts in
            '   the Disaggregated database that aren't found in Legacy (ICE).
            '   Then Left join those records with disaggregated again to determine whether to assign L,T or D or L,D.
            sqlQuery = "Select contractNames.contractCode,contractNames.contractName,iif(IsNull(FQ.code),'L,T', iif(LCASE(Trim(FQ.name)) LIKE 'ice%ice%','D','L,D')) From" & _
                        "(" & _
                            "(" & _
                                "SELECT {nameField} as contractName,{codeField} as contractCode " & _
                                "From [{L_Path}].{L_Table} " & _
                                "WHERE {dateField} = {date_cutoff} " & _
                                "Union " & _
                                "(SELECT D.{nameField} as contractName,D.{codeField} as contractCode " & _
                                "FROM {D_Path}.{D_Table} as D " & _
                                "LEFT JOIN {L_Path}.{L_Table} as L " & _
                                "ON L.{codeField}= D.{codeField} and D.{dateField}=L.{dateField} " & _
                                "WHERE IsNull(L.{codeField}) AND D.{dateField} = {date_cutoff})" & _
                            ") as contractNames " & _
                            "Left Join" & _
                            "(" & _
                                "Select {codeField} as code,{nameField} as name " & _
                                "From [{D_Path}].{D_Table} WHERE {dateField} = {date_cutoff}" & _
                            ") as FQ " & _
                            "ON FQ.code = contractNames.contractCode" & _
                        ")" & _
                        "Order by contractNames.contractName ASC;"
                    
            #If Mac Then
                Dim dict As New Dictionary
            #Else
                Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
            #End If
            
            With dict
                .Item("nameField") = "[Market_and_Exchange_Names]"
                .Item("dateField") = "[Report_Date_as_YYYY-MM-DD]"
                .Item("codeField") = "[CFTC_Contract_Market_Code]"
                .Item("L_Path") = L_Path
                .Item("L_Table") = L_Table
                .Item("D_Path") = D_Path
                .Item("D_Table") = D_Table
                .Item("date_cutoff") = "CDATE" & Format$(Variable_Sheet.Range("Last_Updated_CFTC").Value2, "('yyyy-mm-dd')")
            End With
            
            Call Interpolator(sqlQuery, dict)
                                
            connectionString = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & L_Path & ";"
            
            With Available_Contracts
                For Each QT In .QueryTables
                    If QT.Name Like queryName & "*" Then
                        queryAvailable = True
                        Exit For
                    End If
                Next QT
                
                If Not queryAvailable Then
                    Set QT = .QueryTables.Add(connectionString, .Range("K1"))
                End If
            End With
            
            With QT
                .CommandText = sqlQuery
                .BackgroundQuery = True
                If queryAvailable Then .Connection = connectionString
                .commandType = xlCmdSql
                .MaintainConnection = False
                .Name = queryName
                .RefreshOnFileOpen = False
                .RefreshStyle = xlOverwriteCells
                .SaveData = False
                .fieldNames = False
                
                Set AfterEventHolder = New ClassQTE
                AfterEventHolder.HookUpLatestContracts QT
                
                .Refresh True
            End With
        
        End If
        
        Exit Sub
Propagate:
        PropagateError Err, "Latest_Contracts"
    End Sub
    Sub Latest_Contracts_After_Refresh(RefreshedQueryTable As QueryTable, success As Boolean)
            
        Dim results() As Variant, commodityGroups As Collection, Item As Variant, iRow As Long, lo As ListObject
        Dim appProperties As Collection
        Const commodityColumn As Byte = 4, subGroupColumn As Byte = 5
        
        Set AfterEventHolder = Nothing
        
        Set appProperties = DisableApplicationProperties(True, False, True)
        
        If success Then
            
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
            Set lo = Available_Contracts.ListObjects("Contract_Availability")
            
            With lo
                With .DataBodyRange
                    .SpecialCells(xlCellTypeConstants).ClearContents
                    .Cells(1, 1).Resize(UBound(results, 1), UBound(results, 2)).Value2 = results
                End With
                .Resize .Range.Cells(1, 1).Resize(UBound(results, 1) + 1, .ListColumns.count)
            End With
            
            ClearRegionBeneathTable lo
            
        End If
        
        EnableApplicationProperties appProperties
        
        Exit Sub
        
Code_Not_Found:
         Resume Next_Commodity_Assignment
    End Sub
    Private Sub Interpolator(inputStr$, dict As Object)
    '=======================================================================================================
    ' Replaces text within {} with a value in the paramArray values.
    '=======================================================================================================
        Dim rightBrace As Long, leftSplit$(), Z As Long, D As Long, noEscapeCharacter As Boolean
                                
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
                    leftSplit(Z) = dict.Item(Left$(leftSplit(Z), rightBrace - 1)) & Right$(leftSplit(Z), Len(leftSplit(Z)) - rightBrace)
                    D = D + 1
                End If
    
            End If
    
        Next Z
    
        inputStr = Join(leftSplit, vbNullString)
    
    End Sub
    
    Function GetDataForMultipleContractsFromDatabase(reportType$, versionToQuery As OpenInterestType, dateSortOrder As XlSortOrder, _
                            Optional maxWeeksInPast As Long = -1, Optional alternateCodes As Variant, _
                            Optional includePriceColumn As Boolean = False) As Collection
    '====================================================================================================================================
    '   Summary: Retrieves data for all favorites or select contracts from the database and stores an array for each contract keyed to its contract code.
    '   Inputs:
    '       reportType: One of L,D or T to select which database to target.
    '       versionToQuery: true if Futures + Options data is wanted; otherwise, false.
    '       sortDataAscending: true to sort data in ascending order by date otherwise false for descending.
    '       maxWeeksInPast: Number of weeks in the past in addition to the current week to query for.
    '       alternateCodes: Specific contract codes to filter for from the database.
    '       includePriceColumn: true if you want to return prices as well.
    '   Returns: A collection of arrays keyed to that contracts contract code.
    '====================================================================================================================================
        Dim SQL$, tableName$, adodbConnection As Object, record As Object, SQL2$, _
        favoritedContractCodes$, queryResult() As Variant, fieldNames$, _
        contractClctn As Collection, allContracts As New Collection, oldestWantedDate As Date, mostRecentDate As Date
         
        Const dateField$ = "[Report_Date_as_YYYY-MM-DD]", _
              codeField$ = "[CFTC_Contract_Market_Code]", _
              nameField$ = "[Market_and_Exchange_Names]", dateColumn As Byte = 1
        
        On Error GoTo Finally
        
        If isMissing(alternateCodes) Then
            ' Get a list of all contract codes that have been favorited.
            queryResult = WorksheetFunction.Transpose(Variable_Sheet.ListObjects("Current_Favorites").DataBodyRange.columns(1).Value2)
        Else
            queryResult = alternateCodes
        End If
        
        favoritedContractCodes = Join(QuotedForm(queryResult, "'"), ",")
              
        Set adodbConnection = CreateObject("ADODB.Connection")
        
        If TryGetDatabaseDetails(versionToQuery, reportType, adodbConnection, tableName) Then
            
            'Set record = CreateObject("ADODB.RecordSet")
            
            With adodbConnection
                .Open
                'Get a record of all field names in tha database.
                Set record = .Execute(CommandText:=tableName, Options:=adCmdTable)
            End With
            
            fieldNames = FilterColumnsAndDelimit(RecordSetFieldNames(record, encloseFieldsInBrackets:=True), reportType, includePriceColumn:=includePriceColumn)    'Field names from database returned as an array
            record.Close
                
            mostRecentDate = Variable_Sheet.Range("Last_Updated_CFTC").Value2
            
            SQL2 = "SELECT " & codeField & " FROM " & tableName & " WHERE " & dateField & " = CDATE('" & Format(mostRecentDate, "yyyy-mm-dd") & "') AND " & codeField & " in (" & favoritedContractCodes & ");"
            
            oldestWantedDate = IIf(maxWeeksInPast > 0, DateAdd("ww", -maxWeeksInPast, mostRecentDate), DateSerial(1970, 1, 1))
            
            SQL = "SELECT " & fieldNames & " FROM " & tableName & _
            " WHERE " & codeField & " in (" & SQL2 & ") AND " & dateField & " >=CDATE('" & oldestWantedDate & "') Order BY " & codeField & " ASC," & dateField & " " & IIf(dateSortOrder = xlAscending, "ASC;", "DESC;")
            
            Erase queryResult
            
            With record
                .Open SQL, adodbConnection, adOpenStatic, adLockReadOnly, adCmdText
                queryResult = TransposeData(.GetRows)
                .Close
            End With
            
            adodbConnection.Close
            
            Dim codeColumn As Byte, nameColumn As Byte, iRow As Long, iColumn As Byte, _
            queryRow() As Variant, output As New Collection
            
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
                    .Add CombineArraysInCollection(allContracts(iRow), Append_Type.Multiple_1d), allContracts(iRow)(1)(codeColumn)
                Next iRow
            End With
            
            Set GetDataForMultipleContractsFromDatabase = output
        
        End If
        
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
            With HoldError(Err)
                DisplayErr Err, "GetDataForMultipleContractsFromDatabase"
                PropagateError .HeldError, "GetDataForMultipleContractsFromDatabase", "Error while assigning array to collection."
            End With
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
    
    Public Sub Generate_Database_Dashboard(callingWorksheet As Worksheet, ReportChr$)
    
        Dim contractClctn As Collection, tempData As Variant, output() As Variant, totalStoch() As Variant, _
        outputRow As Long, tempRow As Long, tempCol As Byte, commercialNetColumn As Byte, _
        dateRange As Long, Z As Byte, targetColumn As Long, versionToQuery As OpenInterestType
        
        Dim dealerNetColumn As Byte, assetNetColumn As Byte, levFundNet As Byte, otherNet As Byte, _
        nonCommercialNetColumn As Byte, totalNetColumns As Byte, _
        reportGroup As Variant, reportedGroups() As Variant, producerNet As Byte, swapNet As Byte, managedNet As Byte
        
        Const threeYearsInWeeks As Long = 156, sixMonthsInWeeks As Byte = 26, oneYearInWeeks As Byte = 52, _
        previousWeeksToCalculate As Byte = 1
        
        On Error GoTo No_Data
        
        If callingWorksheet.Shapes("FUT Only").OLEFormat.Object.value = xlOn Then
            versionToQuery = OpenInterestType.FuturesOnly
        Else
            versionToQuery = OpenInterestType.FuturesAndOptions
        End If
        
        Set contractClctn = GetDataForMultipleContractsFromDatabase(ReportChr, versionToQuery, xlAscending, threeYearsInWeeks + previousWeeksToCalculate + 2)
        
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
        
        Dim lo As ListObject
        
        With callingWorksheet
        
            .Range("A1").Value2 = Variable_Sheet.Range("Last_Updated_CFTC").Value2
            Set lo = .ListObjects("Dashboard_Results" & ReportChr)
            
            With lo
                
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
            ClearRegionBeneathTable lo
        End With
        
        Re_Enable
        
        Exit Sub
        
No_Data:
        MsgBox "An error occurred. " & Err.Description
    End Sub
    
    Public Function GetCftcWorksheet(reportType$, getData As Boolean, getCharts As Boolean) As Worksheet
        
        Dim T As Byte, WSA() As Variant
        
        If getData Then
            WSA = Array(LC, DC, TC)
        ElseIf getCharts Then
            WSA = Array(L_Charts, D_Charts, T_Charts)
        Else
            Err.Raise 5, "GetCftcWorksheet", "Neither getData nor getCharts is TRUE."
        End If
        
        On Error GoTo Catch_ReportType_Not_Found
        T = Application.Match(reportType, Array("L", "D", "T"), 0) - 1
        
        Set GetCftcWorksheet = WSA(T)
        
        Exit Function
Catch_ReportType_Not_Found:
        PropagateError Err, "GetCftcWorksheet", reportType & " isn't 1 of 'L,D,T'."
    End Function
    
    Public Function Get_CftcDataTable(report$) As ListObject
    '==================================================================================================
    '   Returns the ListObject used to store data for the report abbreviated by the report paramater.
    '   Paramater:
    '       - report: One of L,D or T.
    '==================================================================================================
        Set Get_CftcDataTable = GetCftcWorksheet(report, True, False).ListObjects(report & "_Data")
    End Function
    
    Public Sub Save_For_Github()
Attribute Save_For_Github.VB_Description = "Saves the workbook for GitHub if on creator computer.\r\n"
Attribute Save_For_Github.VB_ProcData.VB_Invoke_Func = " \n14"
    '=======================================================================================================
    ' Toggles range value that marks the workboook for upload to github.
    '=======================================================================================================
        If IsOnCreatorComputer Then
            Range("Github_Version").Value2 = True
            Custom_SaveAS Environ("USERPROFILE") & "\Desktop\COT-GIT.xlsb"
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
Attribute OverwritePricesAfterDate.VB_Description = "Use Legacy Combined price data to overwrite prices in all other databases and tables."
Attribute OverwritePricesAfterDate.VB_ProcData.VB_Invoke_Func = " \n14"
    
    '======================================================================================================
    'Will generate an array to represent all data within the legacy combined database since a certain date N.
    'Price data will be retrieved for that array and used to update the database.
    '======================================================================================================
        Dim availableContractInfo As Collection, SQL$, adodbConnection As Object, tableName$, queryResult() As Variant, iCount As Long, wantedCodes$
    
        Const dateField$ = "[Report_Date_as_YYYY-MM-DD]", _
              codeField$ = "[CFTC_Contract_Market_Code]", _
              nameField$ = "[Market_and_Exchange_Names]"
        
        Dim rowIndex As Long, ColumnIndex As Byte, recordsWithSameContractCode As Collection, _
        queryRow() As Variant, recordsByDateByCode As New Collection, minDate As Date
        
        minDate = CDate(InputBox("Input date in form YYYY-MM-DD"))
        
        If MsgBox("Is this the date you want? " & Format(minDate, "mmmm d, yyyy"), vbYesNo) <> vbYes Then Exit Sub
        
        Set adodbConnection = CreateObject("ADODB.COnnection")
        
        If TryGetDatabaseDetails(OpenInterestType.FuturesAndOptions, "L", adodbConnection, tableName) Then
            
            wantedCodes = "('" & Join(Application.Transpose(Symbols.ListObjects("Symbols_TBL").DataBodyRange.columns(1).Value2), "','") & "')"
            
            SQL = "SELECT " & Join(Array(dateField, codeField, "Price"), ",") & " FROM " & tableName & " WHERE " & codeField & " IN " & wantedCodes & " AND " & dateField & " >=Cdate('" & Format(minDate, "yyyy-mm-dd") & "') ORDER BY " & dateField & " ASC;"
            
            Const codeColumn As Byte = 2
        
            With adodbConnection
                .Open
                 queryResult = TransposeData(.Execute(SQL, , adCmdText).GetRows)
                .Close
            End With
            
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
            
            Set availableContractInfo = GetAvailableContractInfo
            
            With recordsByDateByCode
            
                For iCount = .count To 1 Step -1
                    
                    Set recordsWithSameContractCode = .Item(iCount)
                    queryResult = CombineArraysInCollection(recordsWithSameContractCode, Append_Type.Multiple_1d)
                    .Remove queryResult(1, codeColumn)
                    
                    If HasKey(availableContractInfo, CStr(queryResult(1, codeColumn))) Then
                        If TryGetPriceData(queryResult, 3, availableContractInfo(queryResult(1, codeColumn)), True, True) Then
                            .Add queryResult, queryResult(1, codeColumn)
                        End If
                    End If
                
                Next iCount
            
            End With
            
            queryResult = CombineArraysInCollection(recordsByDateByCode, Append_Type.Multiple_2d)
            
            On Error GoTo 0
            
            UpdateDatabasePrices queryResult, "L", True, 3
            
            HomogenizeWithLegacyCombinedPrices minimum_date:=minDate
            
        End If
        
        Exit Sub
        
Create_Contract_Collection:
    
        Set recordsWithSameContractCode = New Collection
        recordsByDateByCode.Add recordsWithSameContractCode, queryRow(codeColumn)
        
        Resume Next
    
    End Sub
    Private Sub FindDatabasePathInSameFolder()
    '===========================================================================================================
    ' Looks for MS Access Database files that haven't been renamed within the same folder as the Excel workbook.
    '===========================================================================================================
        Dim legacy As New LoadedData, TFF As New LoadedData, DGG As New LoadedData, _
        strfile$, foundCount As Byte, folderPath$, databasePathRange As Range, databaseMissing As Boolean
        
        On Error GoTo Prompt_User_About_UserForm
        ' Initializing these classes will wipe database paths if they cannot be found.
        legacy.InitializeClass "L"
        DGG.InitializeClass "D"
        TFF.InitializeClass "T"
        
        folderPath = ThisWorkbook.Path & Application.PathSeparator
        ' Filter for Microsoft Access databases.
        strfile = Dir$(folderPath & "*.accdb")
        
        Do While LenB(strfile) > 0
            
            If LCase$(strfile) Like "*disaggregated.accdb" And IsEmpty(DGG.CurrentDatabasePath.Value2) Then
                DGG.CurrentDatabasePath.Value2 = folderPath & strfile
                foundCount = foundCount + 1
            ElseIf LCase$(strfile) Like "*legacy.accdb" And IsEmpty(legacy.CurrentDatabasePath.Value2) Then
                legacy.CurrentDatabasePath.Value2 = folderPath & strfile
                foundCount = foundCount + 1
            ElseIf LCase$(strfile) Like "*tff.accdb" And IsEmpty(TFF.CurrentDatabasePath.Value2) Then
                TFF.CurrentDatabasePath.Value2 = folderPath & strfile
                foundCount = foundCount + 1
            End If
            
            strfile = Dir
        Loop
        
Prompt_User_About_UserForm:
        
        On Error GoTo 0
        With legacy.CurrentDatabasePath.ListObject.DataBodyRange
            Set databasePathRange = .columns(legacy.CurrentDatabasePath.Column - .Column + 1)
        End With
        
        ' If any database path is still empty then display a message.
        With databasePathRange
            If Evaluate("=COUNTIF(" & .Address(external:=True) & ",""<>"")<>" & .Rows.count) Then
                MsgBox "Database paths couldn't be auto-retrieved." & vbNewLine & vbNewLine & _
                "Please use the Database Paths USerform on the [ " & HUB.Name & " ] worksheet to fill in the needed data."
                
                databaseMissing = True
            End If
        End With
        
        If databaseMissing Then Err.Raise 17, "FindDatabasePathInSameFolder", "Missing Database(s)"
        
    End Sub
    Public Function GetStoredReportDetails(reportType$) As LoadedData
        
        Dim storedData As New LoadedData
        storedData.InitializeClass reportType
        Set GetStoredReportDetails = storedData
        
    End Function
    
    Public Function GetContractInfo_DbVersion() As Collection
    '==============================================================================================
    ' Creates a collection of Contract objects keyed to their contract code for each
    ' available contract within the database.
    '==============================================================================================
    
        Dim Available_Data() As Variant, CD As ContractInfo, iRow As Long, _
        pAllContracts As New Collection, priceSymbol$, usingYahoo As Boolean, symbolsRange As Range
        
        On Error GoTo Propagate
        
        Available_Data = Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.Value2
        
        Const codeColumn As Byte = 1, nameColumn As Byte = 2, availabileColumn As Byte = 3, _
        commodityGroupColumn As Byte = 4, subGroupColumn As Byte = 5, hasSymbolColumn As Byte = 6, isFavoriteColumn As Byte = 7
        
        Set symbolsRange = Symbols.ListObjects("Symbols_TBL").DataBodyRange
        
        For iRow = LBound(Available_Data) To UBound(Available_Data)
            
            priceSymbol = vbNullString
            usingYahoo = False
            
            If Available_Data(iRow, hasSymbolColumn) = True Then
                On Error GoTo Catch_SymbolNotFound
                priceSymbol = WorksheetFunction.VLookup(Available_Data(iRow, codeColumn), symbolsRange, 3, False)
                On Error GoTo Propagate
                usingYahoo = LenB(priceSymbol) > 0
            End If
            
            Set CD = New ContractInfo
            
            With CD
                
                .InitializeBasicVersion CStr(Available_Data(iRow, codeColumn)), CStr(Available_Data(iRow, nameColumn)), CStr(Available_Data(iRow, availabileColumn)), CBool(Available_Data(iRow, isFavoriteColumn)), priceSymbol, usingYahoo
                
                On Error GoTo Possible_Duplicate_Key
                pAllContracts.Add CD, Available_Data(iRow, codeColumn)
                On Error GoTo Propagate
                
           End With
    
        Next iRow
        
        Set GetContractInfo_DbVersion = pAllContracts
    
        Exit Function
        
Possible_Duplicate_Key:
        Resume Next
Catch_SymbolNotFound:
        'priceSymbol = Right$(String$(6, "0") & Available_Data(iRow, codeColumn), 6)
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
Attribute Open_Contract_Selection.VB_Description = "Opens the Contract_Selection userform."
Attribute Open_Contract_Selection.VB_ProcData.VB_Invoke_Func = " \n14"
        Dim reportToLoad$
                
        On Error GoTo Failed_To_Get_Type
            With ThisWorkbook
                reportToLoad = .Worksheets(.ActiveSheet.Name).WorksheetReportType
            End With
        On Error GoTo 0
        
        With Contract_Selection
            .SetReport reportToLoad
            .Show
        End With
Finally:
        Exit Sub
        
Failed_To_Get_Type:
        MsgBox ThisWorkbook.ActiveSheet.Name & " does not have a publicly available WorksheetReportType Function."
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
        
        Dim localFields As Collection, availableContracts As Collection, currentWeekNet As Long, previousWeekNet As Long
        
        Const maxWeeksToReturn As Byte = 52, weekCountOfShift As Byte = 1
        
        mostRecentContractCodes = Application.Transpose(Available_Contracts.ListObjects("Contract_Availability").DataBodyRange.columns(1).Value2)
        
        Set contractDataByCode = GetDataForMultipleContractsFromDatabase("L", OpenInterestType.FuturesOnly, xlAscending, maxWeeksToReturn - 1, mostRecentContractCodes)
        
        If Not contractDataByCode Is Nothing Then
        
            Set localFields = GetExpectedLocalFieldInfo("L", True, True, False, True)
            
            Dim commLong As Byte, commShort As Byte, nonCommLong As Byte, nonCommShort As Byte, codeColumn As Byte, _
            iColumn As Byte, columnLong As Byte, columnShort As Byte, oiColumn As Byte ', clusteringAndConcentration()
            
            commLong = localFields("comm_positions_long_all").ColumnIndex
            commShort = localFields("comm_positions_short_all").ColumnIndex
            nonCommLong = localFields("noncomm_positions_long_all").ColumnIndex
            nonCommShort = localFields("noncomm_positions_short_all").ColumnIndex
            codeColumn = localFields("cftc_contract_market_code").ColumnIndex
            oiColumn = localFields("oi_all").ColumnIndex
        
            'ReDim clusteringAndConcentration(1 To UBound(outputA, 1), 1 To 5)
            
            Dim recentDate As Date, commConcLong As Byte, commConcShort As Byte, nonCommConcShort As Byte, nonCommConcLong As Byte, traderCount As Byte
            Dim longTraders As Byte, shortTraders As Byte, clustering() As Double, iCountCluster As Long, dateColumn As Byte
            
            commConcLong = localFields("pct_of_oi_comm_long_all").ColumnIndex
            commConcShort = localFields("pct_of_oi_comm_short_all").ColumnIndex
            nonCommConcLong = localFields("pct_of_oi_noncomm_long_all").ColumnIndex
            nonCommConcShort = localFields("pct_of_oi_noncomm_short_all").ColumnIndex
            traderCount = localFields("traders_tot_all").ColumnIndex
            
            longTraders = localFields("traders_noncomm_long_all").ColumnIndex
            shortTraders = localFields("traders_noncomm_short_all").ColumnIndex
            dateColumn = localFields("report_date_as_yyyy_mm_DD").ColumnIndex
                
            recentDate = Variable_Sheet.Range("Last_Updated_CFTC").Value2
            ReDim outputA(1 To contractDataByCode.count, 1 To 12)
            
            Set availableContracts = GetAvailableContractInfo
            
            For Each contractData In contractDataByCode
                            
                Dim currentWeek As Byte, comparisonWeek As Byte
                
                currentWeek = UBound(contractData, 1):
                
                On Error GoTo Next_ContracData
                
                If UBound(contractData, 1) >= 2 And contractData(currentWeek, dateColumn) = recentDate Then
                     
                    comparisonWeek = currentWeek - (weekCountOfShift)
                     
                    iRow = iRow + 1
        
                    outputA(iRow, 1) = contractData(currentWeek, codeColumn)
                    
                    On Error GoTo Catch_CodeMissing
                        outputA(iRow, 2) = availableContracts(contractData(currentWeek, codeColumn)).ContractNameWithoutMarket
                    On Error GoTo 0
                    
                    ReDim clustering(1 To UBound(contractData, 1), 1 To 2)
                    
                    For iCountCluster = 1 To UBound(contractData, 1)
                        'Longs
                        clustering(iCountCluster, 1) = contractData(iCountCluster, longTraders) / contractData(iCountCluster, traderCount)
                        'Shorts
                        clustering(iCountCluster, 2) = contractData(iCountCluster, shortTraders) / contractData(iCountCluster, traderCount)
                    Next iCountCluster
                    
                    outputA(iRow, 7) = Stochastic_Calculations(CLng(nonCommConcLong), UBound(clustering, 1), contractData, 1, True)(1)
                    'Long clustering
                    outputA(iRow, 8) = Stochastic_Calculations(1, UBound(clustering, 1), clustering, 1, True)(1)
                    outputA(iRow, 9) = Stochastic_Calculations(CLng(nonCommConcShort), UBound(clustering, 1), contractData, 1, True)(1)
                    'clustering
                    outputA(iRow, 10) = Stochastic_Calculations(2, UBound(clustering, 1), clustering, 1, True)(1)
                    
                    ' Loop positions for both groups.
                    For iColumn = 0 To 3
                    
                        Dim columnTarget As Byte
                        
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
                            
                            Dim calc As Byte: calc = 0
                            
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
Next_ContracData:
                On Error GoTo -1
            Next contractData
            
            On Error GoTo 0
            
            Dim tableDataRng As Range, lo As ListObject, currentFilters As Variant, appProperties As Collection
            
            Set lo = WeeklyChanges.ListObjects("PctNetChange")
            Set tableDataRng = lo.DataBodyRange
            
            With tableDataRng
            
                Set appProperties = DisableApplicationProperties(True, False, True)
                
                ChangeFilters lo, currentFilters
                On Error Resume Next
                    .SpecialCells(xlCellTypeConstants).ClearContents
                On Error GoTo 0
                
                .columns(4).Resize(UBound(outputA, 1), UBound(outputA, 2)).Value2 = outputA
                
                ResizeTableBasedOnColumn lo, .columns(4)
                
                ClearRegionBeneathTable lo
                With lo.Sort
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
                    If .SortFields.count > 0 Then .Apply
                End With
                RestoreFilters lo, currentFilters
                
                WeeklyChanges.Range("reflectedDate").Value2 = Variable_Sheet.Range("Last_Updated_CFTC").Value2
                EnableApplicationProperties appProperties
                
                '=SUM(IF(SUBTOTAL(103,OFFSET([Commercial Net change/Total Position],ROW([Commercial Net change/Total Position])-ROW($A$3),0,1))>0,IF(K10<[Commercial Net change/Total Position],1)))+1
            End With
        Else
            MsgBox "Database Unavailable"
        End If
        
        Exit Sub
    
Catch_CodeMissing:
        Resume Next_ContracData
    End Sub
    
    Private Sub AttemptCross()
    
        Dim dataFieldInfoByEditedName As Collection, notionalValue() As Double, iRow As Long, _
        dataFromDatabase As Collection, reportType$, notionalValuesByCode As New Collection, _
        Code As Variant, contractUnits As Variant, prices As Variant
        
        On Error GoTo Finally
        
        ' Setting equal to -1 will allow all data to be retrieved.
        Const maxWeeksInPast As Long = -1, versionToQuery As Long = OpenInterestType.FuturesOnly
        
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
        
        reportType = "L"
        
        Set dataFromDatabase = GetDataForMultipleContractsFromDatabase(reportType, versionToQuery, xlAscending, maxWeeksInPast, wantedContractCodes, True)
    
        Set dataFieldInfoByEditedName = GetExpectedLocalFieldInfo(reportType, True, True, False)
        
'        For Each Code In wantedContractCodes
'
'            contractUnits = WorksheetFunction.index(dataFromDatabase(Code), 0, dataFieldInfoByEditedName("contract_units").ColumnIndex)  '
'            contractUnits = GetNumbers(contractUnits)
'
'            With notionalValuesByCode
'
'                .Add New Collection, Code
'
'                With .Item(Code)
'
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
'
'                End With
'
'            End With
'
'        Next
        
        'Calculate hedge ratio and combine.
        
        Dim contractToLong() As Variant, contractToShort() As Variant, iColumn As Byte, _
        hedgeRatio As Double, output() As Variant, nonCommLong As Byte, commLong As Byte, commShort As Byte, _
        nonCommShort As Byte, iShortRow As Long, iReduction As Long
        
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
                ' NET OI contract to long NonComm
                output(iRow, 2) = contractToLong(iRow, nonCommLong) - contractToLong(iRow, nonCommShort)
                ' NET OI contract to short NonComm
                output(iRow, 3) = contractToShort(iShortRow, nonCommLong) - contractToShort(iShortRow, nonCommShort)
                ' NET OI contract to long Comm
                output(iRow, 4) = contractToLong(iRow, commLong) - contractToLong(iRow, commShort)
                ' NET OI contract to Short Comm
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
            .Range(.Cells(1, 1), .Cells(.Rows.count, UBound(output, 2))).ClearContents
            .Resize(UBound(output, 1), UBound(output, 2)).Value2 = Reverse_2D_Array(output)
            ResizeTableBasedOnColumn outputTable, .columns(1)
            ClearRegionBeneathTable outputTable
            RestoreFilters outputTable, tableFilters
            
        End With
        
Finally:
        DisplayErr Err, "AttemptCross"
        Re_Enable
        
        Exit Sub
        
Exit_Loop:
        Resume PlaceOnSheet
    End Sub

#End If
