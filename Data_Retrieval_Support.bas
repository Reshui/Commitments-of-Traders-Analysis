Attribute VB_Name = "Data_Retrieval_Support"
Option Explicit


Public Sub Retrieve_Historical_Workbooks(ByRef Path_CLCTN As Collection, _
                                               ICE_Contracts As Boolean, _
                                               CFTC_Contracts As Boolean, _
                                               Mac_User As Boolean, _
                                               Report_Type As String, _
                                               combined_data_version As Boolean, _
                                            Optional ByVal CFTC_Start_Date As Date, _
                                            Optional ByVal CFTC_End_Date As Date, _
                                            Optional ByVal ICE_Start_Date As Date, _
                                            Optional ByVal ICE_End_Date As Date, _
                                            Optional ByVal Historical_Archive_Download As Boolean = False)

'======================================================================================================
'Downloads missing data from CFTC
'======================================================================================================

Dim fileNameWithinZip As String, Path_Separator As String, AnnualOF_FilePath As String, Destination_Folder As String, zipFileNameAndPath As String, _
fullFileName As String, multiYearFileExtractedFromZip As String, Partial_Url As String, URL As String, multiYearZipFileFullName As String, combinedOrFutures As String, Multi_Year_URL As String

Dim Queried_Date As Long, Download_Year As Integer, Final_Year As Integer, G As Byte, Contract_Data() As Variant

Const TXT As String = ".txt", ZIP As String = ".zip", CSV As String = ".csv", ID_String As String = "B.A.T"

Const mainFolderName As String = "COT_Historical_MoshiM"

On Error GoTo Failed_To_Download

#If Not Mac Then
    
    Path_Separator = Application.PathSeparator
    
    Destination_Folder = Environ("TEMP") & Path_Separator & mainFolderName & Path_Separator & Report_Type & Path_Separator & IIf(combined_data_version = True, "Combined", "Futures Only")
    
    If Not FileOrFolderExists(Destination_Folder) Then
        
        '/c =execute the command and then exit
        
        Shell ("cmd /c mkdir """ & Destination_Folder & """")
        
        Do Until FileOrFolderExists(Destination_Folder)
            
        Loop
    End If
    
#Else
    '/Users/rondebruin/Library/Containers/com.microsoft.Excel/Data
    
    Path_Separator = "/"
    
    Destination_Folder = BasicMacAvailablePathMac & Path_Separator & mainFolderName & Path_Separator & IIf(combined_data_version = True, "Combined", "Futures Only") 'Keep variable as an empty string.User will decide path
    
    If Not FileOrFolderExists(Destination_Folder) Then
        Call CreateRootDirectories(Destination_Folder)
    End If
    
#End If

With Path_CLCTN

    If CFTC_Contracts Then
    
        If Not combined_data_version Then  'IF Futures Only Workbook
        
            combinedOrFutures = "_Futures_Only"
            
            Select Case Report_Type
            
                Case "L"
                
                    fileNameWithinZip = "annual" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/deacot"

                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deacot1986_2016" & ZIP
                    
                    Contract_Data = Array("FUT86_16")
                    
                Case "D"
                
                    fileNameWithinZip = "f_year" & TXT
                    Partial_Url = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_"

                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_hist_2006_2016" & ZIP
                    
                    Contract_Data = Array("F_DisAgg06_16")
                    
                Case "T"
                
                    fileNameWithinZip = "FinFutYY" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/fut_fin_txt_"
                    
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_fut_txt_2006_2016" & ZIP
                    
                    Contract_Data = Array("F_TFF_2006_2016")
                    
            End Select
        
        Else 'Combined Contracts
        
            combinedOrFutures = "_Combined"
            
            Select Case Report_Type
            
                Case "L"
                
                    fileNameWithinZip = "annualof.txt"
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/deahistfo" 'TXT URL
                    
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deahistfo_1995_2016" & ZIP
                    
                    Contract_Data = Array("Com95_16")
                    
                Case "D"
                
                    fileNameWithinZip = "c_year" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/com_disagg_txt_"
                    'https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016.zip
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016" & ZIP
                    
                    Contract_Data = Array("C_DisAgg06_16")
                    
                Case "T"
                
                    fileNameWithinZip = "FinComYY" & TXT
                    'https://www.cftc.gov/files/dea/history/com_fin_txt_2014.zip
                    Partial_Url = "https://www.cftc.gov/files/dea/history/com_fin_txt_"
                    
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_com_txt_2006_2016" & ZIP
                    
                    Contract_Data = Array("C_TFF_2006_2016")
                    
            End Select
        
        End If
        
        If Year(CFTC_Start_Date) <= 2016 Then 'All report types have a compiled file for data before 2016
        
            Historical_Archive_Download = True
            
            CFTC_Start_Date = DateSerial(2017, 1, 1) 'So we can start dates in 2017 instead
            
        End If
        
        multiYearZipFileFullName = Destination_Folder & Path_Separator & Report_Type & "_COT_MultiYear_Archive" & combinedOrFutures & ZIP
        
        AnnualOF_FilePath = Destination_Folder & Path_Separator & fileNameWithinZip

        Download_Year = Year(CFTC_Start_Date)
        
        Final_Year = Year(CFTC_End_Date)
        
        Queried_Date = CFTC_End_Date
        
        For Download_Year = Download_Year - 1 To Final_Year '-1 is for if historical archive download needs to be executed
                
            If Not Historical_Archive_Download Then 'if not doing a download where multi year files are needed ie 2006-2016
            
                If Download_Year = Year(CFTC_Start_Date) - 1 Then
                    GoTo Skip_Download_Loop 'if on first loop
                Else
                    URL = Partial_Url & Download_Year & ZIP 'Declare URL of Zip file
                End If
                
            ElseIf Historical_Archive_Download Then
                
                URL = Multi_Year_URL
                
            End If
            
            For G = LBound(Contract_Data) To UBound(Contract_Data) 'loop at least once ,iterate through Historical strings if needed
            
                If Historical_Archive_Download Then
                    fullFileName = Destination_Folder & Path_Separator & Report_Type & "_" & Contract_Data(G) & combinedOrFutures & TXT
                ElseIf Final_Year = Download_Year Then
                    fullFileName = Destination_Folder & Path_Separator & Report_Type & "_Weekly_" & Queried_Date & "_" & Download_Year & combinedOrFutures & TXT
                Else
                    fullFileName = Destination_Folder & Path_Separator & Report_Type & "_" & Download_Year & combinedOrFutures & TXT
                End If
                
                If Not FileOrFolderExists(fullFileName) Then   'If wanted workbook doesn't exist
                    
                    If G = LBound(Contract_Data) Then 'Only need to check if Zip file exists once
                    
                        If Historical_Archive_Download Then
                            zipFileNameAndPath = multiYearZipFileFullName
                        Else
                            zipFileNameAndPath = Replace(fullFileName, TXT, ZIP)
                        End If
                        
                        If Not FileOrFolderExists(zipFileNameAndPath) Then

                            #If Mac Then
                                Call DownloadFile(URL, zipFileNameAndPath)
                            #Else
                                Call Get_File(URL, zipFileNameAndPath)
                            #End If
                        End If
                        
                        'Download Zip folder if it doesn't exist
                    End If
                    
                    If Not Historical_Archive_Download Then
                    
                        If FileOrFolderExists(AnnualOF_FilePath) Then Kill AnnualOF_FilePath    'If file within Zip folder exists within file directory then kill it
                    
                        #If Mac Then
                            Call UnzipFile(zipFileNameAndPath, Destination_Folder, fileNameWithinZip)
                        #Else
                            Call entUnZip1File(zipFileNameAndPath, Destination_Folder, fileNameWithinZip) 'Unzip specified file
                        #End If
                        
                        Name AnnualOF_FilePath As fullFileName
                        
                    ElseIf Historical_Archive_Download Then
                    
                        multiYearFileExtractedFromZip = Destination_Folder & Path_Separator & Contract_Data(G) & TXT
                        
                        If FileOrFolderExists(multiYearFileExtractedFromZip) Then Kill multiYearFileExtractedFromZip

                        #If Mac Then
                            Call UnzipFile(zipFileNameAndPath, Destination_Folder, Contract_Data(G) & TXT)
                        #Else
                            Call entUnZip1File(zipFileNameAndPath, Destination_Folder, Contract_Data(G) & TXT) 'Unzip specified file
                        #End If
                        
                        Name multiYearFileExtractedFromZip As fullFileName
                        
                    End If
                        
                End If
                
                .Add fullFileName, fullFileName
                
                If Not Historical_Archive_Download Then Exit For
                
            Next G

Skip_Download_Loop: If Historical_Archive_Download Then Historical_Archive_Download = False

        Next Download_Year
        
    End If
    
    If ICE_Contracts Then
        
        If Year(ICE_Start_Date) < 2011 Then
            ICE_Start_Date = DateSerial(2011, 1, 1)
        End If
        
        Download_Year = Year(ICE_Start_Date)
        Final_Year = Year(ICE_End_Date)
        
        Queried_Date = ICE_End_Date
        
        For Download_Year = Download_Year To Final_Year
        
            URL = "https://www.theice.com/publicdocs/futures/COTHist" & Download_Year & ".csv"
            
            Select Case Download_Year
                Case Final_Year
                    fullFileName = Destination_Folder & Path_Separator & "ICE_Weekly_" & Queried_Date & "_" & Download_Year & ".csv"
                Case Else
                    fullFileName = Destination_Folder & Path_Separator & "ICE_" & Download_Year & ".csv"
            End Select

            'ICE files are available as .csv via url
            If Not FileOrFolderExists(fullFileName) Then
                #If Mac Then
                    Call DownloadFile(URL, fullFileName)
                #Else
                    Call Get_File(URL, fullFileName)
                #End If
            End If
            
            .Add fullFileName, fullFileName

        Next Download_Year
        
    End If
    
End With

Exit Sub

Failed_To_Download:
    
    If Not UUID Then
    
        MsgBox "An error occured while downloading historical workbooks. Please check your internet connection and try again." & _
           vbCrLf & vbCrLf & "If you internet connection is fine, then report this error to MoshiM_UC@outlook.com"
           
           Re_Enable
           
           End
    Else
    
        Stop
        
        Resume Next
        
    End If
    
End Sub
Public Function Stochastic_Calculations(Column_Number_Reference As Integer, Time_Period As Integer, _
                                    arr As Variant, Optional Missing_Weeks As Integer = 1, Optional Cap_Extremes As Boolean = False) As Variant

Dim Array_Column() As Double, x As Integer, Array_Period() As Double, current_row As Integer, _
Array_Final() As Variant, UB As Integer, initialStartRow As Integer, Minimum As Double, Maximum As Double

UB = UBound(arr, 1) 'number of rows in the supplied array[Upper bound in 1st dimension]

initialStartRow = UB - (Missing_Weeks - 1) 'will be equal to UB if only 1 missing week

ReDim Array_Period(1 To Time_Period)  'Temporary Array that will hold a certain range of data

ReDim Array_Final(1 To Missing_Weeks) 'Array that will hold calculated values

ReDim Array_Column(IIf(initialStartRow > Time_Period, initialStartRow - Time_Period, 1) To UB) 'Array composed of all data in a column

If UBound(arr, 2) = 1 Then Column_Number_Reference = 1 'for when a single column is supplied

For x = IIf(initialStartRow > Time_Period, initialStartRow - Time_Period, 1) To UB
    'if starting row of data output is greater than the time period then offset the start of the queried array by the time period
    'otherwise start at 1...there should be checks to ensure there is enough data most of the time
    
    Array_Column(x) = arr(x, Column_Number_Reference)
    
Next x

For current_row = initialStartRow To UB

    If (current_row > Time_Period And Not Cap_Extremes) Or (current_row >= Time_Period And Cap_Extremes) Then   'Only calculate if there is enough data
    
        For x = 1 To Time_Period 'Fill array with the previous Time_Period number of values relative to the current row
            
            If Not Cap_Extremes Then
                Array_Period(x) = Array_Column(current_row - x)
            Else
                Array_Period(x) = Array_Column(current_row - (x - 1))
            End If
            
            If x = 1 Then
                Minimum = Array_Period(x)
                Maximum = Array_Period(x)
            ElseIf Array_Period(x) < Minimum Then
                Minimum = Array_Period(x)
            ElseIf Array_Period(x) > Maximum Then
                Maximum = Array_Period(x)
            End If
                            
        Next x
        'Stochastic calculation
        If Maximum <> Minimum Then
            Array_Final(Missing_Weeks - (UB - current_row)) = CByte(((Array_Column(current_row) - Minimum) / (Maximum - Minimum)) * 100)
        End If
        'ex for determining current location within array:    2 - ( 480 - 479 ) = 1
    End If

Next current_row

Stochastic_Calculations = Array_Final
    
End Function

Public Function Test_For_Data_Addition(Optional WKB As String) As Boolean

'======================================================================================================
'Tests if the CFTC has released new data based on its release calendar
'======================================================================================================

Dim Schedule_AR As Variant, Current_Date As Date, Last_Release As Date

On Error GoTo Default_True

Last_Release = CFTC_Release_Dates(True) 'Returns Local date and time for the most recent release

If WKB <> vbNullString And CDbl(Last_Release) <> 0 Then
    
    If FileDateTime(WKB) < Last_Release Then
        Test_For_Data_Addition = True 'If data has been updated since
    End If                            'the Compilation workbook was made...THEN return True
    
Else

   Debug.Print Last_Release
   Test_For_Data_Addition = True 'default to True [Recreate workbook]
   
End If

Exit Function

Default_True:
    Test_For_Data_Addition = True
    
End Function

Public Function Multi_Week_Addition(My_CLCTN As Collection, Sort_Type As Byte) As Variant 'adds the contents of the NEW array TO the contents of the OLD
  
  
'======================================================================================================
'Combines multiple 1D and 2D arrays
'======================================================================================================
  
Dim finalColumnIndex As Integer, x As Long, finalRowIndex As Long, UB1 As Long, UB2 As Byte, Worksheet_Data() As Variant, _
Item As Variant, Old() As Variant, Block() As Variant, Latest() As Variant, Not_Old As Byte, Is_Old As Byte
   
'Dim Addition_Timer As Double: Addition_Timer = Time

With My_CLCTN
    'check the boundaries of the elements to create the array
    Select Case Sort_Type

        Case Append_Type.Multiple_1d 'Array Elements are 1D | single rows |  "Historical_Parse"

            UB1 = .Count 'The number of items in the dictionary will be the number of rows in the final array

            For Each Item In My_CLCTN 'loop through each item in the row and find the max number of columns
                
                finalRowIndex = UBound(Item) + 1 - LBound(Item) 'Number of Columns if 1 based
            
                If finalRowIndex > UB2 Then UB2 = finalRowIndex
                
            Next Item
            
        Case Append_Type.Multiple_2d 'Indeterminate number of  2D[1-Based] arrays to be joined.
                              '"Historical_Excel_Creation"
            For Each Item In My_CLCTN
                
                UB1 = UBound(Item, 1) + UB1
                
                finalRowIndex = UBound(Item, 2)
                
                If finalRowIndex > UB2 Then UB2 = finalRowIndex
                
            Next Item

        Case Append_Type.Add_To_Old
            
            Not_Old = 1
            
            Do Until .Item(Not_Old)(0) <> Data_Identifier.Old_Data
                Not_Old = Not_Old + 1
            Loop
            
            '3 mod not_old
            
            Is_Old = IIf(Not_Old = 1, 2, 1)
            
            Old = .Item(Is_Old)(1)
                
            finalRowIndex = UBound(Old, 2)
            
            Select Case .Item(Not_Old)(0)         'Number designating array type
            
                Case Data_Identifier.Weekly_Data  'This key is used for when sotring weekly data
                
                    Latest = .Item(Not_Old)(1)
                    
                    finalColumnIndex = UBound(Latest)              'Number of columns in the 1-based 1D array
                    
                    UB1 = UBound(Old, 1) + 1 ' +1 Since there will be only 1 row of additional weekly data
                
                Case Data_Identifier.Block_Data  'This key is used if several weeks have passed
                                                        'This will be a 2d array
                    Block = .Item(Not_Old)(1)
                    
                    finalColumnIndex = UBound(Block, 2)
                    
                    UB1 = UBound(Old, 1) + UBound(Block, 1)
                
            End Select
            
            If finalRowIndex >= finalColumnIndex Then 'Determing number of columns to size the array with
                UB2 = finalRowIndex    'S= # of Columns in the older data
            Else           'T= # of Columns in the new data
                UB2 = finalColumnIndex
            End If

    End Select
    
    ReDim Worksheet_Data(1 To UB1, 1 To UB2)
    
    finalRowIndex = 1
    
    For Each Item In My_CLCTN
        
        Select Case Sort_Type

            Case Append_Type.Multiple_1d 'All items in Collection are 1D
                
                For finalColumnIndex = LBound(Item) To UBound(Item)
                    
                    Worksheet_Data(finalRowIndex, finalColumnIndex + 1 - LBound(Item)) = Item(finalColumnIndex)

                Next finalColumnIndex

                finalRowIndex = finalRowIndex + 1
                
            Case Append_Type.Multiple_2d 'Adding Multiple 2D arrays together
    
                    x = 1
                    
                    For finalRowIndex = finalRowIndex To UBound(Item, 1) + (finalRowIndex - 1)

                        For finalColumnIndex = LBound(Item, 2) To UBound(Item, 2)

                            Worksheet_Data(finalRowIndex, finalColumnIndex) = Item(x, finalColumnIndex)
                            
                        Next finalColumnIndex
                        
                        x = x + 1
                    
                    Next finalRowIndex
            
            Case Append_Type.Add_To_Old 'Adding new Data to a 2D Array..Block is 2D...Latest is 1D
                                        
                Select Case Item(0)                 'Key of item

                    Case Data_Identifier.Old_Data 'Current Historical data on Worksheet
                        
                        For finalRowIndex = LBound(Worksheet_Data, 1) To UBound(Old, 1)

                            For finalColumnIndex = LBound(Old, 2) To UBound(Old, 2)

                                Worksheet_Data(finalRowIndex, finalColumnIndex) = Old(finalRowIndex, finalColumnIndex)

                            Next finalColumnIndex
                            
                        Next finalRowIndex
                        
                    Case Data_Identifier.Block_Data '<--2D Array used when adding to arrays together where order is important
                    
                        x = 1
                        
                        For finalRowIndex = UBound(Worksheet_Data, 1) - UBound(Block, 1) + 1 To UBound(Worksheet_Data, 1)

                            For finalColumnIndex = LBound(Block, 2) To UBound(Block, 2)

                                Worksheet_Data(finalRowIndex, finalColumnIndex) = Block(x, finalColumnIndex)
                                
                            Next finalColumnIndex
                            
                            x = x + 1
                        
                        Next finalRowIndex
                        
                    Case Data_Identifier.Weekly_Data  '1 based 1D "WEEKLY" array
                                      '"OLD" is run first so S is already at the correct incremented value
                        
                        finalRowIndex = UBound(Worksheet_Data, 1)
                        
                        For finalColumnIndex = LBound(Latest) To UBound(Latest)
                            
                            Worksheet_Data(finalRowIndex, finalColumnIndex) = Latest(finalColumnIndex) 'worksheet data is 1 based 2D while Item is 1 BASED 1D

                        Next finalColumnIndex
                                      
                End Select

        End Select
        
    Next Item

End With

Multi_Week_Addition = Worksheet_Data
    
'Debug.Print Timer - Addition_Timer

End Function
Public Function HTTP_Weekly_Data(Last_Update As Long, Report_Type As String, Combined_Version As Boolean, Optional Auto_Retrieval As Boolean = False, _
                                Optional DebugMD As Boolean = False) As Variant

'Functinon must be called exactly once per operation to avoid errors like erasing the Valid_Table_Info array

Dim PowerQuery_Available As Boolean, Power_Query_Failed As Boolean, _
Text_Method_Failed As Boolean, Query_Table_Method_Failed As Boolean, MAC_OS As Boolean, Retrieval_Successful As Boolean

Dim TimedTasks As TimerC

Const PowerQTask As String = "Power Query Retrieval", QueryTask As String = "QueryTable Retrieval", HTTPTask As String = "HTTP Retrieval"

If DebugMD Then Set TimedTasks = New TimerC

#If Mac Then

    MAC_OS = True
    PowerQuery_Available = False 'Use standalone QueryTable rather than QueryTable wrapped in listobject

#Else
    
    On Error GoTo Default_No_Power_Query
    
    If Val(Application.Version) < 16 Then 'IF excel version is prior to Excel 2016 then

        PowerQuery_Available = IsPowerQueryAvailable 'Check if Power Query is available
    Else
        PowerQuery_Available = True
    End If

#End If

Retrieval_Process:

If Retrieval_Successful = False Or DebugMD Then

    On Error GoTo QueryTable_Failed
    
    If DebugMD Then TimedTasks.StartTask QueryTask
        
    HTTP_Weekly_Data = CFTC_Data_QueryTable_Method(Report_Type:=Report_Type, combined_wb:=Combined_Version)
    
    If DebugMD Then
        TimedTasks.EndTask QueryTask
    Else
        Exit Function
    End If
    
End If

If Not MAC_OS Then

PowerQuery_Method:

    If PowerQuery_Available Or DebugMD Then
    
        On Error GoTo PowerQuery_Failed
        
        If DebugMD Then TimedTasks.StartTask PowerQTask
            
        HTTP_Weekly_Data = CFTC_Data_PowerQuery_Method(Report_Type, Combined_Version)
            
        If DebugMD Then
            TimedTasks.EndTask PowerQTask
        Else
            Exit Function
        End If
    
    End If
    
TXT_Method:

    If DebugMD Or IsEmpty(HTTP_Weekly_Data) Then     'TXT file Method
    
        On Error GoTo TXT_Failed
        
        If DebugMD Then TimedTasks.StartTask HTTPTask
            
        HTTP_Weekly_Data = CFTC_Data_Text_Method(Last_Update, Report_Type:=Report_Type, Combined_Version:=Combined_Version)
            
        If DebugMD Then
            TimedTasks.EndTask HTTPTask
        Else
            Exit Function
        End If
        
    End If

End If
                                                                                                                  
Exit_Code:

If DebugMD Then Debug.Print TimedTasks.ToString

If Not Retrieval_Successful And Not Auto_Retrieval Then

    MsgBox "Data retrieval has failed." & vbNewLine & vbNewLine & _
           "If you are on Windows and Power Query is available and you aren't using Excel 2016 then please enable or download Power Query / Get and Transform and try again." & vbNewLine & vbNewLine & _
           "If you are on a MAC or the above step fails then please contact me at MoshiM_UC@outlook.com with your operating system and Excel version."
           
End If

Exit Function

PowerQuery_Failed:
    If DebugMD Then TimedTasks.EndTask PowerQTask
    Resume TXT_Method
    
TXT_Failed:
    If DebugMD Then TimedTasks.EndTask HTTPTask
    Resume Exit_Code
    
QueryTable_Failed:
    
    If Not MAC_OS Then
        Resume PowerQuery_Method
    Else
        Resume Exit_Code
    End If
    
Default_No_Power_Query:

    PowerQuery_Available = False
    Resume Retrieval_Process
    
End Function
'

Public Function CFTC_Data_PowerQuery_Method(Report_Type As String, Combined_Version As Boolean) As Variant
    
'======================================================================================================
'Retrieves the latest Weekly data with Power Query. Only use if on Windows.
'======================================================================================================
    Dim URL As String, Formula_AR() As String, quotation As String, Y As Byte, table_name As String
    
    table_name = Evaluate("=VLOOKUP(""" & Report_Type & """,Report_Abbreviation,2,FALSE)")
    
    quotation = Chr(34)
    
    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(Report_Type, Array("L", "D", "T"), 0) - 1
    
    If Not Combined_Version Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    With ThisWorkbook
        'Change Query URL
        Formula_AR = Split(.Queries(table_name).Formula, quotation, 3) 'Split with quotation mark
        Formula_AR(1) = URL
        .Queries(table_name).Formula = Join(Formula_AR, quotation)
        
    End With

    With Weekly.ListObjects(table_name)
        .QueryTable.Refresh False                               'Refresh Weekly Data Table
        CFTC_Data_PowerQuery_Method = .DataBodyRange.Value2     'Store contents of table in an array
    End With
    
End Function
Public Function CFTC_Data_Text_Method(Last_Update As Long, Report_Type As String, Combined_Version As Boolean) As Variant

Dim File_Path As New Collection, URL As String, Y As Byte

    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(Report_Type, Array("L", "D", "T"), 0) - 1
    
    If Not Combined_Version Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    With File_Path
    
        .Add Environ("TEMP") & "\" & Date & "_" & Report_Type & "_Weekly.txt", "Weekly Text File" 'Add file path of file to be downloaded
    
        Call Get_File(URL, .Item("Weekly Text File")) 'Download the file to the above path
        
    End With
    
    CFTC_Data_Text_Method = Historical_Parse(File_Path, Combined_Version:=Combined_Version, CFTC_TXT:=True, Report_Type:=Report_Type, After_This_Date:=Last_Update) 'return array
    
End Function
Public Function CFTC_Data_QueryTable_Method(Report_Type As String, combined_wb As Boolean) As Variant

Dim Data_Query As QueryTable, Data() As Variant, URL As String, _
 Y As Byte, BB As Boolean, Variables, _
Found_Data_Query As Boolean, Error_While_Refreshing As Boolean

Dim Workbook_Type As String

With Application

    BB = .EnableEvents
    
    .EnableEvents = False
    .DisplayAlerts = False
    
End With

Workbook_Type = IIf(combined_wb, "Combined", "Futures_Only")

For Each Data_Query In QueryT.QueryTables
    If InStr(1, Data_Query.name, Report_Type & "_CFTC_Data_Weekly_" & Workbook_Type) > 0 Then
        Found_Data_Query = True
        Exit For
    End If
Next Data_Query

If Not Found_Data_Query Then 'If QueryTable isn't found then create it

Recreate_Query:
    
    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(Report_Type, Array("L", "D", "T"), 0) - 1
    
    If Not combined_wb Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    Set Data_Query = QueryT.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=QueryT.Range("A1"))
    
    With Data_Query
        
        .BackgroundQuery = False
        .SaveData = False
        .AdjustColumnWidth = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        
        .TextFileColumnDataTypes = Filter_Market_Columns(convert_skip_col_to_general:=True, Report_Type:=Report_Type, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True)
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        
        .name = Report_Type & "_CFTC_Data_Weekly_" & Workbook_Type
        
        On Error GoTo Delete_Connection
        
Name_Connection:

        With .WorkbookConnection
            .RefreshWithRefreshAll = False
            .name = Report_Type & "_Weekly CFTC Data: " & Workbook_Type
        End With
        
    End With
    
    On Error GoTo 0

End If

On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time

With Data_Query

    .Refresh False
    
    With .ResultRange
        .Replace ".", Null, xlWhole
        CFTC_Data_QueryTable_Method = .Value2 'Store Data in an Array
        .ClearContents 'Clear the Range
    End With
    
End With

With Application
    .DisplayAlerts = True
    .EnableEvents = BB
End With

Exit Function

Delete_Connection: 'Error handler is available when editing parameters for a new querytable and the connection name is already taken by a different query

    ThisWorkbook.Connections("Weekly CFTC Data: " & Workbook_Type).Delete
        
    On Error GoTo 0
    
    Resume Name_Connection
    
Failed_To_Refresh:
        
    Data_Query.Delete
    
    If Error_While_Refreshing = True Then
        On Error GoTo 0
        Err.Raise 5
    Else
        Error_While_Refreshing = True
        Resume Recreate_Query
    End If
    
End Function
Public Function Historical_Parse(ByVal File_CLCTN As Collection, Report_Type As String, Combined_Version As Boolean, _
                                  Optional ByRef contract_code As String = vbNullString, _
                                  Optional After_This_Date As Long = 0, _
                                  Optional Kill_Previous_Workbook As Boolean = False, _
                                  Optional Yearly_C As Boolean, _
                                  Optional Specified_Contract As Boolean, _
                                  Optional Weekly_ICE_Data As Boolean, _
                                  Optional CFTC_TXT As Boolean, _
                                  Optional Parse_All_Data As Boolean) As Variant

    Dim Date_Sorted As New Collection, Item As Variant, Escape_Filter_Market_Arrays As Boolean, _
    Contract_WB As Workbook, Contract_WB_Path As String, ICE_Data As Boolean, Mac_UserB As Boolean
    
    Dim ErrorC As Collection, OS_BasedPathSeparator As String
    
    On Error GoTo Historical_Parse_General_Error_Handle
    
    #If Mac Then
        OS_BasedPathSeparator = "/"
        Mac_UserB = True
    #Else
        OS_BasedPathSeparator = "\"
    #End If
    
    With ThisWorkbook
    
        If Not HasKey(.Event_Storage, "Historical Parse Errors") Then
            .Event_Storage.Add New Collection, "Historical Parse Errors"
        End If
        
        Set ErrorC = .Event_Storage("Historical Parse Errors")
        
    End With
    
    Application.ScreenUpdating = False

Select Case True

    Case Yearly_C, Specified_Contract, Parse_All_Data 'Parse all data is for when all data is being downloaded

        Contract_WB_Path = Left$(File_CLCTN(1), InStrRev(File_CLCTN(1), OS_BasedPathSeparator)) 'Folder path of File_CLCTN(1)
    
        If Yearly_C Then 'If Yearly Contracts ie: only missing a chunk of data
        
            Contract_WB_Path = Contract_WB_Path & Report_Type & "_COT_Yearly_Contracts_" & IIf(Combined_Version, "Combined", "Futures_Only") & ".xlsb"
            
        ElseIf Specified_Contract Or Parse_All_Data Then  'If using the new contract macro
        
            Contract_WB_Path = Contract_WB_Path & Report_Type & "_COT_Historical_Archive_" & IIf(Combined_Version, "Combined", "Futures_Only") & ".xlsb"
            
        End If
    
        If Not FileOrFolderExists(Contract_WB_Path) Then 'If the file doesn't exist then compile the text files into a single document and create a workbook from it
            
            Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, Report_Type:=Report_Type, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=Mac_UserB, combined_wb:=Combined_Version)
            
        Else 'If the file exists

            If Test_For_Data_Addition(Contract_WB_Path) Or Yearly_C Or Kill_Previous_Workbook = True Then 'if new data has been added since last workbook was created
                
                On Error Resume Next

                Kill Contract_WB_Path 'Kill and then recreate
                On Error GoTo 0
                
                Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, Report_Type:=Report_Type, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=Mac_UserB, combined_wb:=Combined_Version)
                
            Else
                
                Set Contract_WB = Workbooks.Open(Contract_WB_Path)  'Set a reference
                Contract_WB.Windows(1).Visible = False

            End If

        End If
        
        For Each Item In File_CLCTN
            If Item Like "*ICE*" Then
                ICE_Data = True
                Exit For
            End If
        Next Item
        
        Historical_Parse = Historical_Excel_Aggregation(Contract_WB, combined_workbook:=Combined_Version, Contract_ID:=contract_code, Date_Input:=After_This_Date, Specified_Contract:=Specified_Contract, ICE_Contracts:=ICE_Data)
        
        Contract_WB.Close False 'Close without saving
        
        ICE_Data = False 'Workbook Structure has been homogenized
        
    Case Weekly_ICE_Data 'Result=2D Array stored in Collection, Array isn't filtered
        
        ICE_Data = True
        
        Set Contract_WB = Workbooks.Open(File_CLCTN.Item("ICE"))
        
        With Contract_WB
        
            .Windows(1).Visible = False
        
            Historical_Parse = Historical_Excel_Aggregation(Contract_WB, combined_workbook:=Combined_Version, Date_Input:=After_This_Date, ICE_Contracts:=True)
            
            .Close False
            
            If Combined_Version = False Then Kill File_CLCTN.Item("ICE")
            
        End With
        
    Case CFTC_TXT 'Result=2D Array stored in Collection2D Array(s) stored in Collection from .txt file(s)

        Historical_Parse = Weekly_Text_File(File_CLCTN, Report_Type:=Report_Type, Date_Value:=After_This_Date, Combined_Version:=Combined_Version)
        
End Select

Application.StatusBar = vbNullString

Exit Function

Historical_Parse_General_Error_Handle:

    If CFTC_TXT Or Weekly_ICE_Data Then  'use parent error handler
    
        On Error GoTo 0
        Err.Raise 5
        
    ElseIf Yearly_C Or Specified_Contract Or Parse_All_Data Then
    
        Contract_WB_Path = "An error has occured while running the Historical Parse subroutine. Please email me at MoshiM_UC@outlook.com"
        
        For Each Item In ErrorC
            Contract_WB_Path = Contract_WB_Path & vbNewLine & vbNewLine & Item & vbNewLine & vbNewLine
        Next Item
        
        MsgBox Contract_WB_Path
        
        ThisWorkbook.Event_Storage.Remove "Historical Parse Errors"
        Set ErrorC = Nothing
        
        Re_Enable
        
        End
    
    End If
    
End Function
Public Function Historical_TXT_Compilation(File_Collection As Collection, _
Saved_Workbook_Path As String, OnMAC As Boolean, Report_Type As String, combined_wb As Boolean) As Workbook
    
Dim File_TXT As Variant, FileNumber As Long, Data_STR As String, File_Path() As String

Dim InfoF() As Variant, FilterC As Variant, D As Long, ICE_Filter As Boolean, ICE_Count As Byte, OS_BasedPathSeparator As String

Dim File_Name As String, CFTC_Count As Byte, file_text As String, output_file_number As Long, output_file_name As String 'g ', DD As Double

Const comma As String = ","

On Error GoTo Query_Table_Method_For_TXT_Retrieval
    
If OnMAC Then
    OS_BasedPathSeparator = "/"
Else
    OS_BasedPathSeparator = "\"
End If
    
    
    
    FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, Report_Type:=Report_Type, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=False)
    '^^ retrieve wanted column NUmbers

    ReDim InfoF(1 To UBound(FilterC, 1))
    
    For D = 1 To UBound(FilterC, 1) 'Fill in column numbers for use when supplying column filters to OpenTxt
        InfoF(D) = Array(D, FilterC(D))
    Next D
    
    Erase FilterC
    
    output_file_number = FreeFile
    
    output_file_name = Left$(File_Collection(1), InStrRev(File_Collection(1), OS_BasedPathSeparator)) & "Historic.txt"
    
    If FileOrFolderExists(output_file_name) Then Kill output_file_name
    
    Open output_file_name For Append As output_file_number 'Write contents of string to text File
    
    FileNumber = FreeFile
    
    For Each File_TXT In File_Collection 'Open each file in the collection and write their contents to a string
    
        Application.StatusBar = "Parsing " & File_TXT
        DoEvents
        
        Open File_TXT For Input As FileNumber
            
        File_Name = Right$(File_TXT, Len(File_TXT) - InStrRev(File_TXT, OS_BasedPathSeparator))
        
        If File_Name Like "*ICE*" Then 'IF name has ICE in it
        
            D = 0
            ICE_Count = ICE_Count + 1
            Do Until EOF(FileNumber)
                
                D = D + 1
                Line Input #FileNumber, Data_STR
                
                If (D > 1 And ICE_Count > 1) Or ICE_Count = 1 Then
                    'Only allow printing of headers if on first file
                    Print #output_file_number, Data_STR
                End If
                
            Loop
            
        Else
            
            CFTC_Count = CFTC_Count + 1
            D = 0
            Do Until EOF(FileNumber)
                
                D = D + 1
                Line Input #FileNumber, Data_STR
                
                If (D > 1 And CFTC_Count > 1) Or CFTC_Count = 1 Then
                    'Only allow printing of headers if on first file
                    Print #output_file_number, Data_STR
                End If
                
            Loop
            
        End If
            
        Close FileNumber
        
        'If LCase(File_TXT) Like "*weekly*" Then Kill File_TXT
        
    Next File_TXT
    
On Error GoTo Query_Table_Method_For_TXT_Retrieval

    Close output_file_number
    
    Application.StatusBar = "TXT file compilation was successful. Creating Workbook."
    DoEvents
       
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If
    
    With Workbooks
    
            .OpenText Filename:=output_file_name, origin:=D, StartRow:=1, DataType:=xlDelimited, _
                                    TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, comma:=True, _
                                    FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                                    Local:=False
                                    
        Set Historical_TXT_Compilation = Workbooks(.Count)
        
        'Contract_WB.Windows(1).Visible = False
        
    End With
    
   Historical_TXT_Compilation.Windows(1).Visible = False 'True
    
    On Error Resume Next
        If Not OnMAC Then Historical_TXT_Compilation.SaveAs Saved_Workbook_Path, FileFormat:=xlExcel12
    On Error GoTo 0
        
'ElseIf OnMAC Then

Exit Function
    
Query_Table_Method_For_TXT_Retrieval:
    
    On Error GoTo -1
    
    On Error GoTo Parent_Handler

    InfoF = Query_Text_Files(File_Collection, combined_wb:=combined_wb, Report_Type:=Report_Type) 'Use Querytables
    
    Application.StatusBar = "Data compilation was successful. Creating Workbook."
    DoEvents
    
    Set Historical_TXT_Compilation = Workbooks.Add
    
    With Historical_TXT_Compilation
    
        .Windows(1).Visible = False
        
        With .Worksheets(1)
            .DisplayPageBreaks = False
            .Columns("C:C").NumberFormat = "@" 'Format as text
            .Range("A1").Resize(UBound(InfoF, 1), UBound(InfoF, 2)).Value2 = InfoF
        End With
        
    End With
    
    Exit Function
    
Parent_Handler:

    ThisWorkbook.Event_Storage("Historical Parse Errors").Add "An error occurred while compiling text files."
    Resume Exit_SC
    
Exit_SC:
    
    On Error GoTo 0

    Err.Raise 5

End Function
Public Function Historical_Excel_Aggregation(Contract_WB As Workbook, _
                                        combined_workbook As Boolean, Optional Contract_ID As String, _
                                        Optional Date_Input As Long = 0, _
                                        Optional ICE_Contracts As Boolean = False, _
                                        Optional Specified_Contract As Boolean = False, _
                                        Optional Weekly_CFTC_TXT As Boolean = False, Optional QueryTable_To_Filter As Variant) As Variant

'======================================================================================================
'Filters and sorts data on worksheet 1 of a given workbook
'Generated array will be stored in a collection
'======================================================================================================

Dim VAR_DTA() As Variant, Comparison_Operator As String, _
Table_OBJ As ListObject, DBR As Range, Z As Long  ', TT As Double

Dim Combined_CLMN As Byte, Disaggregated_Filter_STR As String 'Used if filtering ICE Contracts for Futures and Options

Dim Error_Number As Long, Filtering_QueryTable As Boolean, Source_RNG As Range, WS As Worksheet

Const yymmdd_column As Byte = 2
Const Contract_Code_CLMN As Byte = 4 'Column that holds Contract identifiers
Const ICE_Contract_Code_CLMN As Byte = 7
Const Date_Field As Byte = 3

'TT = Timer
On Error GoTo Close_Workbook

'Err.Raise 5

Filtering_QueryTable = Not IsMissing(QueryTable_To_Filter)

If Not Filtering_QueryTable Then
    Application.StatusBar = "Filtering Data."
    DoEvents
    Set WS = Contract_WB.Worksheets(1)
Else
    Set WS = QueryTable_To_Filter.Parent
End If

With WS
    
    If .UsedRange.Cells.Count = 1 Then 'If worksheet is empty then display message
        GoTo Scripts_Failed_To_Collect_Data
    Else
        
        'If .ListObjects.Count = 0 Or Filtering_QueryTable Then
        
            If Weekly_CFTC_TXT Then 'Determine if Worksheet has headers based on if its a Text Document or not
                Z = xlNo
            Else
                Z = xlYes
            End If
            
            If Not Filtering_QueryTable Then
                Set Source_RNG = .UsedRange
            Else
                Set Source_RNG = QueryTable_To_Filter.ResultRange
            End If
            
'            Set Table_OBJ = .ListObjects.Add(SourceType:=xlSrcRange, Source:=Source_RNG, XlListObjectHasHeaders:=Z)
'
'        Else
'            Set Table_OBJ = .ListObjects(1)
'        End If
        
    End If
    
End With
 
If ICE_Contracts Then
    Disaggregated_Filter_STR = IIf(combined_workbook, "*Combined*", "*FutOnly*")
End If

On Error GoTo Close_Workbook

Set DBR = Source_RNG

With DBR
    
    'Set DBR = .DataBodyRange

Check_If_Code_Exists:
    
    If ICE_Contracts Then 'Find a column to be sorted based on the column header
    
        Combined_CLMN = Application.Match("FutOnly_or_Combined", .Rows(1).Value2, 0)
        
    ElseIf Specified_Contract Then 'Store filter information for wanted Contract Code
                                                
        VAR_DTA = Array(Contract_Code_CLMN, UCase(Contract_ID), xlFilterValues, False)
        
    End If
    
    If ICE_Contracts Or Weekly_CFTC_TXT Then 'Weekly_CFTC_TXT should be unique to CFTC Weekly Text Files at the time of writing
        Comparison_Operator = ">="
    Else
        Comparison_Operator = ">"
    End If
    
    If ICE_Contracts Then 'Yearly ICE has already been converted when creating the Excel File
    
        Comparison_Operator = Comparison_Operator & Format(IIf(Date_Input = 0, DateSerial(2000, 1, 1), Date_Input), "YYMMDD") 'Format(Year(Date_Input) - 2000, "00") & Format(Month(Date_Input), "00") & Format(Day(Date_Input), "00")
    Else
        Comparison_Operator = Comparison_Operator & Date_Input
        
    End If
    
    On Error Resume Next
    
    .Parent.ShowAllData
    
    On Error GoTo Close_Workbook
    
    .Sort key1:=DBR.Cells(2, IIf(ICE_Contracts = True, yymmdd_column, Date_Field)), ORder1:=xlAscending, header:=Z, MatchCase:=False

'    With .Sort 'Sort Date Field Old to New
'
'        With .SortFields
'            .Clear
'            .Add Key:=DBR.Cells(2, IIf(ICE_Contracts = True, yymmdd_column, Date_Field)), SortOn:=xlSortOnValues, Order:=xlAscending
'        End With
'
'        .MatchCase = False
'        .Header = xlYes
'        .Apply
'
'    End With
    'Filter Date Field
    
    .AutoFilter Field:=IIf(ICE_Contracts = True, yymmdd_column, Date_Field), Criteria1:=Comparison_Operator, Operator:=xlFilterValues
    
    If ICE_Contracts Then 'Sort by Combined Contracts or Futures Only disaggregated report
        
        .Sort key1:=DBR.Cells(2, Combined_CLMN), ORder1:=xlAscending, header:=xlYes, MatchCase:=False

'        With .Sort 'If ICE contracts then group
'                   'Group by contract Codes currently in this workbook
'            With .SortFields
'                .Clear
'                .Add Key:=DBR.Cells(2, Combined_CLMN), SortOn:=xlSortOnValues, Order:=xlAscending
'            End With
'
'            .MatchCase = False
'            .Header = xlYes
'            .Apply
'
'        End With
    
    End If
    
    With DBR 'Filter for "Combined" if condition met. Filter for wanted contract(s)
    
        If ICE_Contracts Then .AutoFilter Field:=Combined_CLMN, Criteria1:=Disaggregated_Filter_STR, Operator:=xlFilterValues, VisibleDropDown:=False

        If Specified_Contract Then
        
            .AutoFilter Field:=VAR_DTA(0), _
                        Criteria1:=VAR_DTA(1), _
                        Operator:=VAR_DTA(2), _
                        VisibleDropDown:=VAR_DTA(3)
        End If
        
        With .SpecialCells(xlCellTypeVisible)
            
            If .Cells.Count > 1 Then
            
                If Weekly_CFTC_TXT Then
                    VAR_DTA = .Value2
                Else
                
                    If .Areas.Count = 1 Then
                        VAR_DTA = .Offset(1).Resize(.Rows.Count - 1).Value2
                    Else
                        VAR_DTA = .Areas(2).Value2
                    End If
                    
                End If
                
                If ICE_Contracts Then  'Convert Dates from YYMMDD
                
                    For Z = LBound(VAR_DTA, 1) To UBound(VAR_DTA, 1)
                        
                        If IsEmpty(VAR_DTA(Z, Contract_Code_CLMN)) Then
                            VAR_DTA(Z, Date_Field) = DateSerial(Left(VAR_DTA(Z, yymmdd_column), 2) + 2000, Mid(VAR_DTA(Z, yymmdd_column), 3, 2), Right(VAR_DTA(Z, yymmdd_column), 2))
                            VAR_DTA(Z, Contract_Code_CLMN) = VAR_DTA(Z, ICE_Contract_Code_CLMN)
                            VAR_DTA(Z, ICE_Contract_Code_CLMN) = Empty
                        End If
                        
                    Next Z
                    
                End If
            
                Historical_Excel_Aggregation = VAR_DTA
                
                'Erase VAR_DTA
                
            End If
            
        End With 'End .SpecialCells(xlCellTypeVisible)
        
    End With 'End DBR
    
End With 'End Table_OBJ

If Not Filtering_QueryTable Then

    With Application
        .StatusBar = vbNullString
        DoEvents
    End With

End If

'Debug.Print Timer - TT

Exit Function

Close_Workbook: 'Error handler

    If Not Contract_WB Is ThisWorkbook Then
    
        Contract_ID = Contract_WB.FullName
        Contract_WB.Close False
        
        On Error Resume Next
        
        Kill Contract_ID
        
        With Application
            .StatusBar = ""
        End With
        
        ThisWorkbook.Event_Storage("Historical Parse Errors").Add "Error during Historical Filtration function."
                
        Error_Number = Err.Number
        
    End If
    
    Resume Parent_Error_Handler
    
Contract_ID_Not_Found: 'Used when user has input an invalid contract code

    If MsgBox("The Selected Contract Code wasn't found" & vbNewLine & "Would you like to try again with a different Contract Code?", vbYesNo, "Please choose") _
                = vbYes Then
        Contract_ID = UCase(Application.InputBox("Please supply the Contract Code of the desired contract"))

        GoTo Check_If_Code_Exists

    Else
        Application.StatusBar = vbNullString:
        Contract_WB.Close False
        Call Re_Enable
        End 'EXITS ALL CODE Execution
    End If
    
Scripts_Failed_To_Collect_Data:
    
    ThisWorkbook.Event_Storage("Historical Parse Errors").Add _
        "Error:  No data found on worksheet." & vbCrLf & vbCrLf & _
        "Subroutine: ""Historical_Excel_Aggregation""" & vbCrLf & _
        "File name: " & Contract_WB.name
        
    Contract_ID = Contract_WB.FullName
    Contract_WB.Close False
    
    Error_Number = Err.Number
    
    On Error Resume Next
        Kill Contract_ID

    Resume Parent_Error_Handler

Parent_Error_Handler:

    On Error GoTo 0
    
    Err.Raise Error_Number 'Enter historical parse error handler
    
End Function
Public Function Weekly_Text_File(File_Path As Collection, Date_Value As Long, Report_Type As String, Combined_Version As Boolean) As Variant

Dim File_IO As Variant, D As Byte, FilterC() As Variant, InfoF() As Variant, Contract_WB As Workbook

FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, Return_Filter_Columns:=True, Report_Type:=Report_Type, Return_Filtered_Array:=False, Create_Filter:=True)

ReDim InfoF(1 To UBound(FilterC, 1))

For D = 1 To UBound(FilterC, 1)
    InfoF(D) = Array(D, FilterC(D))
Next D

Erase FilterC

#If Mac Then
    D = xlMacintosh
#Else
    D = xlWindows
#End If

For Each File_IO In File_Path

    On Error GoTo Error_While_Opening_Text_File

    With Workbooks
    
        .OpenText Filename:=File_IO, origin:=D, StartRow:=1, DataType:=xlDelimited, _
                            TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, comma:=True, _
                            FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                            Local:=False
                       
        Set Contract_WB = Workbooks(.Count)
    
    End With
    
    With Contract_WB
        
        .Windows(1).Visible = False
        
        On Error GoTo Workbook_Parse_Error
        
        Weekly_Text_File = Historical_Excel_Aggregation(Contract_WB, Date_Input:=Date_Value, Weekly_CFTC_TXT:=True, combined_workbook:=Combined_Version)
    
        .Close False
        
    End With
    
    Kill File_IO

Next_File:

Next File_IO

Exit Function

Error_While_Opening_Text_File:

    ThisWorkbook.Event_Storage("Historical Parse Errors").Add "Error while attempting to open a Weekly based Text File"
        
    On Error Resume Next
    
        Kill File_IO
    
    Resume Enact_Handler
    
Workbook_Parse_Error:
    
    On Error Resume Next
    
        Contract_WB.Close False
        Kill File_IO
        
    Resume Enact_Handler
    
Enact_Handler:
    
    On Error GoTo 0 'Script should then use the error handler found in Historical Parse
    
    Err.Raise 5 'This will enable error handling in the CFTC Weekly Data sub and continue with trying to retrieve data via a QueryTable
    
    Exit Function

End Function
Public Sub Filter_Market_Arrays(ByRef Contract_CLCTN As Collection, Report_Type As String, Optional ICE_Market As Boolean = False)
    
Dim TempB As Variant, FilterC() As Variant, T As Long, Array_Count As Long, Unknown_Filter As Boolean
      
With Contract_CLCTN

    Array_Count = .Count
    
    If Array_Count > 1 Then
        FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, Report_Type:=Report_Type, Return_Filter_Columns:=True, Return_Filtered_Array:=False, ICE:=ICE_Market) '1 Based Positionl array filter
        Unknown_Filter = False
    Else
        Unknown_Filter = True
    End If
    
    For T = .Count To 1 Step -1
        
        TempB = .Item(T)
        
        .Remove T
        
        TempB = Filter_Market_Columns(convert_skip_col_to_general:=True, Return_Filter_Columns:=False, _
                                        Return_Filtered_Array:=True, _
                                        InputA:=TempB, _
                                        ICE:=ICE_Market, _
                                        Column_Status:=FilterC, _
                                        Create_Filter:=Unknown_Filter, Report_Type:=Report_Type)
                                        
        If T = .Count + 1 Then 'If last item in Collection then Simply re-add
            .Add TempB
        Else
            .Add TempB, Before:=T
        End If
        
    Next T

End With

End Sub
Public Function Filter_Market_Columns(Return_Filter_Columns As Boolean, _
                                       Return_Filtered_Array As Boolean, _
                                       convert_skip_col_to_general As Boolean, _
                                       Report_Type As String, _
                                       Optional Create_Filter As Boolean = True, _
                                       Optional ByVal InputA As Variant, _
                                       Optional ICE As Boolean = False, _
                                       Optional ByVal Column_Status As Variant) As Variant
'======================================================================================================
'Generates an array referencing RAW data columns to determine if they should be kept or not
'If and array is given an return_filtered_array=True then the array will be filtered column wise based on the previous array
'======================================================================================================

Dim ZZ As Long, Output() As Variant, V As Byte, Y As Byte, _
contractIdField As Byte, num As Byte, columnInOutput As Byte, finalColumnIndex As Byte

Dim CFTC_Wanted_Columns() As Variant, dateField As Byte, skip_value As Byte, twoDimensionalArray As Boolean

Const nameField As Byte = 1

CFTC_Wanted_Columns = Variable_Sheet.ListObjects(Report_Type & "_User_Selected_Columns").DataBodyRange.Columns(2).Value2

If ICE Then
    dateField = 2
    contractIdField = 7
Else
    dateField = 3
    contractIdField = 4
End If

Select Case Report_Type
    Case "L":
        num = 127
    Case "D":
        num = 187
    Case "T":
        num = 83
End Select

If convert_skip_col_to_general Then
    skip_value = xlGeneralFormat
Else
    skip_value = xlSkipColumn
End If

If Create_Filter = True And IsMissing(Column_Status) Then 'IF column Status is empty or if it is empty

    ReDim Column_Status(1 To UBound(CFTC_Wanted_Columns, 1))
    
    For V = 1 To UBound(CFTC_Wanted_Columns, 1)
        
        If CFTC_Wanted_Columns(V, 1) = True Or V = dateField Or V = contractIdField Then
            '^ allows entry into block regardless of if ICE or CFTC is needed for dates or contract code
        
            Select Case V
            
                Case dateField 'column 2 or 3 depending on if ICE or not
                
                    Column_Status(V) = xlMDYFormat
                    
                Case nameField, contractIdField
                
                    Column_Status(V) = xlTextFormat
                    
                Case 2, 3, 4, 7 'These numbers may overlap with dates column or contract field
                                'The previous cases will prevent it from executing unnecessarily depending on if ICE or not
                    Column_Status(V) = skip_value
                    
                Case Else
                
                    Column_Status(V) = xlGeneralFormat
                    
            End Select
            
        Else
        
            If V = num And convert_skip_col_to_general Then
                Column_Status(V) = xlTextFormat
            Else
                Column_Status(V) = skip_value 'skip these columns
            End If
            
        End If
        
    Next V
    
End If

If Return_Filter_Columns = True Then

    Filter_Market_Columns = Column_Status
    
ElseIf Return_Filtered_Array = True Then
    
     'Don't worry about text files..they are filtered in the same sub that they are opened in
     'FYI dateField would be 2 if doing TXT files...2 is used for ICE contracts because of exchange inconsistency
    On Error Resume Next

    Y = 0

    Do 'Determine the total number of dimensions
    
        Y = Y + 1
        V = LBound(InputA, Y)
        
    Loop Until Err.Number <> 0
    
    On Error GoTo 0
    
    If Y - 1 = 2 Then twoDimensionalArray = True
    
    If twoDimensionalArray Then
        ReDim Output(1 To UBound(InputA, 1), 1 To UBound(filter(Column_Status, xlSkipColumn, False)) + 1)
        finalColumnIndex = UBound(Output, 2)
    Else
        ReDim Output(1 To UBound(filter(Column_Status, xlSkipColumn, False)) + 1)
        finalColumnIndex = UBound(Output, 1)
    End If
    
    Y = 0
    
    For V = LBound(Column_Status) To UBound(Column_Status)
        
        If Column_Status(V) <> xlSkipColumn Then
            
            Select Case V
            
                Case nameField
                    columnInOutput = 2
                Case dateField
                    columnInOutput = 1
                Case contractIdField
                    columnInOutput = finalColumnIndex
                Case Else
                    'Find the next value that excludes the above cases
                    Do
                        Y = Y + 1
                    Loop Until 2 < Y And Y < finalColumnIndex
                    
                    columnInOutput = Y
                    
            End Select
            
            If twoDimensionalArray Then
            
                For ZZ = LBound(Output, 1) To UBound(Output, 1)
                    Output(ZZ, columnInOutput) = InputA(ZZ, V)
                Next ZZ
            
            Else
                Output(columnInOutput) = InputA(V)
            End If
            
        End If
        
    Next V
    
    Filter_Market_Columns = Output
    
End If
    
End Function
Public Function Query_Text_Files(ByVal TXT_File_Paths As Collection, Report_Type As String, combined_wb As Boolean) As Variant

'======================================================================================================
'Queries text files in a collection and adds their contents(array) to a collection
'======================================================================================================

Dim qt As QueryTable, file As Variant, Found_QT As Boolean, Field_Info() As Variant, Output_Arrays As New Collection, _
Field_Info_ICE() As Variant
 
Dim headerCount As Byte
 
For Each qt In QueryT.QueryTables 'Search for the following query if it exists
    If InStr(1, qt.name, "TXT Import") > 0 Then
        Found_QT = True
        Exit For
    End If
Next qt

Field_Info = Filter_Market_Columns(convert_skip_col_to_general:=True, Report_Type:=Report_Type, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True) '^^ CFTC Column filter

If Report_Type = "D" Then 'ICE Data column filter
    Field_Info_ICE = Filter_Market_Columns(convert_skip_col_to_general:=True, Report_Type:=Report_Type, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
    
End If

For Each file In TXT_File_Paths
    
    Application.StatusBar = "Querying: " & file
    DoEvents
    
    If Not Found_QT Then
        
        Set qt = QueryT.QueryTables.Add(Connection:="TEXT;" & file, Destination:=QueryT.Cells(1, 1))
        
        With qt
            .name = "TXT Import"
            .BackgroundQuery = False
            .SaveData = False
            .TextFileCommaDelimiter = True
            .TextFileConsecutiveDelimiter = False
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
        End With
        
        Found_QT = True 'So that this statement isn't executed again
        
    End If
    
    With qt
        
        .Connection = "TEXT;" & file

        If file Like "*.csv" Then 'ICE Workbooks
            .TextFileColumnDataTypes = Field_Info_ICE
        Else
            .TextFileColumnDataTypes = Field_Info
        End If
        
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = False
        .Destination = QueryT.Cells(1, 1)
        
        .Refresh False
        
        headerCount = headerCount + 1
        
        If headerCount = 1 Then
            Output_Arrays.Add .ResultRange.Value2
        Else
            With .ResultRange
                Output_Arrays.Add .Offset(1).Resize(.Rows.Count - 1).Value2
            End With
        End If
        
        With .ResultRange
            .ClearContents 'Clear the Range
        End With
    
    End With

Next file

If Output_Arrays.Count > 1 Then
    Query_Text_Files = Multi_Week_Addition(Output_Arrays, Append_Type.Multiple_2d)
Else
    Query_Text_Files = Output_Arrays(1)
End If

qt.Delete

End Function
Public Sub Retrieve_Tuesdays_CLose(ByRef inputData As Variant, _
price_column As Byte, contractDataOBJ As Contract, overwrite_all_prices As Boolean, dates_in_column_1 As Boolean, Optional ByRef Data_Found As Boolean = False)

'======================================================================================================
'Retrieves price data and stores it in a given array
'======================================================================================================

Dim Use_QueryTable As Boolean, Y As Integer, Start_Date As Date, End_Date As Date, URL As String, _
Year_1970 As Date, x As Long, Yahoo_Finance_Parse As Boolean, Stooq_Parse As Boolean

Dim Price_Data() As String, Initial_Split_CHR As String, D_OHLC_AV() As String

Dim Close_Price As Byte, Secondary_Split_STR As String, Response_STR As String, QT_Connection_Type As String

Dim End_Date_STR As String, Start_Date_STR As String, Query_Name As String, Symbol As String, Reverse_sort_Order As Boolean

Dim qt As QueryTable, QueryTable_Found As Boolean, Using_QueryTable As Boolean, Query_Data() As Variant, date_column As Byte

Const All_Data_Date_Column As Byte = 3

With contractDataOBJ
    Symbol = .PriceSymbol
    If Symbol = vbNullString Then Exit Sub
    Yahoo_Finance_Parse = .UseYahooPrices
    Stooq_Parse = Not Yahoo_Finance_Parse
End With

If Not dates_in_column_1 Then
    date_column = All_Data_Date_Column
Else
    date_column = 1
End If

If inputData(1, date_column) > inputData(UBound(inputData, 1), date_column) Then
'For this sub to work correctly It must be ordered from oldest to newest
    Reverse_sort_Order = True
    inputData = Reverse_2D_Array(inputData)
End If

On Error GoTo Exit_Price_Parse

Start_Date = inputData(1, date_column)
End_Date = inputData(UBound(inputData, 1), date_column)

If Yahoo_Finance_Parse Then 'CSV File
    
    Query_Name = "Yahoo Finance Query"
    
    Year_1970 = DateSerial(1970, 1, 1) 'Yahoo bases there URLs on the date converted to UNIX time
    
    End_Date = DateAdd("d", 1, End_Date) '1 more day than is in range to encapsulate that day
    
    Start_Date_STR = DateDiff("s", Year_1970, Start_Date) 'Convert to UNIX time
    
    End_Date_STR = DateDiff("s", Year_1970, End_Date) 'An extra day is added to encompass the End Day
    
    URL = "https://query1.finance.yahoo.com/v7/finance/download/" & Symbol & _
            "?period1=" & Start_Date_STR & _
            "&period2=" & End_Date_STR & _
            "&interval=1d&events=history&includeAdjustedClose=true"
      
    QT_Connection_Type = "TEXT;"
    
Else 'CSV file from STOOQ
    
    Query_Name = "Stooq Query"
    
    End_Date_STR = Format(End_Date, "yyyymmdd")
    Start_Date_STR = Format(Start_Date, "yyyymmdd")
    
    URL = "https://stooq.com/q/d/l/?s=" & Symbol & "&d1=" & Start_Date_STR & "&d2=" & End_Date_STR & "&i=d"
    
    QT_Connection_Type = "URL;"

End If

End_Date_STR = vbNullString
Start_Date_STR = vbNullString
    
#If Mac Then

    On Error GoTo Exit_Price_Parse
    'On Error GoTo 0
    Using_QueryTable = True

    For Each qt In QueryT.QueryTables           'Determine if QueryTable Exists
        
        If InStr(1, qt.name, Query_Name) > 0 Then 'Instr method used in case Excel appends a number to the name
            QueryTable_Found = True
            Exit For
        End If
        
    Next qt
    
    If Not QueryTable_Found Then Set qt = QueryT.QueryTables.Add(QT_Connection_Type & URL, QueryT.Cells(1, 1))
    
    With qt
    
        If Not QueryTable_Found Then
        
            .BackgroundQuery = False
            .name = Query_Name
            
            On Error GoTo Workbook_Connection_Name_Already_Exists 'deletes the connection has the name and then rename
            
                .WorkbookConnection.name = Replace$(Query_Name, "Query", "Prices")
                
            On Error GoTo Exit_Price_Parse
            
        Else
            .Connection = QT_Connection_Type & URL
        End If
        
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SaveData = False
        
        On Error GoTo Remove_QT_And_Connection 'Delete both the Querytable and the connection and exit the sub

         .Refresh False
        
        On Error GoTo Exit_Price_Parse
        
        With .ResultRange
            
            If Yahoo_Finance_Parse Or Stooq_Parse Then 'an array of csv values not separated
            
                Query_Data = .value
                
            End If
            
            .ClearContents
            
        End With
        
    End With
    
    Set qt = Nothing
    Query_Name = vbNullString
    QT_Connection_Type = vbNullString
    
#Else

    On Error GoTo Exit_Price_Parse
    
    'Dim HTTP2 As New MSXML2.XMLHTTP60
    
    Dim HTTP2 As Object
    
    Set HTTP2 = CreateObject("Msxml2.ServerXMLHTTP")

    With HTTP2
        .Open "GET", URL, False
        .send
        Response_STR = .responseText
    End With
    
    Set HTTP2 = Nothing
    
#End If

URL = vbNullString

On Error GoTo Exit_Price_Parse
  
If Yahoo_Finance_Parse Or Stooq_Parse Then 'Parsing CSV Files
    
    If Not Using_QueryTable Then
        
        If InStr(1, Response_STR, 404) = 1 Or Len(Response_STR) = 0 Then Exit Sub 'Something likely wrong with the URl
        
        If Yahoo_Finance_Parse Then
            
            Initial_Split_CHR = Mid$(Response_STR, InStr(1, Response_STR, "Volume") + Len("volume"), 1) 'Finding Splitting_Charachter
        
        ElseIf Stooq_Parse Then
        
            Initial_Split_CHR = vbNewLine
            
        End If
        
        Price_Data = Split(Response_STR, Initial_Split_CHR)
           
    Else
    
        ReDim Price_Data(0 To UBound(Query_Data, 1) - 1) 'redim to fit all rows of the query array
         
        For x = 0 To UBound(Query_Data, 1) - 1 'Add everything  to array
            Price_Data(x) = Query_Data(x + 1, 1)
        Next x
        
        Erase Query_Data
        
    End If
    
    If overwrite_all_prices Then
        'Data Table has been selected to have all price data overwritten
        For Y = LBound(inputData, 1) To UBound(inputData, 1)
            inputData(Y, price_column) = Empty
        Next Y

    End If
    
    Secondary_Split_STR = Chr(44)
    x = LBound(Price_Data) + 1 'Skip headers
    
    Close_Price = 4 'Base 0 location of close prices within the queried array
    
End If

If Len(Response_STR) > 0 Then Response_STR = vbNullString
If Len(Initial_Split_CHR) > 0 Then Initial_Split_CHR = vbNullString

Y = 1

Start_Date = CDate(Left$(Price_Data(x), InStr(1, Price_Data(x), Secondary_Split_STR) - 1))

Do Until inputData(Y, date_column) >= Start_Date
'Align the data based on the date

    If Y + 1 <= UBound(inputData, 1) Then
        Y = Y + 1
    Else
        If Reverse_sort_Order Then inputData = Reverse_2D_Array(inputData)
        Exit Sub
    End If
    
Loop
 
For Y = Y To UBound(inputData, 1)

    On Error GoTo Error_While_Splitting
    
    Do Until Start_Date >= inputData(Y, date_column)
    'Loop until price dates meet or exceed wanted date
    '>= used in case there isnt  a price for the requested date
Increment_X:

        x = x + 1
        
        If x > UBound(Price_Data) Then
            Exit For
        Else
            Start_Date = CDate(Left$(Price_Data(x), InStr(1, Price_Data(x), Secondary_Split_STR) - 1))
        End If
        
    Loop

    On Error Resume Next
    
    If Start_Date = inputData(Y, date_column) Then
    
        D_OHLC_AV = Split(Price_Data(x), Secondary_Split_STR)
                
        If Not IsNumeric(D_OHLC_AV(Close_Price)) Then 'find first value that came before that isn't empty
        
            inputData(Y, price_column) = Empty
            
        ElseIf CDbl(D_OHLC_AV(Close_Price)) = 0 Then
        
            inputData(Y, price_column) = Empty
            
        Else
        
            inputData(Y, price_column) = CDbl(D_OHLC_AV(Close_Price))
            If Not Data_Found Then Data_Found = True
            
        End If
        
        Erase D_OHLC_AV
            
    End If
    
Ending_INcrement_X:
    
Next Y

Exit_Price_Parse:

    Erase Price_Data
    
    If Reverse_sort_Order Then inputData = Reverse_2D_Array(inputData)
    
Exit Sub

Remove_QT_And_Connection:
    
    qt.Delete
    
    Exit Sub
    
Workbook_Connection_Name_Already_Exists:

    ThisWorkbook.Connections(Replace(Query_Name, "Query", "Prices")).Delete
    
    qt.WorkbookConnection.name = Replace(Query_Name, "Query", "Prices")
    Resume Next

Error_While_Splitting:

    If Err.Number = 13 Then 'type mismatch error from using cdate on a non-date string
        Resume Increment_X
    Else
        Exit Sub
    End If
    
End Sub

#If Not DatabaseFile Then

    Public Sub Paste_To_Range(Optional Table_DataB_RNG As Range, Optional Data_Input As Variant, _
        Optional Sheet_Data As Variant, Optional Historical_Paste As Boolean = False, _
        Optional Target_Sheet As Worksheet, _
        Optional Overwrite_Data As Boolean = False)
    
    Dim Model_Table As ListObject, Invalid_STR() As String, I As Long, _
    Invalid_Found() As Variant, A_Dim As Long, ITR_Y As Long, ITR_X As Long
    
    If Not Historical_Paste Then 'If Weekly/Block data addition
        
        If Not Overwrite_Data Then ' If not completely overwriting the worksheet
        
            For I = UBound(Data_Input, 1) To 1 Step -1 'Search in reverse order for dates that are too old to add to sheet
            
                If Data_Input(I, 1) <= Sheet_Data(UBound(Sheet_Data, 1), 1) Then 'If the date for the compared row is less than or equal to the most recent date recorded on the worksheet
                    
                    If I = UBound(Data_Input, 1) Then Exit Sub 'If on the first loop then all data in the input array is already on the worksheet
                    
                    ReDim Invalid_Found(1 To UBound(Data_Input, 1) - I, 1 To UBound(Data_Input, 2)) 'Array sized to hold only the new data
                        
                    For ITR_Y = I + 1 To UBound(Data_Input, 1) 'Fill array
                    
                        A_Dim = A_Dim + 1
                        
                        For ITR_X = 1 To UBound(Data_Input, 2)
                        
                            Invalid_Found(A_Dim, ITR_X) = Data_Input(ITR_Y, ITR_X)
                        
                        Next ITR_X
                        
                    Next ITR_Y
                    
                    Data_Input = Invalid_Found
                    
                    Exit For
                    
                End If
                
            Next I
        
        Else
        
            Table_DataB_RNG.ClearContents
            
            'Table_DataB_RNG.ListObject.AutoFilter.ShowAllData
            
        End If
        
        On Error GoTo No_Table
        
        With Table_DataB_RNG
            
            .Worksheet.DisplayPageBreaks = False
    
            .Cells(IIf(Overwrite_Data = False, .Rows.Count + 1, 1), 1).Resize(UBound(Data_Input, 1), UBound(Data_Input, 2)).Value2 = Data_Input 'bottom row +1,1st column
            'Overwritten range depends on Overwrite Data Boolean, If true then overwrite all data on the worksheet
    
            With .ListObject
            
                If Not Overwrite_Data Then 'If just appending data
                
                    If .DataBodyRange.Rows.Count <> UBound(Data_Input, 1) + UBound(Sheet_Data, 1) Then
        
                        .Resize .Range.Resize(UBound(Data_Input, 1) + UBound(Sheet_Data, 1) + 1, .DataBodyRange.Columns.Count)
                        'resize to fit all data +1 to accomodate for headers
                    End If
                
                ElseIf .DataBodyRange.Rows.Count <> UBound(Data_Input, 1) Then
                    
                    .Resize .Range.Resize(UBound(Data_Input, 1) + 1, .DataBodyRange.Columns.Count)
                
                End If
                
            End With
            
        End With 'pastes the bottom row of the array if bottom date is greater than previous
        
    ElseIf Historical_Paste = True Then 'pastes to active sheet and retrieves headers from sheet 15
    
        On Error GoTo PROC_ERR_Paste
    
        Set Model_Table = ContractDetails(1).TableSource
            
        With Model_Table
        
            .DataBodyRange.Copy 'copy and paste formatting
            
            Target_Sheet.Range(.HeaderRowRange.Address).Value2 = .HeaderRowRange.Value2       'table headers
        
        End With
        
        With Target_Sheet
        
            .Range("A2").Resize(UBound(Sheet_Data, 1), UBound(Sheet_Data, 2)).Value2 = Sheet_Data  'Apply data to worksheet
            
            Set Model_Table = .ListObjects.Add(xlSrcRange, , , xlYes) 'create table from range on Target_Sheet
            
            .Range("A3").ListObject.DataBodyRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                            SkipBlanks:=False, Transpose:=False 'Paste Conditional Formats
                   
            .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:=vbNullString, SubAddress:= _
                   "'" & HUB.name & "'!A1", TextToDisplay:=.Cells(1, 1).value 'create hyperlink in top left cell
            
            On Error GoTo Re_Name '{Finding Valid Worksheet Name}
            
            .name = Split(Sheet_Data(UBound(Sheet_Data, 1), 2), " -")(0)
        
        End With
        
        Application.StatusBar = "Data has been added to sheet. Calculating Formulas"
            
    End If
    
    On Error GoTo 0
    
    Exit Sub
        
Re_Name:
    
       MsgBox " If you were attempting to add a new contract then the Worksheet name could not be changed automatically."
        
        Resume Next
        
PROC_ERR_Paste:
    
        MsgBox "Error: (" & Err.Number & ") " & Err.description, vbCritical
    
        Resume Next
        
No_Table:
        MsgBox "If you are seeing this then either a table could not be found in cell A1 or your version " & _
        "of Excel does not support the listbody object. Further data will not be updated. Contact me via email."
        Call Re_Enable: End
        
    End Sub
    
#End If

