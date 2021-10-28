Attribute VB_Name = "Data_Retrieval_Support"
Option Explicit


Public Sub Retrieve_Historical_Workbooks(ByRef Path_CLCTN As Collection, _
                                               ICE_Contracts As Boolean, _
                                               CFTC_Contracts As Boolean, _
                                               Mac_User As Boolean, _
                                            Optional ByVal CFTC_Start_Date As Date, _
                                            Optional ByVal CFTC_End_Date As Date, _
                                            Optional ByVal ICE_Start_Date As Date, _
                                           Optional ByVal ICE_End_Date As Date, _
                                            Optional ByVal Historical_Archive_Download As Boolean = False)

'
'Downloads files with a get request if on Windows
'If mac then it returns a collection of URLs for user to download
'
Dim File_Within_Zip As String, Path_Separator As String, AnnualOF_FilePath As String, Queried_Date As Long, _
Destination_Folder As String, File_Path As String, Zip_FileN As String, _
Download_Year As Long, Final_Year As Long, FileN As String, MT_Year As String, URL As String, Partial_Url As String

Dim Multi_Year_URL As String, g As Long, Contract_Data() As Variant, Multi_Year_ZipN As String, _
File_Name_End As String, Combined_WB As Boolean

Const TXT As String = ".txt", ZIP As String = ".zip", CSV As String = ".csv"

Const AnimeT As String = "B.A.T"

On Error GoTo Failed_To_Download

Path_Separator = Application.PathSeparator

Combined_WB = Combined_Workbook(Variable_Sheet)

If Not Mac_User Then

    Destination_Folder = Environ("TEMP") & Path_Separator & "COT_Historical_MoshiM" & Path_Separator & Data_Retrieval.TypeF & Path_Separator & IIf(Combined_WB = True, "Combined", "Futures Only")
    
    If Dir(Destination_Folder, vbDirectory) = vbNullString Then
        Shell ("cmd /c mkdir """ & Destination_Folder & """")
        
        Do Until Dir(Destination_Folder, vbDirectory) <> vbNullString
        Loop
    End If
    
Else

    Destination_Folder = vbNullString 'Keep variable as an empty string.User will decide path
    
End If

With Path_CLCTN

    If CFTC_Contracts Then
    
        If Not Combined_WB Then 'IF Futures Only Workbook
        
            File_Name_End = "_Futures_Only"
            
            Select Case Data_Retrieval.TypeF
            
                Case "L"
                
                    File_Within_Zip = "annual" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/deacot"

                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deacot1986_2016" & ZIP
                    
                    Contract_Data = Array("FUT86_16")
                    
                Case "D"
                
                    File_Within_Zip = "f_year" & TXT
                    Partial_Url = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_"

                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_hist_2006_2016" & ZIP
                    
                    Contract_Data = Array("F_DisAgg06_16")
                    
                Case "T"
                
                    File_Within_Zip = "FinFutYY" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/fut_fin_txt_"
                    
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_fut_txt_2006_2016" & ZIP
                    
                    Contract_Data = Array("F_TFF_2006_2016")
                    
            End Select
        
        Else 'Combined Contracts
        
            File_Name_End = "_Combined"
            
            Select Case Data_Retrieval.TypeF
            
                Case "L"
                
                    File_Within_Zip = "annualof.txt"
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/deahistfo" 'TXT URL
                    
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deahistfo_1995_2016" & ZIP
                    
                    Contract_Data = Array("Com95_16")
                    
                Case "D"
                
                    File_Within_Zip = "c_year" & TXT
                    
                    Partial_Url = "https://www.cftc.gov/files/dea/history/com_disagg_txt_"
                    'https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016.zip
                    Multi_Year_URL = "https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016" & ZIP
                    
                    Contract_Data = Array("C_DisAgg06_16")
                    
                Case "T"
                
                    File_Within_Zip = "FinComYY" & TXT
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
        
        Multi_Year_ZipN = Destination_Folder & Path_Separator & Data_Retrieval.TypeF & "_COT_MultiYear_Archive" & File_Name_End & ZIP
        
        If Not Mac_User Then AnnualOF_FilePath = Destination_Folder & Path_Separator & File_Within_Zip
        
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
            
            For g = LBound(Contract_Data) To UBound(Contract_Data) 'loop at least once ,iterate through Historical strings if needed
            
                If Historical_Archive_Download Then
                 
                    FileN = Destination_Folder & Path_Separator & Data_Retrieval.TypeF & "_" & Contract_Data(g) & File_Name_End & TXT
                     
                ElseIf Final_Year = Download_Year Then
                 
                    FileN = Destination_Folder & Path_Separator & Data_Retrieval.TypeF & "_Weekly_" & Queried_Date & "_" & Download_Year & File_Name_End & TXT
                 
                Else
                 
                    FileN = Destination_Folder & Path_Separator & Data_Retrieval.TypeF & "_" & Download_Year & File_Name_End & TXT
                
                End If
                
                If Not Mac_User Then
                
                    If Dir(FileN) = vbNullString Then   'If wanted workbook doesn't exist
                        
                        If g = LBound(Contract_Data) Then 'Only need to check if Zip file exists once
                        
                            If Historical_Archive_Download Then
                            
                                Zip_FileN = Multi_Year_ZipN
                                
                            Else
                            
                                Zip_FileN = Replace(FileN, TXT, ZIP)
                                
                            End If
                            
                            If Dir(Zip_FileN) = vbNullString Then Call Get_File(URL, Zip_FileN)
                            'Download Zip folder if it doesn't exist
                        End If
                        
                        If Not Historical_Archive_Download Then
                        
                            If Dir(AnnualOF_FilePath) <> vbNullString Then Kill AnnualOF_FilePath 'If file within Zip folder exists within file directory then kill it
                        
                            Call entUnZip1File(Zip_FileN, Destination_Folder, File_Within_Zip) 'Unzip specified file
                            
                            Name Destination_Folder & Path_Separator & File_Within_Zip As FileN
                            
                        ElseIf Historical_Archive_Download Then
                        
                            MT_Year = Destination_Folder & Path_Separator & Contract_Data(g) & TXT
                            
                            If Dir(MT_Year) <> vbNullString Then Kill MT_Year
                            
                            Call entUnZip1File(Zip_FileN, Destination_Folder, Contract_Data(g) & TXT)
                            
                            Name MT_Year As FileN
                        
                        End If
                            
                    End If
                    
                    .Add FileN, FileN 'Add file name to collection
                    
                ElseIf Mac_User Then
                
                    FileN = Replace(FileN, Path_Separator, vbNullString)
                    
                    If (g = LBound(Contract_Data) And Historical_Archive_Download) Or Not Historical_Archive_Download Then
                    
                        .Add Array(URL, FileN), FileN
                        
                    Else
                    
                        .Add Array(AnimeT, FileN), FileN
                        
                    End If
                    
                End If
                
                If Not Historical_Archive_Download Then Exit For
                
            Next g

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
                    FileN = Destination_Folder & Path_Separator & "ICE_Weekly_" & Queried_Date & "_" & Download_Year & ".csv"
                Case Else
                    FileN = Destination_Folder & Path_Separator & "ICE_" & Download_Year & ".csv"
            End Select

            If Not Mac_User Then
            
                If Dir(FileN) = vbNullString Then Call Get_File(URL, FileN)
                
                .Add FileN, FileN
            
            ElseIf Mac_User Then
            
                FileN = Replace(FileN, Path_Separator, vbNullString)
                
                .Add Array(URL, FileN), FileN
            
            End If

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
Public Function Stochastic_Calculations(Column_Number_Reference As Long, Time_Period As Long, _
                                    arr As Variant, Optional Missing_Weeks As Long = 1, Optional Cap_Extremes As Boolean = False) As Variant

Dim Array_Column() As Double, X As Long, Array_Period() As Double, Current_Row As Long, _
Array_Final() As Variant, UB As Long, Initial As Long, Minimum As Double, Maximum As Double

UB = UBound(arr, 1) 'number of rows in the supplied array[Upper bound in 1st dimension]

Initial = UB - (Missing_Weeks - 1) 'will be equal to UB if only 1 missing week

ReDim Array_Period(1 To Time_Period)  'Temporary Array that will hold a certain range of data

ReDim Array_Final(1 To Missing_Weeks) 'Array that will hold calculated values

ReDim Array_Column(IIf(Initial > Time_Period, Initial - Time_Period, 1) To UB) 'Array composed of all data in a column

If UBound(arr, 2) = 1 Then Column_Number_Reference = 1 'for when a single column is supplied

For X = IIf(Initial > Time_Period, Initial - Time_Period, 1) To UB
    'if starting row of data output is greater than the time period then offset the start of the queried array by the time period
    'otherwise start at 1...there should be checks to ensure there is enough data most of the time
    
    Array_Column(X) = arr(X, Column_Number_Reference)
    
Next X

For Current_Row = Initial To UB

    If (Current_Row > Time_Period And Not Cap_Extremes) Or (Current_Row >= Time_Period And Cap_Extremes) Then   'Only calculate if there is enough data
    
        For X = 1 To Time_Period 'Fill array with the previous Time_Period number of values relative to the current row
            
            If Not Cap_Extremes Then
                Array_Period(X) = Array_Column(Current_Row - X)
            Else
                Array_Period(X) = Array_Column(Current_Row - (X - 1))
            End If
            
            If X = 1 Then
                Minimum = Array_Period(X)
                Maximum = Array_Period(X)
            Else
            
                If Array_Period(X) < Minimum Then
                
                    Minimum = Array_Period(X)
                    
                ElseIf Array_Period(X) > Maximum Then
                
                    Maximum = Array_Period(X)
                    
                End If
                
            End If
            
        Next X
        'Stochastic calculation
        If Maximum <> Minimum Then
            Array_Final(Missing_Weeks - (UB - Current_Row)) = CLng(((Array_Column(Current_Row) - Minimum) / (Maximum - Minimum)) * 100)
        End If
        'ex for determining current location within array:    2 - ( 480 - 479 ) = 1
    End If

Next Current_Row

Stochastic_Calculations = Array_Final
    
End Function

Public Sub Paste_To_Range(Optional Table_DataB_RNG As Range, Optional Data_Input As Variant, _
    Optional Sheet_Data As Variant, Optional Historical_Paste As Boolean = False, _
    Optional Target_Sheet As Worksheet, _
    Optional Overwrite_Data As Boolean = False)
 
Dim Model_Table As ListObject, Invalid_STR() As String, i As Long, _
Invalid_Found() As Variant, A_Dim As Long, ITR_Y As Long, ITR_X As Long

If Not Historical_Paste Then 'If Weekly/Block data addition
    
    If Not Overwrite_Data Then ' If not completely overwriting the worksheet
    
        For i = UBound(Data_Input, 1) To 1 Step -1 'Search in reverse order for dates that are too old to add to sheet
        
            If Data_Input(i, 1) <= Sheet_Data(UBound(Sheet_Data, 1), 1) Then 'If the date for the compared row is less than or equal to the most recent date recorded on the worksheet
                
                If i = UBound(Data_Input, 1) Then Exit Sub 'If on the first loop then all data in the input array is already on the worksheet
                
                ReDim Invalid_Found(1 To UBound(Data_Input, 1) - i, 1 To UBound(Data_Input, 2)) 'Array sized to hold only the new data
                    
                For ITR_Y = i + 1 To UBound(Data_Input, 1) 'Fill array
                
                    A_Dim = A_Dim + 1
                    
                    For ITR_X = 1 To UBound(Data_Input, 2)
                    
                        Invalid_Found(A_Dim, ITR_X) = Data_Input(ITR_Y, ITR_X)
                    
                    Next ITR_X
                    
                Next ITR_Y
                
                Data_Input = Invalid_Found
                
                Exit For
                
            End If
            
        Next i
    
    Else
    
        Table_DataB_RNG.ClearContents
        
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

    Set Model_Table = CFTC_Table(ThisWorkbook, Model_S)
        
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
               "'" & HUB.Name & "'!A1", TextToDisplay:=.Cells(1, 1).Value 'create hyperlink in top left cell
        
        On Error GoTo Re_Name '{Finding Valid Worksheet Name}
        
        .Name = Split(Sheet_Data(UBound(Sheet_Data, 1), 2), " -")(0)
    
    End With
    
    Application.StatusBar = "Data has been added to sheet. Calculating Formulas"
        
End If

On Error GoTo 0

Exit Sub
    
Re_Name:

   MsgBox " If you were attempting to add a new contract then the Worksheet name could not be changed automatically."
    
    Resume Next
    
PROC_ERR_Paste:

    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical

    Resume Next
    
No_Table:
    MsgBox "If you are seeing this then either a table could not be found in cell A1 or your version " & _
    "of Excel does not support the listbody object. Further data will not be updated. Contact me via email."
    Call Re_Enable: End
    
End Sub
Public Function Test_For_Data_Addition(Optional WKB As String) As Boolean

'If the CFTC has updated data since the last time the compiled workbook[not this file] was edited then return TRUE

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

Public Function Multi_Week_Addition(My_CLCTN As Collection, Sort_Type As Long) As Variant 'adds the contents of the NEW array TO the contents of the OLD
  
'
'Combines all arrays with less than 3 dimensions into a single 2D array
'

Dim T As Long, X As Long, S As Long, UB1 As Long, UB2 As Long, Worksheet_Data() As Variant, _
Item As Variant, Old() As Variant, Block() As Variant, Latest() As Variant, Not_Old As Long, Is_Old As Long
   
'Dim Addition_Timer As Double: Addition_Timer = Time

With My_CLCTN
    'check the boundaries of the elements to create the array
    Select Case Sort_Type

        Case Append_Type.Multiple_1d 'Array Elements are 1D | single rows |  "Historical_Parse"

            UB1 = .Count 'The number of items in the dictionary will be the number of rows in the final array

            For Each Item In My_CLCTN 'loop through each item in the row and find the max number of columns
                
                S = UBound(Item) + 1 - LBound(Item) 'Number of Columns if 1 based
            
                If S > UB2 Then UB2 = S
                
            Next Item
            
        Case Append_Type.Multiple_2d 'Indeterminate number of  2D[1-Based] arrays to be joined.
                              '"Historical_Excel_Creation"
            For Each Item In My_CLCTN
                
                UB1 = UBound(Item, 1) + UB1
                
                S = UBound(Item, 2)
                
                If S > UB2 Then UB2 = S
                
            Next Item

        Case Append_Type.Add_To_Old
        
            'Normal Weekly operation [appending 1 based 1D or 2D array to 1 based 2D]
            'Used in Subs Sort_L and Block_Query
            
            Not_Old = 1
            
            Do Until .Item(Not_Old)(0) <> Data_Identifier.Old_Data
                Not_Old = Not_Old + 1
            Loop
            
            Is_Old = IIf(Not_Old = 1, 2, 1)
            
            If Is_Old <> 1 Then 'Ensure that it's the first item in the Collection
            
                Old = .Item(Is_Old)
                .Remove Is_Old
                .Add Old, "Old", Before:=1
                
            End If
            
            Old = .Item(Is_Old)(1)              'Assign to element that holds the array
                
            S = UBound(Old, 2)                  'Number of Columns used for the Old data
            
            Select Case .Item(Not_Old)(0)       'Number designating array type
            
                Case Data_Identifier.Weekly_Data  'This key is used for when sotring weekly data
                
                    Latest = .Item(Not_Old)(1)
                    
                    T = UBound(Latest)              'Number of columns in the 1-based 1D array
                    
                    UB1 = UBound(Old, 1) + 1 ' +1 Since there will be only 1 row of additional weekly data
                
                Case Data_Identifier.Block_Data  'This key is used if several weeks have passed
                                                        'This will be a 2d array
                    Block = .Item(Not_Old)(1)
                    
                    T = UBound(Block, 2)
                    
                    UB1 = UBound(Old, 1) + UBound(Block, 1)
                
            End Select
            
            If S >= T Then 'Determing number of columns to size the array with
                UB2 = S    'S= # of Columns in the older data
            Else           'T= # of Columns in the new data
                UB2 = T
            End If

    End Select
    
    ReDim Worksheet_Data(1 To UB1, 1 To UB2)
    
    S = 1
    
    For Each Item In My_CLCTN
        
        Select Case Sort_Type

            Case Append_Type.Multiple_1d 'All items in Collection are 1D
                
                For T = LBound(Item) To UBound(Item) 'Worksheet_Data is 1 based
                    
                    Worksheet_Data(S, T + 1 - LBound(Item)) = Item(T)

                Next T

                S = S + 1
                
            Case Append_Type.Multiple_2d 'Adding Multiple 2D arrays together
    
                    X = 1
                    
                    For S = S To UBound(Item, 1) + (S - 1)

                        For T = LBound(Item, 2) To UBound(Item, 2)

                            Worksheet_Data(S, T) = Item(X, T)
                            
                        Next T
                        
                        X = X + 1
                    
                    Next S
            
            Case Append_Type.Add_To_Old 'Adding new Data to a 2D Array..Block is 2D...Latest is 1D
                                        
                Select Case Item(0)                 'Key of item

                    Case Data_Identifier.Old_Data 'Current Historical data on Worksheet
                        
                        For S = LBound(Worksheet_Data, 1) To UBound(Old, 1)

                            For T = LBound(Old, 2) To UBound(Old, 2)

                                Worksheet_Data(S, T) = Old(S, T)

                            Next T
                            
                        Next S
                        
                    Case Data_Identifier.Block_Data '<--2D Array used when adding to arrays together where order is important
                    
                        X = 1
                        
                        For S = UBound(Worksheet_Data, 1) - UBound(Block, 1) + 1 To UBound(Worksheet_Data, 1)

                            For T = LBound(Block, 2) To UBound(Block, 2)

                                Worksheet_Data(S, T) = Block(X, T)
                                
                            Next T
                            
                            X = X + 1
                        
                        Next S
                        
                    Case Data_Identifier.Weekly_Data  '1 based 1D "WEEKLY" array
                                      '"OLD" is run first so S is already at the correct incremented value
                        For T = LBound(Latest) To UBound(Latest)
                            
                            Worksheet_Data(S, T) = Latest(T) 'worksheet data is 1 based 2D while Item is 1 BASED 1D

                        Next T
                                      
                End Select

        End Select
        
    Next Item

End With

Multi_Week_Addition = Worksheet_Data
    
'Debug.Print Timer - Addition_Timer

End Function

Public Function HTTP_Weekly_Data(Last_Update As Long, Optional Auto_Retrieval As Boolean = False, _
                                Optional Valid_Tables_Available As Boolean = False, _
                                Optional DebugMD As Boolean = False) As Variant

'
'Retrieves the latest week's data
'
Dim PowerQuery_Available As Boolean, Power_Query_Failed As Boolean, _
Text_Method_Failed As Boolean, Query_Table_Method_Failed As Boolean, MAC_OS As Boolean

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

If Valid_Tables_Available = False Then  'Valid_Table_Info is a public array that holds valid table info

    Data_Retrieval.Valid_Table_Info = Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info")
    
End If

If Not MAC_OS And (PowerQuery_Available Or DebugMD) Then

    On Error GoTo Power_Query_Failed
    
    'If UUID Then GoTo TXT_Method ''''

    HTTP_Weekly_Data = CFTC_Data_PowerQuery_Method
    
    If DebugMD Then GoTo TXT_Method
    
ElseIf Not MAC_OS And (Not PowerQuery_Available Or DebugMD) Then  'TXT file Method

TXT_Method:

    On Error GoTo TXT_Failed
  
    HTTP_Weekly_Data = CFTC_Data_Text_Method(Last_Update)

    If DebugMD Then GoTo Query_Table_Method
    
Else

Query_Table_Method:

    On Error GoTo Query_Table_Failed

    HTTP_Weekly_Data = CFTC_Data_QueryTable_Method
    
End If

If Valid_Tables_Available = False Then Erase Data_Retrieval.Valid_Table_Info 'if Valid_Table_Info wasn't allocated before running this function
                                              'then this function is being run as an automatic check and Valid_Table_Info should be erased
Exit Function

Power_Query_Failed:

    Resume TXT_Method
    
TXT_Failed:

    Resume Query_Table_Method
    
Query_Table_Failed:
    
    If Auto_Retrieval Then 'if an error has occured and this is part of an automatic data retrieval and date comparison
        Exit Function      ' Then return nothing from the function
    Else
        GoTo Failed_Data_Retrieval 'show error message
    End If

Failed_Data_Retrieval: 'Both Power Query and HTTP methods have failed
    
    MsgBox "Data retrieval has failed." & vbNewLine & vbNewLine & _
           "If you are on Windows and Power Query is available and you aren't using Excel 2016 then please enable or download Power Query / Get and Transform and try again." & vbNewLine & vbNewLine & _
           "If you are on a MAC or the above step fails then please contact me at MoshiM_UC@outlook.com with your operating system and Excel version."
               
'clean memory,Re-enable Application properties and cease macro processing

End_All_Processes: 'Used when function is called manually from worksheet button
    
    Erase Data_Retrieval.Valid_Table_Info
    
    Call Re_Enable
    
    End
Default_No_Power_Query:
    PowerQuery_Available = False
    Resume Retrieval_Process
    
End Function
Public Function CFTC_Data_PowerQuery_Method() As Variant
    
    With Weekly.ListObjects("Weekly")
        .QueryTable.Refresh False                               'Refresh Weekly Data Table
        CFTC_Data_PowerQuery_Method = .DataBodyRange.Value2     'Store contents of table in an array
    End With
    
End Function
Public Function CFTC_Data_Text_Method(Last_Update As Long) As Variant

Dim File_Path As New Collection

Dim URL As String, Y As Long

    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(Data_Retrieval.TypeF, Array("L", "D", "T"), 0) - 1
    
    If Not Combined_Workbook(Variable_Sheet) Then 'Futures Only
        URL = URL & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        URL = URL & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    With File_Path
    
        .Add Environ("TEMP") & "\" & Date & "_" & Data_Retrieval.TypeF & "_Weekly.txt", "Weekly Text File" 'Add file path of file to be downloaded
    
        Call Get_File(URL, .Item("Weekly Text File")) 'Download the file to the above path
        
    End With
    
    CFTC_Data_Text_Method = Historical_Parse(File_Path, CFTC_TXT:=True, After_This_Date:=Last_Update)  'return array
    
End Function
Public Function CFTC_Data_QueryTable_Method() As Variant

Dim Data_Query As QueryTable, Temp_C As New Collection, Data() As Variant, URL As String, Valid_Codes() As Variant, _
Column_Filter() As Variant, Y As Long, Parsed() As Variant, ZZ As Long, EE As Variant, BB As Boolean, Variables, _
Found_Data_Query As Boolean, Combined_WB As Boolean, Error_While_Refreshing As Boolean

Dim Workbook_Type As String

With Application

    BB = .EnableEvents
    
    .EnableEvents = False
    .DisplayAlerts = False

    Valid_Codes = Variable_Sheet.ListObjects("Table_WSN").DataBodyRange.Columns(1).Value2

End With

Combined_WB = Combined_Workbook(Variable_Sheet)

Workbook_Type = IIf(Combined_WB, "Combined", "Futures_Only")

For Each Data_Query In QueryT.QueryTables
    If InStr(1, Data_Query.Name, "CFTC_Data_Weekly_" & Workbook_Type) > 0 Then
        Found_Data_Query = True
        Exit For
    End If
Next Data_Query

If Not Found_Data_Query Then 'If QueryTable isn't found then create it

Recreate_Query:
    
    Column_Filter = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True)

    URL = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(Data_Retrieval.TypeF, Array("L", "D", "T"), 0) - 1
    
    If Not Combined_WB Then 'Futures Only
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
        
        .TextFileColumnDataTypes = Column_Filter
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        
        .Name = "CFTC_Data_Weekly_" & Workbook_Type
        
        On Error GoTo Delete_Connection
        
Name_Connection:

        With .WorkbookConnection
            .RefreshWithRefreshAll = False
            .Name = "Weekly CFTC Data: " & Workbook_Type
        End With
        
    End With
    
    On Error GoTo 0
    
    Erase Column_Filter

End If

On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time

With Data_Query

    .Refresh False
    
    With .ResultRange
        Data = .Value2 'Store Data in an Array
        .ClearContents 'Clear the Range
    End With
    
End With
    
On Error GoTo 0

ReDim Parsed(1 To UBound(Data, 1), 1 To UBound(Data, 2))

For ZZ = 1 To UBound(Data, 1)

    For Y = 1 To UBound(Data, 2)
    
        Select Case Y
        
            Case 1
            
                Parsed(ZZ, Y) = Data(ZZ, 2) 'Dates placed in 1st column
                
            Case 2
            
                 Parsed(ZZ, Y) = Data(ZZ, 1) 'Market Name -2nd Column
                
            Case UBound(Parsed, 2)
            
                Parsed(ZZ, Y) = Data(ZZ, 3) 'Contract Code-Last Column
                
            Case Else
            
                Parsed(ZZ, Y) = Data(ZZ, Y + 1) 'shift everything else left 1
                
        End Select
        
    Next Y
    
Next ZZ

ReDim Data(1 To 1, 1 To UBound(Data, 2))

With Application

    For ZZ = 1 To UBound(Parsed, 1) 'Loop each row of parsed
          
        If Not IsError(.Match(Parsed(ZZ, UBound(Parsed, 2)), Valid_Codes, 0)) Then
            
            For Y = 1 To UBound(Data, 2)
                Data(1, Y) = Parsed(ZZ, Y)
            Next Y
            
            Temp_C.Add Data, Data(1, UBound(Data, 2))
        
        End If
        
    Next ZZ
    
'    For ZZ = 1 To UBound(Valid_Codes, 1)
'        If IsError(.Match(Valid_Codes(ZZ, 1), Temp_C.Keys, 0)) Then Debug.Print Valid_Codes(ZZ, 1)
'    Next ZZ
    
    ReDim Parsed(1 To Temp_C.Count, 1 To UBound(Parsed, 2))
    
    ZZ = 1
    
    For Each EE In Temp_C 'combine elements into a single array
    
        For Y = 1 To UBound(Parsed, 2)
            Parsed(ZZ, Y) = EE(1, Y)
        Next Y
        
        ZZ = ZZ + 1
        
    Next EE
    
    CFTC_Data_QueryTable_Method = Parsed
    
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
Public Function Historical_Parse(ByVal File_CLCTN As Collection, _
                                  Optional ByRef Contract_Code As String = vbNullString, _
                                  Optional After_This_Date As Long = 0, _
                                  Optional Kill_Previous_Workbook As Boolean = False, _
                                  Optional Yearly_C As Boolean, _
                                  Optional Specified_Contract As Boolean, _
                                  Optional Weekly_ICE_Data As Boolean, _
                                  Optional CFTC_TXT As Boolean, _
                                  Optional Parse_All_Data As Boolean)



'
'Parses data from files when given path collection
'Compiled Data will be filtered
'

Dim Date_Sorted As New Collection, Item As Variant, Escape_Filter_Market_Arrays As Boolean, _
Contract_WB As Workbook, Contract_WB_Path As String, ICE_Data As Boolean, Mac_UserB As Boolean, Combined_Version As Boolean

Dim FilterC() As Variant, Output() As Variant, ZZ As Long, Y As Long, ErrorC As Collection

On Error GoTo Historical_Parse_General_Error_Handle

#If Mac Then
    Mac_UserB = True
#End If

With ThisWorkbook

    If Not HasKey(.Event_Storage, "Historical Parse Errors") Then
    
        .Event_Storage.Add New Collection, "Historical Parse Errors"
        
    End If
    
    Set ErrorC = .Event_Storage("Historical Parse Errors")
    
End With

Combined_Version = Combined_Workbook(Variable_Sheet)

Application.ScreenUpdating = False

Select Case True

    Case Yearly_C, Specified_Contract, Parse_All_Data 'Parse all data is for when all data is being downloaded
           
        If Not Mac_UserB Then 'Excecute if on Windows PC otherwise created file will be deleted at end
            
            Contract_WB_Path = Left$(File_CLCTN(1), InStrRev(File_CLCTN(1), Application.PathSeparator)) 'Folder path of File_CLCTN(1)
        
            If Yearly_C Then 'If Yearly Contracts ie: only missing a chunk of data
            
                Contract_WB_Path = Contract_WB_Path & Data_Retrieval.TypeF & "_COT_Yearly_Contracts_" & IIf(Combined_Version, "Combined", "Futures_Only") & ".xlsb"
                
            ElseIf Specified_Contract Or Parse_All_Data Then  'If using the new contract macro
            
                Contract_WB_Path = Contract_WB_Path & Data_Retrieval.TypeF & "_COT_Historical_Archive_" & IIf(Combined_Version, "Combined", "Futures_Only") & ".xlsb"
                
            End If
        
            If Dir(Contract_WB_Path) = vbNullString Then 'If the file doesn't exist then compile the text files into a single document and create a workbook from it
                
                Call Historical_TXT_Compilation(File_CLCTN, Contract_WB, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=Mac_UserB)
                
            Else 'If the file exists
    
                If Test_For_Data_Addition(Contract_WB_Path) Or Kill_Previous_Workbook = True Then 'if new data has been added since last workbook was created
                
                    Kill Contract_WB_Path 'Kill and then recreate
                    
                    Call Historical_TXT_Compilation(File_CLCTN, Contract_WB, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=Mac_UserB)
                    
                Else
                    
                    Set Contract_WB = Workbooks.Open(Contract_WB_Path)  'Set a reference
                    
                    Contract_WB.Windows(1).Visible = False
    
                End If
    
            End If
           
        ElseIf Mac_UserB Then 'Workbook will be created but will not be saved...Workbook path supplied is an empty string
        
            Call Historical_TXT_Compilation(File_CLCTN, Contract_WB, Saved_Workbook_Path:=Contract_WB_Path, OnMAC:=Mac_UserB)
            
        End If
        
        If Data_Retrieval.TypeF = "D" Then ICE_Data = True 'Checking if Disaggregated Workbook
        
        Call Historical_Excel_Aggregation(Contract_WB, CLCTN_A:=Date_Sorted, Contract_ID:=Contract_Code, Date_Input:=After_This_Date, Specified_Contract:=Specified_Contract)
        
        Contract_WB.Close False 'Close without saving
        
        ICE_Data = False 'Workbook Structure has been homogenized
        
    Case Weekly_ICE_Data 'Result=2D Array stored in Collection, Array isn't filtered
        
        ICE_Data = True
        
        Set Contract_WB = Workbooks.Open(File_CLCTN.Item("ICE"))
        
        With Contract_WB
        
            .Windows(1).Visible = False
        
            Call Historical_Excel_Aggregation(Contract_WB, CLCTN_A:=Date_Sorted, Date_Input:=After_This_Date, Weekly_ICE_Contracts:=True)
            
            .Close False
            
            Kill File_CLCTN.Item("ICE")
            
        End With
        
    Case CFTC_TXT 'Result=2D Array stored in Collection2D Array(s) stored in Collection from .txt file(s)

        Call Weekly_Text_File(File_CLCTN, StorageC:=Date_Sorted, Date_Value:=After_This_Date)
        
End Select

If (Yearly_C Or Specified_Contract Or Parse_All_Data Or CFTC_TXT) And Date_Sorted.Count > 0 Then 'Data columns have already been filtered and just need to be rearranged
    
    Escape_Filter_Market_Arrays = True 'enabled so it isnt filtered again
    
    With Date_Sorted
    
        FilterC = Date_Sorted(1) 'there should only be one item in the collection
        
        .Remove .Count
        
        ReDim Output(1 To UBound(FilterC, 1), 1 To UBound(FilterC, 2))
        
        For ZZ = 1 To UBound(Output, 1)
        
            For Y = 1 To UBound(Output, 2)
                    
                Select Case Y
                    Case 1:                  Output(ZZ, Y) = FilterC(ZZ, 2) 'Dates placed in 1st column
                    Case 2:                  Output(ZZ, Y) = FilterC(ZZ, 1) 'Market Name -2nd Column
                    Case UBound(Output, 2):  Output(ZZ, Y) = FilterC(ZZ, 3) 'Contract Code-Last Column
                    Case Else:               Output(ZZ, Y) = FilterC(ZZ, Y + 1) 'shift everything else left 1
                End Select
                
            Next Y
            
        Next ZZ
        
        .Add Output, "Filtered Columns"
        Erase Output
        Erase FilterC
        
    End With
    
End If

With Date_Sorted
    
    If .Count = 0 Then 'If there are no items in the dictionary
        
        Historical_Parse = WorksheetFunction.Transpose(Array(After_This_Date, "B.A.T"))
    
        Exit Function
    
    End If
    
    If Not Escape_Filter_Market_Arrays Then Filter_Market_Arrays Date_Sorted, ICE_Data 'converts and prunes certain elements
    
    If .Count > 1 Then 'Join an indefinite number of 2D arrays
    
        Historical_Parse = Multi_Week_Addition(Date_Sorted, Multiple_2d)
        
    ElseIf .Count = 1 Then
    
        Historical_Parse = .Item(1)
        
    End If

End With

Application.StatusBar = vbNullString

Exit Function

Historical_Parse_General_Error_Handle:

    If CFTC_TXT Or Weekly_ICE_Data Then  'use parent error handler
    
        On Error GoTo 0
        Err.Raise 5
        
    ElseIf Yearly_C Or Specified_Contract Or Parse_All_Data Then
    
        Contract_WB_Path = "An error has occured while running the Historical Parse subroutine. Please email me at MoshiM_UC@outlook.com"
        
        For Y = 1 To ErrorC.Count
            
            Contract_WB_Path = Contract_WB_Path & vbNewLine & vbNewLine & ErrorC(Y) & vbNewLine & vbNewLine
            
        Next Y
        
        MsgBox Contract_WB_Path
        
        ThisWorkbook.Event_Storage.Remove "Historical Parse Errors"
        Set ErrorC = Nothing
        
        Re_Enable
        
        End
    
    End If
    
End Function
Public Sub Historical_TXT_Compilation(File_Collection As Collection, ByRef Contract_WB As Workbook, _
Saved_Workbook_Path As String, OnMAC As Boolean)
    
Dim File_TXT As Variant, FileNumber As Long, Data_STR As String, File_Path() As String

Dim InfoF() As Variant, FilterC As Variant, D As Long, ICE_Filter As Boolean, UB As Long

Dim Combined_WB As Boolean, File_Name As String ', DD As Double

'If Not OnMAC Then

On Error GoTo Query_Table_Method_For_TXT_Retrieval

    Combined_WB = Combined_Workbook(Variable_Sheet)
    
    FilterC = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=False)
    '^^ retrieve wanted column NUmbers
    UB = UBound(FilterC, 1)
    ReDim InfoF(1 To UB)
    
    For D = 1 To UBound(FilterC, 1) 'Fill in column numbers for use when supplying column filters to OpenTxt
        InfoF(D) = Array(D, FilterC(D))
    Next D
    
    Erase FilterC
    
    FileNumber = FreeFile
    
    For Each File_TXT In File_Collection 'Openeach file in the collection and write their contents to a string
    
        Application.StatusBar = "Parsing " & File_TXT
        DoEvents
        
        Open File_TXT For Input As FileNumber
            
            File_Name = Right$(File_TXT, Len(File_TXT) - InStrRev(File_TXT, Application.PathSeparator))
            
            If File_Name Like "*ICE*" Then 'IF name has ICE in it
            
                File_Path = Split(Input$(LOF(FileNumber), FileNumber), vbNewLine) 'split by new line characters
                
                For D = LBound(File_Path) To UBound(File_Path) 'Move the location of contract codes to the 4th field
                    
                    On Error GoTo Invalid_Row_Found
                    
                    If File_Path(D) <> vbNullString Then
                    
                        FilterC = Split(File_Path(D), Chr(44))
                        
                        If D <> LBound(File_Path) And ((InStr(1, LCase(FilterC(0)), "option") > 0 And Combined_WB = True) Or (InStr(1, LCase(FilterC(0)), "option") = 0 And Combined_WB = False)) Then
                        
                            FilterC(3) = FilterC(6)
                            FilterC(1) = DateSerial(Left(FilterC(1), 2) + 2000, Mid(FilterC(1), 3, 2), Right(FilterC(1), 2))
                            
                            If UBound(FilterC) > UB - 1 Then ReDim Preserve FilterC(LBound(FilterC) To UB - 1)
                            
                            File_Path(D) = Join$(FilterC, Chr(44))
                            
                        ElseIf D <> LBound(File_Path) Then 'Keep headers

                            File_Path(D) = vbNullString
                            
                        End If
                        
                    End If
Next_Text_Line:
                Next D
                
                Data_STR = Join$(File_Path, vbNewLine) & Data_STR
                
            Else
            
                Data_STR = Input$(LOF(FileNumber), FileNumber) & Data_STR
                
            End If
            
        Close FileNumber
        
        'If LCase(File_TXT) Like "*weekly*" Then Kill File_TXT
        
    Next File_TXT
    
On Error GoTo Query_Table_Method_For_TXT_Retrieval

    File_Name = Left$(File_Collection(1), InStrRev(File_Collection(1), Application.PathSeparator)) & "Historic.txt"
    
    Application.StatusBar = "Creating file " & File_Name
    DoEvents
    
    Open File_Name For Output As FileNumber 'Write contents of string to text File
        
        Print #FileNumber, Data_STR
        
    Close FileNumber
    
    Application.StatusBar = "TXT file compilation was successful. Creating Workbook."
    DoEvents
       
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If
    
    With Workbooks
    
            .OpenText Filename:=File_Name, origin:=D, StartRow:=1, DataType:=xlDelimited, _
                                    TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, COMMA:=True, _
                                    FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                                    Local:=False
                                    
        Set Contract_WB = Workbooks(.Count)
        
        'Contract_WB.Windows(1).Visible = False
        
    End With
    
    Contract_WB.Windows(1).Visible = False
    
    On Error Resume Next
        If Not OnMAC Then Contract_WB.SaveAs Saved_Workbook_Path, FileFormat:=xlExcel12
    On Error GoTo 0
        
'ElseIf OnMAC Then

Exit Sub

Invalid_Row_Found:
    
    File_Path(D) = vbNullString
    Resume Next_Text_Line
    
Query_Table_Method_For_TXT_Retrieval:
    
    On Error GoTo -1
    
    On Error GoTo Parent_Handler

    InfoF = Query_Text_Files(File_Collection) 'Use Querytables
    
    Application.StatusBar = "Data compilation was successful. Creating Workbook."
    DoEvents
    
    Set Contract_WB = Workbooks.Add
    
    With Contract_WB
    
        .Windows(1).Visible = False
        
        With .Worksheets(1)
        
            .DisplayPageBreaks = False
            
            .Columns("C:C").NumberFormat = "@" 'Format as text
            
            .Range("A1").Resize(UBound(InfoF, 1), UBound(InfoF, 2)).Value = InfoF
        
        End With
        
    End With
    
    Exit Sub
    
Parent_Handler:

    ThisWorkbook.Event_Storage("Historical Parse Errors").Add "An error occurred while compiling text files."
    Resume Exit_SC
    
Exit_SC:
    
    On Error GoTo 0

    Err.Raise 5

End Sub
Public Sub Historical_Excel_Creation(ByRef Contract_WB As Workbook, File_Collection As Collection, _
                                     Workbook_Path As String, Mac_User As Boolean)
                                     
'
' Creates an Excel file for historical Multi-Week data from ALL contracts
'

Dim File_IO As Variant, OpenWBS As Range, Array_CLCTN As New Collection, _
AC As Variant, Opened_Workbook As Workbook, TB As ListObject, Y As Long

Application.DisplayAlerts = False

Set Contract_WB = Workbooks.Add     'Create a new workbook

With Contract_WB

    .Windows(1).Visible = False

    With .Worksheets(1)
        
        .Columns("D:D").NumberFormat = "@" 'format as text

        .DisplayPageBreaks = False

    End With

End With

For Each File_IO In File_Collection  'open all files
    
    Application.StatusBar = "Parsing file: " & File_IO
    
    Set Opened_Workbook = Workbooks.Open(File_IO)
    
    With Opened_Workbook
        
        .Windows(1).Visible = False
        
         With .Worksheets(1).UsedRange
        
            If File_IO = File_Collection(1) Then
               AC = .Value2 'All data including header
            Else
               AC = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Value2 'Don't include header
            End If
         
         End With
         
         If Opened_Workbook.Name Like "*ICE*" Then
         
            On Error Resume Next
            
            For Y = LBound(AC, 1) To UBound(AC, 1)
                AC(Y, 4) = AC(Y, 7) 'Write contract codes to the same column as CFTC Workbooks
                AC(Y, 3) = CLng(DateSerial(Left(AC(Y, 2), 2) + 2000, Mid(AC(Y, 2), 3, 2), Right(AC(Y, 2), 2)))
                AC(Y, 188) = AC(Y, 191) 'move combined strings
                AC(Y, 191) = Empty
            Next Y
            
            On Error GoTo 0
            
        End If
        
        Array_CLCTN.Add AC
        
        Erase AC
        
        .Close False 'Close workbook stored in Path Collection
        
    End With
    
    If Not Mac_User Then If File_IO Like "*Weekly*" Then Kill File_IO
    
Next File_IO
    
On Error GoTo Error_While_Compiling

With Contract_WB
    
    With Array_CLCTN
    
        If .Count > 1 Then
            AC = Multi_Week_Addition(Array_CLCTN, Multiple_2d) 'compile all the CFTC arrays into a single array
        Else
            AC = .Item(1)
        End If
        
    End With
    
    With .Worksheets(1)
    
        .Range("A1").Resize(UBound(AC, 1), UBound(AC, 2)).Value2 = AC
        
        Set TB = .ListObjects.Add(SourceType:=xlSrcRange, Source:=.UsedRange, XlListObjectHasHeaders:=xlYes)
        
    End With
    
    If Not Mac_User Then .SaveAs Filename:=Workbook_Path, FileFormat:=xlExcel12

End With
    
With Application

    .DisplayAlerts = True

    '.StatusBar = "Excel file has been created and saved. Parsing for relevant Contract Codes."

End With

Exit Sub

Error_While_Compiling:
    
    MsgBox "An error occured while creating Excel workbook for historical data."
    
    On Error GoTo -1
    
    On Error Resume Next
    
    Contract_WB.Close: Call Re_Enable: End

End Sub
Public Sub Historical_Excel_Aggregation(Contract_WB As Workbook, _
                                        ByRef CLCTN_A As Collection, _
                                        Optional Contract_ID As String, _
                                        Optional Date_Input As Long = 0, _
                                        Optional Weekly_ICE_Contracts As Boolean = False, _
                                        Optional Specified_Contract As Boolean = False, _
                                        Optional Weekly_CFTC_TXT As Boolean = False)

'
'Filters Data for wanted contracts
'

Dim Date_Field As Long, VAR_DTA() As Variant, Valid_Codes() As Variant, Comparison_Operator As String, _
Contract_Code_CLMN As Long, Table_OBJ As ListObject, Futures_NOptions_Filter() As String, _
Contract_Code_List As Long, Futures_NOptions_List As Long, DBR As Range, OPI As Long, ICE_Codes() As String ', TT As Double

Dim Combined_CLMN As Long, Execute_Combined_Filter As Boolean, Disaggregated_Filter_STR As String 'Used if filtering ICE Contracts for Futures and Options


Dim Error_Number As Long

'TT = Timer
On Error GoTo Close_Workbook

'Err.Raise 5
Application.StatusBar = "Filtering Data."
DoEvents

With Contract_WB.Worksheets(1)
    
    If .UsedRange.Cells.Count = 1 Then 'If worksheet is empty then display message
        
        GoTo Scripts_Failed_To_Collect_Data
        
    Else
    
        If .ListObjects.Count = 0 Then
        
            If Weekly_CFTC_TXT Then 'Determine if Worksheet has headers based on if its a Text Document or not
                OPI = xlNo
            Else
                OPI = xlYes
            End If
            
            Set Table_OBJ = .ListObjects.Add(SourceType:=xlSrcRange, Source:=.UsedRange, XlListObjectHasHeaders:=OPI)
        
        Else
        
            Set Table_OBJ = .ListObjects(1)
            
        End If
        
    End If
    
End With

If Weekly_ICE_Contracts Then

    Contract_Code_CLMN = 7 'Column that holds Contract identifiers
    Date_Field = 2
    Disaggregated_Filter_STR = IIf(Combined_Workbook(Variable_Sheet), "*Combined*", "*FutOnly*")
    Execute_Combined_Filter = True
    
Else 'CFTC Contracts/Yearly ICE contracts...ICE Contracts have already had their dates converted

    Contract_Code_CLMN = 3 'Column that holds Contract identifiers
    Date_Field = 2
    
End If

With WorksheetFunction
    
    On Error Resume Next
    
    Valid_Codes = .Transpose(.Index(Data_Retrieval.Valid_Table_Info, 0, 1)) 'Contract Codes currently found in ThisWorkbook...1D Array
    
    If Err.Number <> 0 Then
    
        Valid_Codes = .Transpose(.Index(Application.Run("'" & ThisWorkbook.Name & "'!Get_Worksheet_Info"), 0, 1))  'Contract Codes currently found in ThisWorkbook...1D Array
        Err.Clear
        
    End If
    
End With

On Error GoTo Close_Workbook

With Table_OBJ
    
    Set DBR = .DataBodyRange

Check_If_Code_Exists:
    
    If Specified_Contract And IsError(Application.Match(Contract_ID, DBR.Columns(Contract_Code_CLMN).Value2, 0)) Then 'Test if contract code exists
        
        GoTo Contract_ID_Not_Found 'Prompt user for a different input if Contract code isn't available
            
    End If
    
    If Weekly_ICE_Contracts Then 'Find a column to be sorted based on the column header
    
        Combined_CLMN = Application.Match("FutOnly_or_Combined", .HeaderRowRange.Value2, 0)
        
    End If
    
    Application.AddCustomList ListArray:=Valid_Codes
    
    Contract_Code_List = Application.CustomListCount
    
    If Specified_Contract Then 'Store filter information for wanted Contract Code
                                                
        VAR_DTA = Array(Contract_Code_CLMN, UCase(Contract_ID), xlFilterValues, False)
            
    Else 'Store filter information for wanted Contract Codes
    
        VAR_DTA = Array(Contract_Code_CLMN, Valid_Codes, xlFilterValues, False)
        
    End If
    
    Erase Valid_Codes
    
    If Weekly_ICE_Contracts Or Weekly_CFTC_TXT Then 'Weekly_CFTC_TXT should be unique to CFTC Weekly Text Files at the time of writing
        Comparison_Operator = ">="
    Else
        Comparison_Operator = ">"
    End If
    
    If Weekly_ICE_Contracts Then 'Yearly ICE has already been converted when creating the Excel File
    
        Comparison_Operator = Comparison_Operator & Format(Year(Date_Input) - 2000, "00") & Format(Month(Date_Input), "00") & Format(Day(Date_Input), "00")
    Else
        Comparison_Operator = Comparison_Operator & Date_Input
        
    End If
    
    .AutoFilter.ShowAllData
    
    With .Sort 'Sort Date Field Old to New
    
        With .SortFields
            .Clear
            .Add Key:=DBR.Cells(2, Date_Field), SortOn:=xlSortOnValues, Order:=xlAscending
        End With
        
        .MatchCase = False
        .Header = xlYes
        .Apply
        
    End With
    'Filter Date Field
    DBR.AutoFilter Field:=Date_Field, Criteria1:=Comparison_Operator, Operator:=xlFilterValues
    
    With .Sort 'If ICE contracts then group
               'Group by contract Codes currently in this workbook
        With .SortFields
        
            .Clear
            
            If Execute_Combined_Filter Then 'Sort by Combined Contracts or Futures Only
                .Add Key:=DBR.Cells(2, Combined_CLMN), SortOn:=xlSortOnValues, Order:=xlAscending
            End If
            
            .Add Key:=DBR.Cells(2, Contract_Code_CLMN), SortOn:=xlSortOnValues, CustomOrder:=Contract_Code_List
            'Sort by Contract Codes
        End With
        
        .MatchCase = False
        .Header = xlYes
        .Apply
        
    End With
    
    With DBR 'Filter for "Combined" if condition met. Filter for wanted contract(s)
    
        If Execute_Combined_Filter Then .AutoFilter Field:=Combined_CLMN, Criteria1:=Disaggregated_Filter_STR, Operator:=xlFilterValues, VisibleDropDown:=False
            
        .AutoFilter Field:=VAR_DTA(0), _
                Criteria1:=VAR_DTA(1), _
                 Operator:=VAR_DTA(2), _
          VisibleDropDown:=VAR_DTA(3)

        With .SpecialCells(xlCellTypeVisible)
            
            If .Cells.Count > 1 Then
            
                VAR_DTA = .Value2
                
                If Weekly_ICE_Contracts Then  'Convert Dates from YYMMDD
                
                    For OPI = LBound(VAR_DTA, 1) To UBound(VAR_DTA, 1)
                        VAR_DTA(OPI, Date_Field) = CLng(DateSerial(Left(VAR_DTA(OPI, Date_Field), 2) + 2000, Mid(VAR_DTA(OPI, Date_Field), 3, 2), Right(VAR_DTA(OPI, Date_Field), 2)))
                    Next OPI
                    
                End If
            
                CLCTN_A.Add VAR_DTA
                
                Erase VAR_DTA
                
            End If
            
        End With 'End .SpecialCells(xlCellTypeVisible)
        
    End With 'End DBR
    
End With 'End Table_OBJ

With Application

    If Futures_NOptions_List > 0 Then .DeleteCustomList Futures_NOptions_List
    If Contract_Code_List > 0 Then .DeleteCustomList Contract_Code_List
    
    .StatusBar = vbNullString
    DoEvents
    
End With

'Debug.Print Timer - TT

Exit Sub

Close_Workbook: 'Error handler
    
    Contract_ID = Contract_WB.FullName
    Contract_WB.Close False
    
    On Error Resume Next
    
    Kill Contract_ID
    
    With Application
        If Futures_NOptions_List > 0 Then .DeleteCustomList Futures_NOptions_List
        If Contract_Code_List > 0 Then .DeleteCustomList Contract_Code_List
        Application.StatusBar = ""
    End With
    
    ThisWorkbook.Event_Storage("Historical Parse Errors").Add "Error during Historical Filtration function."
            
    Error_Number = Err.Number
    
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
        "File name: " & Contract_WB.Name
        
    Contract_ID = Contract_WB.FullName
    Contract_WB.Close False
    
    Error_Number = Err.Number
    
    On Error Resume Next
        Kill Contract_ID

    Resume Parent_Error_Handler

Parent_Error_Handler:

    On Error GoTo 0
    
    Err.Raise Error_Number 'Enter historical pars error handler
    
End Sub

Public Sub Weekly_Text_File(File_Path As Collection, ByRef StorageC As Collection, Date_Value As Long)

Dim File_IO As Variant, D As Long, FilterC() As Variant, InfoF() As Variant, _
Contract_WB As Workbook, ZZ As Long, Y As Long, Output() As Variant

FilterC = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True)

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
                            TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, COMMA:=True, _
                            FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                            Local:=False
                       
        Set Contract_WB = Workbooks(.Count)
    
    End With
    
    With Contract_WB
        
        .Windows(1).Visible = False
        
        On Error GoTo Workbook_Parse_Error
        
        Call Historical_Excel_Aggregation(Contract_WB, CLCTN_A:=StorageC, Date_Input:=Date_Value, Weekly_CFTC_TXT:=True)
    
        .Close False
        
    End With
    
    Kill File_IO

Next_File:

Next File_IO

Exit Sub

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
    
    Exit Sub

End Sub
Public Sub Filter_Market_Arrays(ByRef Contract_CLCTN As Collection, Optional ICE_Market As Boolean = False)
    
'
'Returns Filtered columns from each array in the collection
'

Dim TempB As Variant, FilterC() As Variant, T As Long, Array_Count As Long, Unknown_Filter As Boolean
      
With Contract_CLCTN

    Array_Count = .Count
    
    If Array_Count > 1 Then
        FilterC = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, ICE:=ICE_Market) '1 Based Positionl array filter
        Unknown_Filter = False
    Else
        Unknown_Filter = True
    End If
    
    For T = .Count To 1 Step -1
        
        TempB = .Item(T)
        
        .Remove T
        
        TempB = Filter_Market_Columns(Return_Filter_Columns:=False, _
                                        Return_Filtered_Array:=True, _
                                        InputA:=TempB, _
                                        ICE:=ICE_Market, _
                                        Column_Status:=FilterC, _
                                        Create_Filter:=Unknown_Filter)
                                        
        If T = .Count + 1 Then 'If last item in Collection then Simply re-add
            .Add TempB
        Else
            .Add TempB, Before:=T
        End If
        
    Next T

End With

End Sub
Public Function Filter_Market_Columns(ByVal Return_Filter_Columns As Boolean, _
                                       ByVal Return_Filtered_Array As Boolean, _
                                       Optional ByVal Create_Filter As Boolean = True, _
                                       Optional ByVal InputA As Variant, _
                                       Optional ByVal ICE As Boolean = False, _
                                       Optional ByVal Column_Status As Variant) As Variant

Dim ZZ As Long, Output() As Variant, V As Long, D As Long, Y As Long, _
Contract_ID As Long, Found_Value As Boolean

Dim CFTC_Wanted_Columns() As Variant, Dates_Column As Long, Data_Start As Long

CFTC_Wanted_Columns = Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange.Columns(2).Value2

If ICE Then
    Dates_Column = 2
    Contract_ID = 7
Else
    Dates_Column = 3
    Contract_ID = 4
End If

For ZZ = Contract_ID + 1 To UBound(CFTC_Wanted_Columns, 1)
    
    If CFTC_Wanted_Columns(ZZ, 1) = True Then
        Data_Start = ZZ
        Exit For
    End If
    
Next ZZ

If Create_Filter = True Then 'IF column Status is empty or if it is empty

    ReDim Column_Status(1 To UBound(CFTC_Wanted_Columns, 1))
    
    For ZZ = 1 To UBound(CFTC_Wanted_Columns, 1)
        
        If CFTC_Wanted_Columns(ZZ, 1) = True Or ZZ = Dates_Column Or ZZ = Contract_ID Then
            '^ allows entry into block regardless of if ICE or CFTC is needed for dates or contract code
        
            Select Case ZZ
            
                Case Dates_Column 'column 2 or 3 depending on if ICE or not
                
                    Column_Status(ZZ) = xlMDYFormat
                    
                Case 1, Contract_ID
                
                    Column_Status(ZZ) = xlTextFormat
                    
                Case 2, 3, 4, 7 'These numbers may overlap with dates column or contract field
                                'The previous cases will prevent it from executing unnecessarily depending on if ICE or not
                    Column_Status(ZZ) = xlSkipColumn
                    
                Case Else
                
                    Column_Status(ZZ) = xlGeneralFormat
                    
            End Select
            
        Else
        
            Column_Status(ZZ) = xlSkipColumn 'skip these columns
            
        End If
        
    Next ZZ
    
End If

If Return_Filter_Columns = True Then

    Filter_Market_Columns = Column_Status
    
ElseIf Return_Filtered_Array = True Then
    
     'Don't worry about text files..they are filtered in the same sub that they are opened in
     'FYI Dates_Column would be 2 if doing TXT files...2 is used for ICE contracts because of exchange inconsistency
    On Error Resume Next

    Y = 0

    Do 'Determine the total number of dimensions
    
        Y = Y + 1
        ZZ = LBound(InputA, Y)
        
    Loop Until Err.Number <> 0
    
    Select Case Y - 1
    
        Case 2 '2 Dimensions
        
            ReDim Output(1 To UBound(InputA, 1), 1 To UBound(Filter(Column_Status, xlSkipColumn, False)) + 1) 'Output Array
            
            D = Data_Start
            Y = 0
            
            For V = LBound(Column_Status) To UBound(Column_Status) 'Loop filter array and Test if skip column
                
                If Column_Status(V) <> xlSkipColumn Then
                
                    Y = Y + 1
                    
                    Select Case Y
                    
                        Case 1, 2, UBound(Output, 2) 'These elements are shifted below
                        
                        Case Else 'Open Interest is used as default value for D since it is the first value after the contract codes that is needed
                            
                            If Y > 3 Then D = D + 1      'Increment when Y is changed and not meeting previous cases
                            
                            Found_Value = False
                            
                            Do While D <= UBound(Column_Status) 'Search to right for next valid column
                            
                                If Column_Status(D) <> xlSkipColumn Then
                                    Found_Value = True
                                    Exit Do
                                End If
                                D = D + 1
                                
                            Loop
                            
                    End Select
                    
                    For ZZ = LBound(Output, 1) To UBound(Output, 1) 'Loop rows of INPUT Array and fill entire column
                        
                        Select Case Y
                        
                            Case 1
                            
                                Output(ZZ, Y) = InputA(ZZ, Dates_Column) 'Dates placed in 1st column
                                
                            Case 2
                            
                                Output(ZZ, Y) = InputA(ZZ, 1) 'Market Name -2nd Column
                                
                            Case UBound(Output, 2)
        
                                Output(ZZ, Y) = InputA(ZZ, Contract_ID)
                
                            Case Else
                                
                                If Found_Value = True Then Output(ZZ, Y) = InputA(ZZ, D) 'shift everything else left 1
                                
                        End Select
        
                    Next ZZ
                    
                End If
                
            Next V
    
        Case Else '1
        
            MsgBox "Need 2D array."
            
    End Select
    
    Filter_Market_Columns = Output
    
End If
    
End Function
Public Function Query_Text_Files(ByVal TXT_File_Paths As Collection) As Variant

'Const File As String = "C:\Users\Yliyah\Desktop\L_2017_TXT_Futures_Only.txt"

Dim QT As QueryTable, file As Variant, Found_QT As Boolean, Field_Info() As Variant, Output_Arrays As New Collection, _
Field_Info_ICE() As Variant, dd() As Variant, E As Long, Max_Filtered_Columns As Long, ICE_Data As Boolean, Z As Long

Dim Combined_WB As Boolean

Combined_WB = Combined_Workbook(Variable_Sheet) 'Determine Workbook Type Futures ONly or combined with Options
    
For Each QT In QueryT.QueryTables 'Search for the following query if it exists
    If InStr(1, QT.Name, "TXT Import") > 0 Then
        Found_QT = True
        Exit For
    End If
Next QT

Field_Info = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True) '^^ CFTC Column filter

If Data_Retrieval.TypeF = "D" Then 'ICE Data column filter
    Field_Info_ICE = Filter_Market_Columns(Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
    
    Max_Filtered_Columns = UBound(Filter(Field_Info_ICE, xlSkipColumn, False)) + 1 'number of columns that should be in the array at the end
End If

For Each file In TXT_File_Paths
    
    Application.StatusBar = "Querying: " & file
    DoEvents
    
    If Not Found_QT Then
        
        Set QT = QueryT.QueryTables.Add(Connection:="TEXT;" & file, Destination:=QueryT.Cells(1, 1))
        
        With QT
            .Name = "TXT Import"
            .BackgroundQuery = False
            .SaveData = False
            .TextFileCommaDelimiter = True
            .TextFileConsecutiveDelimiter = False
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
        End With
        
        Found_QT = True 'So that this statement isn't executed again
        
    End If
    
    With QT
        
        .Connection = "TEXT;" & file

        If file Like "*.csv" Then 'ICE Workbooks
            .TextFileColumnDataTypes = Field_Info_ICE
            ICE_Data = True
        Else
            .TextFileColumnDataTypes = Field_Info
        End If
        
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = False
        .Destination = QueryT.Cells(1, 1)
        
        .Refresh False
        
        With .ResultRange
            
            If ICE_Data Then
            
                dd = .Value2
                
                On Error Resume Next
                
                    For E = 1 To UBound(dd, 1)

                        If E <> 1 And (InStr(1, LCase(dd(E, 1)), "option") > 0 And Combined_WB) Or (InStr(1, LCase(dd(E, 1)), "option") = 0 And Not Combined_WB) Then
                        
                            dd(E, 2) = CLng(DateSerial(Left(dd(E, 2), 2) + 2000, Mid(dd(E, 2), 3, 2), Right(dd(E, 2), 2))) 'convert dates to proper format
                        
                        ElseIf E <> 1 Then 'Arrray Row doesn't match Workbook Type
                        
                            For Z = 1 To UBound(dd, 2)
                                dd(E, Z) = Empty
                            Next Z
                            
                        End If
                        
                    Next E
                
                On Error GoTo 0
                
                If UBound(dd, 2) <> Max_Filtered_Columns Then ReDim Preserve dd(1 To UBound(dd, 1), 1 To Max_Filtered_Columns)
                
                Output_Arrays.Add dd
                
                ICE_Data = False
                
            Else
            
                Output_Arrays.Add .Value2
                
            End If
            
            .ClearContents
            
        End With
    
    End With

Next file

'Rearranging of array elements is done in Historical_Parse

If Output_Arrays.Count > 1 Then
    Query_Text_Files = Multi_Week_Addition(Output_Arrays, Data_Retrieval.Multiple_2d)
Else
    Query_Text_Files = Output_Arrays(1)
End If

QT.Delete

End Function
Private Sub Edit_Columns_List()

Dim ColumnsL() As Variant, X As Long, Z As Long, Table_RNG As Range, FilterA() As Variant

Set Table_RNG = Variable_Sheet.ListObjects("User_Selected_Columns").DataBodyRange

ColumnsL = Table_RNG.Value2

FilterA = Filter_Market_Columns(True, False, True, , False)

For X = 1 To UBound(ColumnsL, 1)

    Z = 1
    Do Until Not IsNumeric(Mid(ColumnsL(X, 1), Z, 1)) And Not Mid(ColumnsL(X, 1), Z, 1) = " "
        Z = Z + 1
    Loop
    
    ColumnsL(X, 1) = Right(ColumnsL(X, 1), Len(ColumnsL(X, 1)) - (Z - 1))
    ColumnsL(X, 2) = IIf(FilterA(X) = xlSkipColumn, False, True)
    
Next X

Table_RNG.Value = ColumnsL

End Sub
Public Sub Retrieve_Tuesdays_CLose(ByRef Table_Data_Addition As Variant, ByVal Price_Column As Long, ByVal Contract_Code_Column As Long, Workbook_INfo As Variant)

'
'Retrieves Price data for dates in column 1 of an array
'

Dim Use_QueryTable As Boolean, Y As Long, Start_Date As Date, End_Date As Date, URL As String, _
Year_1970 As Date, Symbol As String, X As Long, Yahoo_Finance_Parse As Boolean, Stooq_Parse As Boolean

Dim Price_Data() As String, Initial_Split_CHR As String, D_OHLC_AV() As String, Quandl_API_KEY As String, QuandL_Parse As Boolean

Dim Close_Price As Long, Secondary_Split_STR As String, Response_STR As String, QT_Connection_Type As String

Dim End_Date_STR As String, Start_Date_STR As String, Query_Name As String

Dim QT As QueryTable, QueryTable_Found As Boolean, Using_QueryTable As Boolean, Query_Data() As Variant

Const Symbol_Column_in_Workbook_Info As Long = 5

On Error GoTo Exit_Price_Parse

With Application
    
    Symbol = UCase(Workbook_INfo(.Match(Table_Data_Addition(UBound(Table_Data_Addition, 1), Contract_Code_Column), .Index(Workbook_INfo, 0, 1), 0), Symbol_Column_in_Workbook_Info))
    
End With

If Symbol = vbNullString Then Exit Sub

Start_Date = Table_Data_Addition(1, 1)

End_Date = Table_Data_Addition(UBound(Table_Data_Addition, 1), 1)

If Symbol Like "*=F" Or Symbol Like "*=X" Or Symbol Like "^*" Then 'CSV File

    Yahoo_Finance_Parse = True
    
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
    
ElseIf Symbol Like "*.F" Then 'CSV file

    Stooq_Parse = True
    
    Query_Name = "Stooq Query"
    
    End_Date_STR = Format(End_Date, "yyyymmdd")
    Start_Date_STR = Format(Start_Date, "yyyymmdd")
    
    URL = "https://stooq.com/q/d/l/?s=" & Symbol & "&d1=" & Start_Date_STR & "&d2=" & End_Date_STR & "&i=d"
    
    QT_Connection_Type = "URL;"
    
ElseIf Symbol Like "*CHRIS/*" Then
    
    QuandL_Parse = True
    
    Query_Name = "QuandL Query"
    
    End_Date_STR = Format(End_Date, "yyyy-mm-dd;@")
    
    Start_Date_STR = Format(Start_Date, "yyyy-mm-dd;@")
    
    Quandl_API_KEY = Range("QuandL_Key").Value
    
    If Len(Quandl_API_KEY) = 0 Then Exit Sub
    
    'column index=4 means that only dates and close price will be imported
    URL = "https://www.quandl.com/api/v3/datasets/" & Symbol & "1?column_index=4&order=asc&" & _
    "start_date=" & Start_Date_STR & _
    "&end_date=" & End_Date_STR & _
    "&api_key=" & Quandl_API_KEY
    
    QT_Connection_Type = "TEXT;"
    
    Quandl_API_KEY = vbNullString
    
Else

    Exit Sub
    
'    YAhoo_Finance_Parse = True
'    QT_Connection_Type = "TEXT;"
    
End If

End_Date_STR = vbNullString
Start_Date_STR = vbNullString
    
#If Mac Then

    On Error GoTo Exit_Price_Parse
    'On Error GoTo 0
    Using_QueryTable = True

    For Each QT In QueryT.QueryTables           'Determine if QueryTable Exists
        
        If InStr(1, QT.Name, Query_Name) > 0 Then 'Instr method used in case Excel appends a number to the name
            QueryTable_Found = True
            Exit For
        End If
        
    Next QT
    
    If Not QueryTable_Found Then Set QT = QueryT.QueryTables.Add(QT_Connection_Type & URL, QueryT.Cells(1, 1))
    
    With QT
    
        If Not QueryTable_Found Then
        
            .BackgroundQuery = False
            .Name = Query_Name
            
            On Error GoTo Workbook_Connection_Name_Already_Exists 'deletes the connection has the name and then rename
            
                .WorkbookConnection.Name = Replace$(Query_Name, "Query", "Prices")
                
            On Error GoTo Exit_Price_Parse
            
        Else
            .Connection = QT_Connection_Type & URL
        End If
        
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SaveData = False
        
        If QuandL_Parse Then
            .TextFileCommaDelimiter = False
            .TextFileConsecutiveDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileOtherDelimiter = "["
        End If
        
        On Error GoTo Remove_QT_And_Connection 'Delete both the Querytable and the connection and exit the sub

         .Refresh False
        
        On Error GoTo Exit_Price_Parse
        
        With .ResultRange
            
            If Yahoo_Finance_Parse Or Stooq_Parse Or QuandL_Parse Then 'an array of csv values not separated
            
                Query_Data = .Value
                
            End If
            
            .ClearContents
            
        End With
        
    End With
    
    Set QT = Nothing
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
        
        If InStr(1, Response_STR, 404) = 1 Then Exit Sub 'Something likely wrong with the URl
        
        If Yahoo_Finance_Parse Then
            
            Initial_Split_CHR = Mid$(Response_STR, InStr(1, Response_STR, "Volume") + Len("volume"), 1) 'Finding Splitting_Charachter
        
        ElseIf Stooq_Parse Then
        
            Initial_Split_CHR = vbNewLine
            
        End If
        
        Price_Data = Split(Response_STR, Initial_Split_CHR)
           
    Else
    
        ReDim Price_Data(0 To UBound(Query_Data, 1) - 1) 'redim to fit all rows of the query array
         
        For X = 0 To UBound(Query_Data, 1) - 1 'Add everything  to array
            Price_Data(X) = Query_Data(X + 1, 1)
        Next X
        
        Erase Query_Data
        
    End If
    
    Secondary_Split_STR = Chr(44)
    
    X = LBound(Price_Data) + 1 'Skip headers
    
    Close_Price = 4 'Base 0 location of close prices within the queried array
    
ElseIf QuandL_Parse Then
    
    If Not Using_QueryTable Then
        
        If InStr(1, Response_STR, "quandl_error") > 0 Then Exit Sub
  
        Response_STR = Split(Response_STR, "[[" & Chr(34))(1)
        Response_STR = Split(Response_STR, "]]")(0)

        Price_Data = Split(Response_STR, "]" & Chr(44) & "[" & Chr(34)) '],["
        
        Secondary_Split_STR = Chr(34) & Chr(44) 'split subsequent elements with    ",
        
    Else
        
        ReDim Price_Data(0 To UBound(Query_Data, 2) - 4) 'need base 0 and then skip first 3 values
        
        For X = 0 To UBound(Price_Data) 'Array elements have a string begginning with "]" attached to the end ,so remove that string
        
            Price_Data(X) = Left$(Query_Data(1, X + 4), InStr(1, Query_Data(1, X + 4), "]") - 1)
            
        Next X
        
        Erase Query_Data
        
        Secondary_Split_STR = Chr(44) 'price_data elements with a commaa
        
    End If
    
    Close_Price = 1 'Base 0 location of closing prices within array D_OHLC_AV
    
    X = 0
    
End If

If Len(Response_STR) > 0 Then Response_STR = vbNullString
If Len(Initial_Split_CHR) > 0 Then Initial_Split_CHR = vbNullString

Y = 1

Start_Date = CDate(CDate(Left$(Price_Data(X), InStr(1, Price_Data(X), Secondary_Split_STR) - 1)))

Do Until Table_Data_Addition(Y, 1) >= Start_Date 'Loop until a matching date is found from the input date with the minimum data in the queried prices

    If Y + 1 <= UBound(Table_Data_Addition, 1) Then
        Y = Y + 1
    Else
        Exit Sub
    End If
    
Loop

For Y = Y To UBound(Table_Data_Addition, 1)

    On Error GoTo Error_While_Splitting
    
    Do Until Start_Date >= Table_Data_Addition(Y, 1) 'Loop until price dates meet or exceed wanted date
    '>= used in case there isnt  a price for the requested date
Increment_X:

        X = X + 1
        
        If X > UBound(Price_Data) Then
            Exit Sub 'Exits Main Loop
        Else
            Start_Date = CDate(Left$(Price_Data(X), InStr(1, Price_Data(X), Secondary_Split_STR) - 1))
        End If
        
    Loop

    If Start_Date = Table_Data_Addition(Y, 1) Then 'IF wanted date is found
    
        D_OHLC_AV = Split(Price_Data(X), Secondary_Split_STR)
        
        On Error Resume Next
        
        If Not IsNumeric(D_OHLC_AV(Close_Price)) Then 'find first value that came before that isn't empty
        
            Table_Data_Addition(Y, Price_Column) = Empty
            
        ElseIf CDbl(D_OHLC_AV(Close_Price)) = 0 Then
        
            Table_Data_Addition(Y, Price_Column) = Empty
            
        Else
        
            Table_Data_Addition(Y, Price_Column) = CDbl(D_OHLC_AV(Close_Price))
            
        End If
        
Ending_INcrement_X:
        
        Erase D_OHLC_AV
        
        If X + 1 <= UBound(Price_Data) Then
            X = X + 1
        Else
            Exit Sub
        End If
    
    End If
    
Next Y

Exit_Price_Parse:

    Erase Price_Data

Exit Sub

Remove_QT_And_Connection:
    
    QT.Delete
    
    Exit Sub
    
Workbook_Connection_Name_Already_Exists:

    ThisWorkbook.Connections(Replace(Query_Name, "Query", "Prices")).Delete
    
    QT.WorkbookConnection.Name = Replace(Query_Name, "Query", "Prices")
    Resume Next

Error_While_Splitting:

    If Err.Number = 13 Then 'type mismatch error from using cdate on a non-date string
        Resume Increment_X
    Else
        Exit Sub
    End If
    
End Sub
Public Function Detect_Old_To_New(Table_O As ListObject, Column_Key As Long) As Boolean
'
'Determine Sort order of Table Object
'
Dim SF As SortFields, g As Long, Do_Manual_Check As Boolean

Set SF = Table_O.Sort.SortFields

If SF.Count > 0 Then

    For g = 1 To SF.Count
    
        With SF(g)
        
            If .Key.Column = Column_Key Then
            
                Select Case .Order
                    Case xlDescending
                        Detect_Old_To_New = False
                    Case xlAscending
                        Detect_Old_To_New = True
                    Case Else
                        Do_Manual_Check = True
                End Select
                
                If Do_Manual_Check = False Then
                    Exit Function
                Else
                    Exit For
                End If
            End If
            
        End With
        
    Next g
    
Else
    Do_Manual_Check = True
End If

On Error GoTo G_Too_Large

If Do_Manual_Check = True Then

    With Table_O.DataBodyRange.SpecialCells(xlCellTypeVisible)
        For g = 1 To 3
            If .Cells(g, Column_Key) > .Cells(g + 1, Column_Key) Then 'if greater than the cell below it
                Exit Function
            End If
        Next g
    End With
    
End If

Detect_Old_To_New = True

Exit Function

G_Too_Large:
    If Err.Number = 9 And g > 2 Then
        Detect_Old_To_New = True
    Else
        Detect_Old_To_New = True 'Default to true...possible place holder
    End If
    
End Function
Public Function Reverse_2D_Array(Data_I As Variant) As Variant

Dim Z As Long, B As Long, OT() As Variant

ReDim OT(1 To UBound(Data_I, 1), 1 To UBound(Data_I, 2))

For Z = UBound(Data_I, 1) To 1 Step -1
    For B = 1 To UBound(Data_I, 2)
        OT(UBound(Data_I, 1) - Z + 1, B) = Data_I(Z, B)
    Next B
Next Z

Reverse_2D_Array = OT

End Function
Sub Sort_Table(TB As ListObject, Key As Long, Order As Long)

    With TB.Sort
    
        With .SortFields
            .Clear
            .Add Key:=TB.Range.Cells(1, Key), SortOn:=xlSortOnValues, Order:=Order
        End With
        
        .MatchCase = False
        .Header = xlYes
        .Apply
        
    End With

End Sub

