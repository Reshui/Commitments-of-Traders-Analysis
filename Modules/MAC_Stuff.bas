Attribute VB_Name = "MAC_Stuff"
'Private Function Select_File_Or_Files_Mac()
'
'    Dim MyPath As String
'    Dim MyScript As String
'    Dim MyFiles As String
'    Dim MySplit As Variant
'    Dim N As Long
'    Dim FName As String
'    Dim mybook As Workbook
'
'    On Error Resume Next
'
'    MyPath = MacScript("return (path to documents folder) as String")
'    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"
'
'    ' In the following statement, change true to false in the line "multiple
'    ' selections allowed true" if you do not want to be able to select more
'    ' than one file. Additionally, if you want to filter for multiple files, change
'    ' {""com.microsoft.Excel.xls""} to
'    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
'    ' if you want to filter on xls and csv files, for example.
'    MyScript = _
'    "set applescript's text item delimiters to "","" " & vbNewLine & _
'               "set theFiles to (choose file of type " & _
'             " {""public.comma-separated-values-text""} " & _
'               "with prompt ""Please select a file or files"" default location alias """ & _
'               MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
'               "set applescript's text item delimiters to """" " & vbNewLine & _
'               "return theFiles"
'
'    MyFiles = MacScript(MyScript)
'
'    Dim returnList() As String
'
'    On Error GoTo 0
'
'    If MyFiles <> "" Then
'
'        With Application
'            .ScreenUpdating = False
'            .EnableEvents = False
'        End With
'
'        'MsgBox MyFiles
'        MySplit = Split(MyFiles, ",")
'        ReDim returnList(LBound(MySplit) To UBound(MySplit))
'        For N = LBound(MySplit) To UBound(MySplit)
'
'            returnList(N) = MySplit(N)
'
'        Next N
'
'        With Application
'            .ScreenUpdating = True
'            .EnableEvents = True
'        End With
'
'        Select_File_Or_Files_Mac = returnList
'
'    Else
'        ReDim returnList(0 To 0)
'        returnList(0) = "False"
'        Select_File_Or_Files_Mac = returnList
'    End If
'
'End Function

Public Function Select_Folder_On_Mac() As String

    Dim folderPath As String
    Dim RootFolder As String
    Dim scriptstr As String

    On Error Resume Next
    RootFolder = MacScript("return (path to desktop folder) as String")
    'Or use RootFolder = "Macintosh HD:Users:YourUserName:Desktop:TestMap:"
    'Note : for a fixed path use : as seperator in 2011 and 2016

    If Val(Application.Version) < 15 Then
        scriptstr = "(choose folder with prompt ""Select the folder""" & _
            " default location alias """ & RootFolder & """) as string"
    Else
        scriptstr = "return posix path of (choose folder with prompt ""Select the folder""" & _
            " default location alias """ & RootFolder & """) as string"
    End If

    folderPath = MacScript(scriptstr)
    On Error GoTo 0

    Select_Folder_On_Mac = folderPath
    
End Function
'*******Function that do all the work that will be called by the macro*********

Function GetFilesOnMacWithOrWithoutSubfolders(ByVal Folder_Path As String, Level As Long, ExtChoice As Long, _
                                              FileFilterOption As Long, FileNameFilterStr As String) As String
'Ron de Bruin,Version 4.0: 27 Sept 2015
'http://www.rondebruin.nl/mac.htm
'Thanks to DJ Bazzie Wazzie and Nigel Garvey(posters on MacScripter)
    Dim ScriptToRun As String

    Dim FileNameFilter As String
    Dim Extensions As String

    If Folder_Path = "" Then Exit Function

    Select Case ExtChoice
        Case 0: Extensions = "(xls|xlsx|xlsm|xlsb)"  'xls, xlsx , xlsm, xlsb
        Case 1: Extensions = "xls"    'Only  xls
        Case 2: Extensions = "xlsx"    'Only xlsx
        Case 3: Extensions = "xlsm"    'Only xlsm
        Case 4: Extensions = "xlsb"    'Only xlsb
        Case 5: Extensions = "csv"    'Only csv
        Case 6: Extensions = "txt"    'Only txt
        Case 7: Extensions = ".*"    'All files with extension, use *.* for everything
        Case 8: Extensions = "(xlsx|xlsm|xlsb)"  'xlsx, xlsm , xlsb
        Case 9: Extensions = "(csv|txt)"   'csv and txt files
        'You can add more filter options if you want,
    End Select

    Select Case FileFilterOption
        Case 0: FileNameFilter = "'.*/[^~][^/]*\\." & Extensions & "$' "  'No Filter
        Case 1: FileNameFilter = "'.*/" & FileNameFilterStr & "[^~][^/]*\\." & Extensions & "$' "    'Begins with
        Case 2: FileNameFilter = "'.*/[^~][^/]*" & FileNameFilterStr & "\\." & Extensions & "$' "    ' Ends With
        Case 3: FileNameFilter = "'.*/([^~][^/]*" & FileNameFilterStr & "[^/]*|" & FileNameFilterStr & "[^/]*)\\." & Extensions & "$' "   'Contains
    End Select

    Folder_Path = MacScript("tell text 1 thru -2 of " & Chr(34) & Folder_Path & _
                           Chr(34) & " to return quoted form of it's POSIX Path")
    Folder_Path = Replace(Folder_Path, "'\''", "'\\''")

    If Val(Application.Version) < 15 Then
        ScriptToRun = ScriptToRun & "set foundPaths to paragraphs of (do shell script """ & "find -E " & _
                      Folder_Path & " -iregex " & FileNameFilter & "-maxdepth " & _
                      Level & """)" & Chr(13)
        ScriptToRun = ScriptToRun & "repeat with thisPath in foundPaths" & Chr(13)
        ScriptToRun = ScriptToRun & "set thisPath's contents to (POSIX file thisPath) as text" & Chr(13)
        ScriptToRun = ScriptToRun & "end repeat" & Chr(13)
        ScriptToRun = ScriptToRun & "set astid to AppleScript's text item delimiters" & Chr(13)
        ScriptToRun = ScriptToRun & "set AppleScript's text item delimiters to return" & Chr(13)
        ScriptToRun = ScriptToRun & "set foundPaths to foundPaths as text" & Chr(13)
        ScriptToRun = ScriptToRun & "set AppleScript's text item delimiters to astid" & Chr(13)
        ScriptToRun = ScriptToRun & "foundPaths"
    Else
        ScriptToRun = ScriptToRun & "do shell script """ & "find -E " & _
                      Folder_Path & " -iregex " & FileNameFilter & "-maxdepth " & _
                      Level & """ "
    End If
    
    On Error Resume Next
    GetFilesOnMacWithOrWithoutSubfolders = MacScript(ScriptToRun)
    On Error GoTo 0
    
End Function



