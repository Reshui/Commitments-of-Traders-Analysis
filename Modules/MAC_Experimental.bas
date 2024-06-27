Attribute VB_Name = "MAC_Experimental"
Const mDelimiter$ = "*"
Const AppleScriptFileName$ = "COT_HelperScripts_MoshiM.scpt"
Const posixPathSeparator$ = "/"
Option Explicit
' https://stackoverflow.com/questions/41946262/mac-excel-generate-utf-8
#If Mac Then

    #If VBA7 Then
        ' 64 bit Office:mac
        Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
        Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
        Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As LongPtr
        Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
        Private file As LongPtr
    #Else
        ' 32 bit Office:mac
        Private Declare Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As Long
        Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
        Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal items As Long, ByVal stream As Long) As Long
        Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
        Private file As Long
        
    #End If
    
    Public Function ExecuteShellCommandMAC(command As String, Optional ByRef exitCode As Long) As String
        
        On Error GoTo Propogate
        file = popen(command, "r")
        
        If file = 0 Then
          Exit Function
        End If
        
        While feof(file) = 0
          Dim chunk As String
          Dim read As Long
          chunk = Space(50)
          read = fread(chunk, 1, Len(chunk) - 1, file)
          If read > 0 Then
            chunk = Left$(chunk, read)
            ExecuteShellCommandMAC = ExecuteShellCommandMAC & chunk
          End If
        Wend
        
        exitCode = pclose(file) ' 0 = success
        Exit Function
Propogate:
        PropagateError e, "ExecuteShellCommandMAC"
    End Function
    
    Function writeToFile(str As String, fileName As String)
        'escape double quotes
        str = Replace(str, """", "\\\""")
    
        ' use Apple Script and shell commands to create and write file
        MacScript ("do shell script ""printf '" & str & "'> " & fileName & " "" ")
    
        ' print file path
        Debug.Print "file path: " & MacScript("do shell script ""pwd""") & "/" & fileName
    End Function

#End If
Sub DownloadFileMAC(fileUrl$, savedFileName$)

    Dim argList$(), returnCode As Byte
    
    ReDim argList(1)
    
    argList(0) = savedFileName
    argList(1) = fileUrl
    
    argList = QuotedForm(argList)

    'Where Script File needs to be
    '~/Library/Application Scripts/[bundle id]/
    '~/Library/Application Scripts/com.microsoft.Excel/COT_HelperScripts_MoshiM.applescript

    'Folder that I can download and access files from
    'Environ("HOME") =/Users/rondebruin/Library/Containers/com.microsoft.Excel/Data

    On Error GoTo PossibleErrorInArguements

    #If MAC_OFFICE_VERSION >= 15 Then
        Dim argSubmission$
        
        argSubmission = mDelimiter & Join(argList, mDelimiter)
        returnCode = AppleScriptTask(AppleScriptFileName, "DownloadFile", argSubmission)
        
    #Else
        
        Dim shellCommand$
        
        shellCommand = "set result to (do shell script "" curl -Lo " & argList(0) & " " & argList(1) & """)" & vbNewLine & "return result"
        '--follow redirections and output to a file
        returnCode = MacScript(shellCommand)
        
    #End If

PossibleErrorInArguements:
    If returnCode <> 0 Or Err.Number <> 0 Then
        MsgBox "An error occured while attempting to download a file." _
        & vbNewLine & vbNewLine & _
        "Arguements:" & vbNewLine & _
        argList(0) & vbNewLine & argList(1)

        Re_Enable
        End
    End If

End Sub
Sub UnzipFile(zipFullPath$, directoryToExtractTo$, ByVal filesToExtract As Variant)
    
    'https://docs.oracle.com/cd/E88353_01/html/E37839/unzip-1.html
    
    Dim argList$(), returnCode As Byte

    ReDim argList(2)
    
    argList(0) = zipFullPath
    argList(1) = filesToExtract
    argList(2) = directoryToExtractTo

    argList = QuotedForm(argList)
    
    argList(1) = Join(argList(1), " ")


    #If MAC_OFFICE_VERSION >= 15 Then
        
        Dim argSubmission$
        
        argSubmission = mDelimiter & Join(argList, mDelimiter)
        
        returnCode = AppleScriptTask(AppleScriptFileName, "UnzipFiles", argSubmission)
        
    #Else
    
        Dim shellCommand$
        
        shellCommand = "set result to (do shell script ""/usr/bin/unzip -uao " & argList(0) & " " & argList(1) & " -d " & argList(2) & """)" & vbNewLine & "return result"
        
        returnCode = MacScript(shellCommand)
        
    #End If
    'do shell script "unzip -d /Users/abc/Desktop/ /Users/abc/Desktop/1/2/3/5/abc.zip"

PossibleErrorInArguements:

    If returnCode <> 0 Or Err.Number <> 0 Then
        MsgBox "An error occured while attempting to unzip a file." _
        & vbNewLine & vbNewLine & _
        "Arguements:" & vbNewLine & _
        argList(0) & vbNewLine & argList(1) & vbNewLine & argList(2)

        Re_Enable
        End
    End If
    
End Sub
Public Sub CreateRootDirectories(folderPath$)

    Dim folderHiearchy$(), partialRootFolder$(), currentFolderDepth As Byte, currentFolderName$

    currentFolderName = folderPath
    ' Each folder will need to be created if it doesn't exist.
    ' Store folder names in array.
    folderHiearchy = Split(folderPath, posixPathSeparator)
    
    partialRootFolder = folderHiearchy
    ' Set currentFolderDepth equal to the last value
    ' Go in reverse order untila valid folder is found.
    currentFolderDepth = UBound(folderHiearchy)
    
    'Loop until a valid folder path is found or if the next loop is out of bounds.
    Do Until FileOrFolderExists(currentFolderName) Or currentFolderDepth - 1 < LBound(partialRootFolder)
        
        currentFolderDepth = currentFolderDepth - 1

        ReDim Preserve partialRootFolder(LBound(partialRootFolder) To currentFolderDepth)

        currentFolderName = Join(partialRootFolder, posixPathSeparator)

    Loop
    
    If Not FileOrFolderExists(currentFolderName) Then
        MsgBox "Error in Mac Root directory creation step."
        Re_Enable
        End
    End If
        
    Do While currentFolderDepth < UBound(folderHiearchy)
        currentFolderDepth = currentFolderDepth + 1
        currentFolderName = currentFolderName & posixPathSeparator & folderHiearchy(currentFolderDepth)
        MkDir currentFolderName
    Loop

End Sub

Function BasicMacAvailablePath$()
    
    Dim OfficeFolder$

    OfficeFolder = MacScript("return POSIX path of (path to library folder) as string")
    
    If Right$(OfficeFolder, 1) <> "/" Then OfficeFolder = OfficeFolder & "/"
    
    OfficeFolder = OfficeFolder & "Group Containers/UBF8T346G9.Office"

End Function

Function CreateFolderinMacOffice(NameFolder$) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 13-July-2020
    Dim OfficeFolder$
    Dim PathToFolder$
    Dim TestStr$

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
        "/Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir$(PathToFolder & "*", vbDirectory)
    On Error GoTo 0
    If LenB(TestStr) = 0 Then
        MkDir PathToFolder
        'You can use this msgbox line for testing if you want
        'MsgBox "You find the new folder in this location :" & PathToFolder
    End If
    CreateFolderinMacOffice = PathToFolder
End Function

'Function FileOrFolderExistsOnYourMac(FileOrFolderstr$, FileOrFolder As Long) As Boolean
'    'Ron de Bruin : 13-Dec-2020, for Excel 2016 and higher
'    'Function to test if a file or folder exist on your Mac
'    'Use 1 as second argument for File and 2 for Folder
'    Dim ScriptToCheckFileFolder$
'    Dim FileOrFolderPath$
'
'    If FileOrFolder = 1 Then
'        'File test
'        On Error Resume Next
'        FileOrFolderPath = Dir$(FileOrFolderstr & "*")
'        On Error GoTo 0
'        If LenB(FileOrFolderPath) > 0 Then FileOrFolderExistsOnYourMac = True
'    Else
'        'folder test
'        On Error Resume Next
'        FileOrFolderPath = Dir$(FileOrFolderstr & "*", vbDirectory)
'        On Error GoTo 0
'        If LenB(FileOrFolderPath) > 0 Then FileOrFolderExistsOnYourMac = True
'    End If
'
'End Function
Public Function DetermineIfScriptableMAC() As Boolean

    #If MAC_OFFICE_VERSION >= 15 Then
        
        If FileOrFolderExists(ReturnFullPathToScriptFileMAC) Then
            DetermineIfScriptableMAC = True
        End If
        
    #Else
        DetermineIfScriptableMAC = True
    #End If
        
End Function

Public Function ReturnFullPathToScriptFileMAC$()
    
    Dim temp$
    temp = MacScript("return POSIX path of (path to library folder) as string")
     
    If Right$(temp, 1) <> "/" Then temp = temp & "/"
    
    ReturnFullPathToScriptFileMAC = temp & "Application Scripts/com.microsoft.Excel/" & AppleScriptFileName
    
End Function
Public Sub GateMacAccessToWorkbook()
    
    Dim stopScripts As Boolean
    
    #If Mac Then
    
        #If Not DatabaseFile Then
'            If Not DetermineIfScriptableMAC() Then
'                MsgBox "Couldn't find File : " & ReturnFullPathToScriptFileMAC & vbNewLine & vbNewLine & _
'                       "Please download the script file from the DropBox folder and place it in the given location. Create the file path if necessary."
'                stopScripts = True
'            End If
        #Else
            stopScripts = True
            MsgBox "File is unavailable to MAC users. Use an alternate version available in the DropBox folder."
        #End If
        
        If stopScripts Then
            Re_Enable
            End
        End If
        
    #End If
    
End Sub
