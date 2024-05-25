Attribute VB_Name = "MAC_Experimental"
Const mDelimiter As String = "*"
Const AppleScriptFileName As String = "COT_HelperScripts_MoshiM.scpt"
Const posixPathSeparator As String = "/"

Sub DownloadFile(fileUrl As String, savedFileName As String)

    Dim argList() As String, returnCode As Byte
    
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
        Dim argSubmission As String
        
        argSubmission = mDelimiter & Join(argList, mDelimiter)
        returnCode = AppleScriptTask(AppleScriptFileName, "DownloadFile", argSubmission)
        
    #Else
        
        Dim shellCommand As String
        
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
Sub UnzipFile(zipFullPath As String, directoryToExtractTo As String, ByVal filesToExtract As Variant)
    
    'https://docs.oracle.com/cd/E88353_01/html/E37839/unzip-1.html
    
    Dim argList() As String, returnCode As Byte

    ReDim argList(2)
    
    argList(0) = zipFullPath
    argList(1) = filesToExtract
    argList(2) = directoryToExtractTo

    argList = QuotedForm(argList)
    
    argList(1) = Join(argList(1), " ")


    #If MAC_OFFICE_VERSION >= 15 Then
        
        Dim argSubmission As String
        
        argSubmission = mDelimiter & Join(argList, mDelimiter)
        
        returnCode = AppleScriptTask(AppleScriptFileName, "UnzipFiles", argSubmission)
        
    #Else
    
        Dim shellCommand As String
        
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
Public Function QuotedForm(ByRef Item, Optional Enclosing_CHR As String = """") As Variant

    Dim Z As Long, subArray As Variant, subArrayIndex As Byte
    
    If IsArray(Item) Then
    
        For Z = LBound(Item) To UBound(Item)
        
            If IsArray(Item(Z)) Then
            
                subArray = Item(Z)
                
                For subArrayIndex = LBound(subArray) To UBound(subArray)
                    If Not subArray(subbarayindex) Like Enclosing_CHR & "*" & Enclosing_CHR Then subArray(subbarayindex) = Enclosing_CHR & subArray(subbarayindex) & Enclosing_CHR
                Next subArrayIndex
                    
                Item(Z) = subArray
                
            Else
                If Not Item(Z) Like Enclosing_CHR & "*" & Enclosing_CHR Then Item(Z) = Enclosing_CHR & Item(Z) & Enclosing_CHR
            End If
                     
        Next Z
        
    Else
        If Not Item Like Enclosing_CHR & "*" & Enclosing_CHR Then Item = Enclosing_CHR & Item & Enclosing_CHR
    End If
    
    QuotedForm = Item
    
End Function

Public Sub CreateRootDirectories(folderPath As String)

    Dim folderHiearchy() As String, partialRootFolder() As String, currentFolderDepth As Byte, currentFolderName As String

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

Function BasicMacAvailablePath() As String
    
    Dim OfficeFolder As String

    OfficeFolder = MacScript("return POSIX path of (path to library folder) as string")
    
    If Right$(OfficeFolder, 1) <> "/" Then OfficeFolder = OfficeFolder & "/"
    
    OfficeFolder = OfficeFolder & "Group Containers/UBF8T346G9.Office"

End Function

Function CreateFolderinMacOffice(NameFolder As String) As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 13-July-2020
    Dim OfficeFolder As String
    Dim PathToFolder As String
    Dim TestStr As String

    OfficeFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    
    OfficeFolder = Replace(OfficeFolder, "/Desktop", "") & _
        "/Library/Group Containers/UBF8T346G9.Office/"

    PathToFolder = OfficeFolder & NameFolder

    On Error Resume Next
    TestStr = Dir(PathToFolder & "*", vbDirectory)
    On Error GoTo 0
    If LenB(TestStr) = 0 Then
        MkDir PathToFolder
        'You can use this msgbox line for testing if you want
        'MsgBox "You find the new folder in this location :" & PathToFolder
    End If
    CreateFolderinMacOffice = PathToFolder
End Function

Function FileOrFolderExistsOnYourMac(FileOrFolderstr As String, FileOrFolder As Long) As Boolean
    'Ron de Bruin : 13-Dec-2020, for Excel 2016 and higher
    'Function to test if a file or folder exist on your Mac
    'Use 1 as second argument for File and 2 for Folder
    Dim ScriptToCheckFileFolder As String
    Dim FileOrFolderPath As String
    
    If FileOrFolder = 1 Then
        'File test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr & "*")
        On Error GoTo 0
        If LenB(FileOrFolderPath) > 0 Then FileOrFolderExistsOnYourMac = True
    Else
        'folder test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr & "*", vbDirectory)
        On Error GoTo 0
        If LenB(FileOrFolderPath) > 0 Then FileOrFolderExistsOnYourMac = True
    End If

End Function
Public Function DetermineIfScriptableMAC() As Boolean

    #If MAC_OFFICE_VERSION >= 15 Then
        
        If FileOrFolderExists(ReturnFullPathToScriptFileMAC) Then
            DetermineIfScriptableMAC = True
        End If
        
    #Else
        DetermineIfScriptableMAC = True
    #End If
        
End Function

Public Function ReturnFullPathToScriptFileMAC() As String
    
    Dim temp As String
    temp = MacScript("return POSIX path of (path to library folder) as string")
     
    If Right$(temp, 1) <> "/" Then temp = temp & "/"
    
    ReturnFullPathToScriptFileMAC = temp & "Application Scripts/com.microsoft.Excel/" & AppleScriptFileName
    
End Function
Public Sub GateMacAccessToWorkbook()
    
    Dim stopScripts As Boolean
    
    #If Not DatabaseFile And Mac Then
        If Not DetermineIfScriptableMAC() Then
            MsgBox "Couldn't find File : " & ReturnFullPathToScriptFileMAC & vbNewLine & vbNewLine & _
                   "Please download the script file from the DropBox folder and place it in the given location. Create the file path if necessary."
            stopScripts = True
        End If
    #ElseIf Mac Then
        stopScripts = True
        MsgBox "File is unavailable to MAC users."
    #End If
    
    If stopScripts Then
        Re_Enable
        End
    End If
    
End Sub
Function FileOrFolderExists(FileOrFolderstr As String) As Boolean
'Ron de Bruin : 1-Feb-2019
'Function to test whether a file or folder exist on a Mac in office 2011 and up
'Uses AppleScript to avoid the problem with long names in Office 2011,
'limit is max 32 characters including the extension in 2011.
    Dim ScriptToCheckFileFolder As String
    Dim TestStr As String
    
    #If Not Mac Then
        FileOrFolderExists = LenB(Dir(FileOrFolderstr, vbDirectory)) > 0
        Exit Function
    #End If
    
    If Val(Application.Version) < 15 Then
        ScriptToCheckFileFolder = "tell application " & QuotedForm("System Events") & _
        "to return exists disk item (" & QuotedForm(FileOrFolderstr) & " as string)"
        FileOrFolderExists = MacScript(ScriptToCheckFileFolder)
    Else
        On Error Resume Next
        TestStr = Dir(FileOrFolderstr & "*", vbDirectory)
        On Error GoTo 0
        If LenB(TestStr) > 0 Then FileOrFolderExists = True
    End If

End Function




