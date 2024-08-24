VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonParserB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' ============================================================================
''' JSON parser for VBA.
''' https://github.com/Reshui/VBA-JSON-Parser
''' ============================================================================
''' @author: MoshiM_UC@outlook.com
''' ============================================================================
''' MIT License (http://www.opensource.org/licenses/mit-license.php)
'''
''' Copyright (c) 2024 Yliyah Miller
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================
#If Mac Then

    #If VBA7 Then
        ' 64 bit Office:mac
        Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
        Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
        Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As LongPtr
        Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr
        'Private file As LongPtr
    #Else
        ' 32 bit Office:mac
        Private Declare Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As Long
        Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
        Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal items As Long, ByVal stream As Long) As Long
        Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long
        'Private file As Long
    #End If

#ElseIf VBA7 Then
    Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As LongPtr
    Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTime As SystemTime) As LongPtr
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As LongPtr
#Else
    Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedTimeUTC As SystemTime) As Long
    Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (tzInfo As TimeZoneInformation, localTime As SystemTime, returnedLocalTimeUTC As SystemTime) As Long
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (tzInfo As TimeZoneInformation) As Long
#End If

Private Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TimeZoneInformation
    Bias As Long 'The bias is the difference, in minutes, between Coordinated Universal Time (UTC) and local time.
    standardName(31) As Integer
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SystemTime
    DaylightBias As Long
    'TimeZoneKeyName(127) As Integer
    'DynamicDaylightTimeDisabled As Boolean
End Type

Option Explicit

Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const INDEX_OFFSET As Byte = 1
Private Const RightCurly$ = "}", LeftCurly$ = "{", LeftSquare$ = "[", RightSquare$ = "]", DecimalChar$ = ".", _
QuoteCharacter$ = """", BackslashCharacter$ = "\", SpaceCharacter$ = " ", ColonCharacter$ = ":", Comma$ = ",", BINARY_OFFSET As Byte = 2

Private pConvertStampsToDate As Boolean
Private pSpecialCharactersA$(6)

Public pConvertDatesToUTC As Boolean
Public pConvertToLocalTime As Boolean

#Const DebugActive = False

Private Sub Class_Initialize()
    pSpecialCharactersA(0) = QuoteCharacter: pSpecialCharactersA(1) = BackslashCharacter
    pSpecialCharactersA(2) = vbCr: pSpecialCharactersA(3) = vbTab
    pSpecialCharactersA(4) = vbLf: pSpecialCharactersA(5) = VBA.vbBack: pSpecialCharactersA(6) = VBA.vbFormFeed
End Sub

Public Function Deserialize(ByRef json$, Optional convertTimeStampsToDate As Boolean = False, Optional convertDatesToUTC As Boolean = False, Optional convertUtcToLocal As Boolean) As Object
'========================================================================================================================================
' Summary : Deserializes a json string.
' Paramaters :
'   [convertTimeStampToDate] - Set to True to auto attempt converting timestamps to dates.
'       It is recommended that you submit a value of False as some values may be converted when you don't want them to be ex: "2000" will be converted to year 2000.
'   [convertDatesToUTC] - True if you want to convert time stamps to UTC if a timezone offset is provided.
'   [json] - JSON string to parse.
'========================================================================================================================================
    pConvertStampsToDate = convertTimeStampsToDate
    pConvertDatesToUTC = convertDatesToUTC
    pConvertToLocalTime = convertUtcToLocal
    
    Set Deserialize = ParseJsonString(json)
End Function
Public Function DeserializeBytes(ByRef jsonBytes() As Byte, Optional convertTimeStampsToDate As Boolean = False, Optional convertDatesToUTC As Boolean = False, Optional convertUtcToLocal As Boolean) As Object
    'Warning bytes must be UTF-16
    pConvertStampsToDate = convertTimeStampsToDate
    pConvertDatesToUTC = convertDatesToUTC
    pConvertToLocalTime = convertUtcToLocal
    Set DeserializeBytes = ParseJsonString(CStr(jsonBytes))
End Function

Private Function ParseJsonString(ByVal json$) As Object
'====================================================================================================================================================================================
' Summary:
'   Non-Recursive JSON parser to convert a JSON string into a dictionary/collection.
' Parameters:
'      [json] - JSON string to parse.
' Returns:
'   Either a Collection or Dictionary
' Reference:
'   https://www.json.org/json-en.html
' Note:
'   If on MAC ensure you have downloaded the Dictionary class from https://github.com/VBA-tools/VBA-Dictionary
'====================================================================================================================================================================================
    Dim value$, isValueJsonObject As Boolean, isValueJsonArray As Boolean, isContainerJsonObject As Boolean, isValueAvailable As Boolean

    Dim jsonLength&, keyForValue$, valueStartIndex&, valueEndIndex&, jsonContainer As Object, _
    isValidatedString As Boolean, heirarchyStack As New Collection, cursorLocation&, nextCursorLocation&, isKeyAvailable As Boolean
    
    Const cursorLimit$ = "MoshiM_UC@outlook.com"
    Const cursorLocationAfterCompletion$ = "https://github.com/Reshui?tab=repositories"
    Const arrayPattern$ = "[[]*[]]", objectPattern$ = "{*}"
    On Error GoTo CATCH_FAILURE
    ' This is basically a trim that also works on vbLf,vbTab,vbCr and doesn't return a new string just start and end locations.
    If TryGetWhitespaceCleanedIndexes(json, 1, Len(json), cursorLocation, jsonLength) Then
        If Mid$(json, cursorLocation, 1) = LeftCurly And Mid$(json, jsonLength, 1) = RightCurly Then
            isValueJsonObject = True
        ElseIf Mid$(json, cursorLocation, 1) = LeftSquare And Mid$(json, jsonLength, 1) = RightSquare Then
            isValueJsonArray = True
        End If
    End If
    
    If isValueJsonObject Or isValueJsonArray Then

        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
        Set jsonContainer = heirarchyStack(1)
        
        AddValueToObject jsonContainer, isValueJsonObject, addingCursorLimit:=True, cursorLimit:=jsonLength, keyForValue:=cursorLimit
        AddValueToObject jsonContainer, isValueJsonObject, addingCursorLimit:=True, cursorLimit:=jsonLength, keyForValue:=cursorLocationAfterCompletion

        isContainerJsonObject = isValueJsonObject
        
        Do While cursorLocation < jsonLength
            ' For certain JSON strings when an object/array was closed on the previous loop then the
            ' cursor may be set to the next end of an object/array but still have more text to parse.
            ' and therefore cursorLocation = jsonContainer(cursorLocationAfterCompletion) = jsonContainer(cursorLimit)
            ' Ex: }},{"hello":1}
            '     >^
            If cursorLocation <> jsonContainer.Item(cursorLimit) Then
                ' Finds the next token that encloses a value in conjunction with the current cursorLocation.
                nextCursorLocation = GetNextEnclosingDelimiterLocation(json, cursorLocation, jsonContainer.Item(cursorLimit))
                
                If nextCursorLocation > cursorLocation Then
                    If isContainerJsonObject Then
                        ' Adjust by INDEX_OFFSET to ignore enclosing token delimiters.
                        GetValueRangeAndKey json, cursorLocation + INDEX_OFFSET, nextCursorLocation - INDEX_OFFSET, valueStartIndex, valueEndIndex, isValueAvailable, keyForValue, isKeyAvailable
                    End If
        
                    If Not isContainerJsonObject Or (isKeyAvailable And isContainerJsonObject) Then
                                    
                        If Not isContainerJsonObject Then
                            'Adjust by INDEX_OFFSET to exclude token delimiters.
                            isValueAvailable = TryGetWhitespaceCleanedIndexes(json, cursorLocation + INDEX_OFFSET, nextCursorLocation - INDEX_OFFSET, valueStartIndex, valueEndIndex)
                        End If
                                            
                        If isValueAvailable Then
                            'If value is enclosed in {} or [] then parse a new object or array.
                            isValueJsonObject = Mid$(json, valueStartIndex, valueEndIndex - valueStartIndex + 1) Like objectPattern
                            If Not isValueJsonObject Then isValueJsonArray = Mid$(json, valueStartIndex, valueEndIndex - valueStartIndex + 1) Like arrayPattern
                        End If
    
                        If isValueAvailable And (isValueJsonObject Or isValueJsonArray) Then
                            
                            CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
                            isValueJsonObject = False: isValueJsonArray = False
                            ' Create heirarchy by nesting child within parent.
                            With heirarchyStack
                                AddValueToObject jsonContainer, isContainerJsonObject, newChild:=.Item(.count), keyForValue:=keyForValue
                                Set jsonContainer = .Item(.count)
                            End With
    
                            isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                            ' Add key to limit the cursor range for object.
                            AddValueToObject jsonContainer, isContainerJsonObject, addingCursorLimit:=True, cursorLimit:=valueEndIndex, keyForValue:=cursorLimit
                            ' Index to update cursorLocation to once array or object has been completed.
                            AddValueToObject jsonContainer, isContainerJsonObject, addingCursorLimit:=True, cursorLimit:=nextCursorLocation, keyForValue:=cursorLocationAfterCompletion
                            cursorLocation = valueStartIndex
                        Else
                            If isValueAvailable Then
                                ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                                ' Throws an error if any other unescaped quotes.
                                value = ProcessCharactersInRange(json, valueStartIndex, valueEndIndex, isValidatedString)
                            Else
                                ' Code will convert this to null in AddValueToObject() below.
                                value = vbNullString: isValidatedString = False
                            End If
                            AddValueToObject jsonContainer, isContainerJsonObject, value, isValueString:=isValidatedString, keyForValue:=keyForValue
                            If isValueAvailable Then value = vbNullString
                            cursorLocation = nextCursorLocation
                        End If
    
                        If isKeyAvailable Then
                            keyForValue = vbNullString: isKeyAvailable = False
                        End If
                    Else
                        'Should already be a raised error in GetValueRangeAndKey().
                        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="JSON parser failed. Key unavailable for object."
                    End If
                End If
            End If
            
            With jsonContainer
                ' If cursorLocation has exceeded jsonContainer Cursor Limit; remove it from the stack.
                If cursorLocation >= .Item(cursorLimit) Then
                    cursorLocation = .Item(cursorLocationAfterCompletion)
                    With heirarchyStack
                        ' Only update if not Parent of all nodes.
                        If .count > 1 Then
                            With jsonContainer
                                .Remove cursorLimit
                                .Remove cursorLocationAfterCompletion
                            End With
                            .Remove .count
                            Set jsonContainer = .Item(.count)
                            isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                        End If
                    End With
                End If
            End With
            
        Loop
        
        With jsonContainer
            .Remove cursorLimit
            .Remove cursorLocationAfterCompletion
        End With
        Set ParseJsonString = jsonContainer
        
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Json isn't an object > {*} or array > [*]."
    End If

    Exit Function
CATCH_FAILURE:
    #If DebugActive Then
        Stop: Resume
    #Else
        PropagateError Err, "ParseJsonString"
    #End If
End Function
Public Sub Unit_Test()
    Dim success As Boolean, json$, ob As Object
    With Me
        '.ParseJsonString "{""apple"":""tree"",""inner object"":1,""array"":[1,2,3,{""apple"":22,""berry"":45}]}"
        '.ParseJsonString "{""ground\n\\\""beaver"":""tr\\ee"",""inner object"":[{""APP"":1,""Sun"":""Moon"",""Ocean"":""Beach""},[""a"",""b"",""c""],[0,1,2,[5,6,7,8],9,10]],""array"":[1,2,3,{""apple"":22,""berry"":45}]}"
        Set ob = .Deserialize("{""L"":{""True"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""724 ms"",""Status"":0,""Combined"":""True"",""Id"":""Legacy""},""False"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""0 ms"",""Status"":0,""Combined"":""False"",""Id"":""Legacy""}},""D"":{""True"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""0 ms"",""Status"":0,""Combined"":""True"",""Id"":""Disaggregated""},""False"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""0 ms"",""Status"":0,""Combined"":""False"",""Id"":""Disaggregated""}},""T"":{""True"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""0 ms"",""Status"":0,""Combined"":""True"",""Id"":""TFF""},""False"":{""Latest Date"":""2024-06-25"",""Time Elapsed"":""0 ms"",""Status"":0,""Combined"":""False"",""Id"":""TFF""}}}", True)
        Debug.Print Serialize(ob, False, True)
        
        '.ParseJsonString HttpGet("http://worldtimeapi.org/api/timezone/America/New_York", success)
    End With
End Sub
Public Function Serialize(container As Object, encloseStringsInEscapedQuotes As Boolean, Optional prettyPrint As Boolean = False, Optional ByVal indentLevel& = 0) As String
'=======================================================================================================================================
' Summary:
'   Converts a Collection or Dictionary to a Json string.
' Parameters:
'   [container] - Object to convert.
'   [encloseStringsInEscapedQuotes] - True to enclose text in escaped quotes instead of unescaped quotes.
'        Example: {\"hello\":1} instead of {"hello":1}
'=======================================================================================================================================
    Dim output$, cursorLocation&, isContainerJsonObject As Boolean, keyOrIndex() As Variant, valuesAvailalbe As Boolean, _
    iCount&, isValueString As Boolean, variableQuote$, currentIndent&, whitespaceL1$, whitespaceL2$, isValueDate As Boolean
    
    variableQuote = IIf(encloseStringsInEscapedQuotes, "\""", """")
    
    output = String$(1000, " ")
    valuesAvailalbe = True
    
    On Error GoTo NoValuesAvailable
    
    If TypeName(container) = "Dictionary" Then
        isContainerJsonObject = True
        keyOrIndex = container.keys
    ElseIf TypeOf container Is Collection Then
        ReDim keyOrIndex(1 To container.count)
        For iCount = 1 To UBound(keyOrIndex)
            keyOrIndex(iCount) = iCount
        Next iCount
    Else
        On Error GoTo 0
        Err.Raise 13, "Serialize", "Container must be either a Collection or Dictionary."
    End If
    
DetermineIndent:
    On Error GoTo 0
    If prettyPrint Then
        If indentLevel <> 0 Then whitespaceL1 = Space$(indentLevel * 4)
        whitespaceL2 = Space$((indentLevel + 1) * 4)
    End If
    
    If isContainerJsonObject Then
        'If indentLevel <> 0 Then OverwriteBufferSegment output, whitespaceL1, cursorLocation
        OverwriteBufferSegment output, LeftCurly, cursorLocation
    Else
        OverwriteBufferSegment output, LeftSquare, cursorLocation
    End If
    
    If prettyPrint Then OverwriteBufferSegment output, vbNewLine, cursorLocation
    
    If valuesAvailalbe Then
        With container
            For iCount = LBound(keyOrIndex) To UBound(keyOrIndex)
                
                If prettyPrint Then OverwriteBufferSegment output, whitespaceL2, cursorLocation
                
                If isContainerJsonObject Then
                    OverwriteBufferSegment output, variableQuote, cursorLocation
                    OverwriteBufferSegment output, CStr(keyOrIndex(iCount)), cursorLocation, True
                    OverwriteBufferSegment output, variableQuote, cursorLocation
                    OverwriteBufferSegment output, ColonCharacter, cursorLocation
                End If
                    
                If IsObject(.Item(keyOrIndex(iCount))) Then
                    If Not isContainerJsonObject And prettyPrint Then
                        OverwriteBufferSegment output, vbNewLine, cursorLocation
                    End If
                    OverwriteBufferSegment output, Serialize(.Item(keyOrIndex(iCount)), encloseStringsInEscapedQuotes, prettyPrint, indentLevel + 1), cursorLocation
                Else
                
                    isValueDate = IsDate(.Item(keyOrIndex(iCount)))
                    isValueString = isValueDate Or VarType(.Item(keyOrIndex(iCount))) = vbString
                               
                    If isValueString Then OverwriteBufferSegment output, variableQuote, cursorLocation
                    
                    Select Case .Item(keyOrIndex(iCount))
                        Case True, False
                            OverwriteBufferSegment output, LCase$(.Item(keyOrIndex(iCount))), cursorLocation, isValueString
                        Case Else
                            If IsNull(.Item(keyOrIndex(iCount))) Then
                                OverwriteBufferSegment output, "null", cursorLocation, isValueString
                            Else
                                OverwriteBufferSegment output, IIf(isValueDate, Format$(.Item(keyOrIndex(iCount)), "yyyy-mm-ddThh:mm:ss"), .Item(keyOrIndex(iCount))), cursorLocation, isValueString
                            End If
                    End Select
                    
                    If isValueString Then OverwriteBufferSegment output, variableQuote, cursorLocation
                    
                End If
    
                If iCount <> UBound(keyOrIndex) Then
                    OverwriteBufferSegment output, Comma, cursorLocation
                End If
                
                If prettyPrint Then OverwriteBufferSegment output, vbNewLine, cursorLocation
                
            Next iCount
        End With
    End If
    
    If prettyPrint And indentLevel <> 0 Then OverwriteBufferSegment output, whitespaceL1, cursorLocation
    OverwriteBufferSegment output, IIf(isContainerJsonObject, RightCurly, RightSquare), cursorLocation
    
    Serialize = Left$(output, cursorLocation)
    Exit Function
    
NoValuesAvailable:
    valuesAvailalbe = False
    Resume DetermineIndent
End Function

Private Sub CreateObjectOnStack(stack As Collection, createJsonObject As Boolean)

    If createJsonObject Then
        stack.Add GetNewDictionary()
    Else
        stack.Add New Collection
    End If

End Sub

Private Function GetNextEnclosingDelimiterLocation(json$, cursorLocation&, maxCursorLocation&) As Long
'=======================================================================================================================================
' Summary:
'   Returns the location of the next token index that encapsulates a complete value based on [cursorLocation].
'   Searches for commas and determines if an object or array is between that comma and [cursorLocation]
'   If an array or object exists, determine which comes first and loop until bounds of object are found, then search for one of ,]}
' Parameters:
'   [json] - String to search.
'   [cursorLocation] - Location to start searching from.
'   [maxCursorLocation] - Max returnable index.
'=======================================================================================================================================
    Dim nextSquareCloseBinary&, nextSquareOpenBinary&, nextValidCommaTokenIndexBinary&, currentIndexBinary&, nextOpenBinary&, nextCloseBinary&, minValues(1) As Long
    
    Dim maxIndexBinary&, nextCurlyCloseBinary&, nextCurlyOpenBinary&, previousCloseBinary&, tokenIndicator$(1), nextEnclosingTokenIndex&, escapeEnclosingArrayOrObject As Boolean

    On Error GoTo Propagate
    currentIndexBinary = ConvertToBinaryIndex(cursorLocation)
    maxIndexBinary = ConvertToBinaryIndex(maxCursorLocation)
    
    nextCurlyOpenBinary = SeekNextUnquotedBinaryCharacterIndex(json, LeftCurly, currentIndexBinary + BINARY_OFFSET, maxIndexBinary, True)
    nextCurlyCloseBinary = SeekNextUnquotedBinaryCharacterIndex(json, RightCurly, currentIndexBinary + BINARY_OFFSET, maxIndexBinary, True)
    nextSquareOpenBinary = SeekNextUnquotedBinaryCharacterIndex(json, LeftSquare, currentIndexBinary + BINARY_OFFSET, maxIndexBinary, True)
    nextSquareCloseBinary = SeekNextUnquotedBinaryCharacterIndex(json, RightSquare, currentIndexBinary + BINARY_OFFSET, maxIndexBinary, True)
    nextValidCommaTokenIndexBinary = SeekNextUnquotedBinaryCharacterIndex(json, Comma, currentIndexBinary + BINARY_OFFSET, maxIndexBinary, True)
    
    If nextValidCommaTokenIndexBinary > 0 Then
        '======================================================================================================================================
        ' Determine which of [ or { is greater than cursorLocation AND appears first.
        ' If nextValidCommaTokenIndexBinary is greater than that byte,
        ' then determine if the Comma is encapsulated. If it is, then find the next Comma that isn't encapsulated to find the end of the value.
        '======================================================================================================================================
        If IsValueInRange(nextCurlyOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary, False) And (nextCurlyOpenBinary < nextSquareOpenBinary Or nextSquareOpenBinary = 0) Then
            tokenIndicator(0) = LeftCurly: tokenIndicator(1) = RightCurly
            nextOpenBinary = nextCurlyOpenBinary: nextCloseBinary = nextCurlyCloseBinary
            escapeEnclosingArrayOrObject = True
        ElseIf IsValueInRange(nextSquareOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary, False) And (nextSquareOpenBinary < nextCurlyOpenBinary Or nextCurlyOpenBinary = 0) Then
            tokenIndicator(0) = LeftSquare: tokenIndicator(1) = RightSquare
            nextOpenBinary = nextSquareOpenBinary: nextCloseBinary = nextSquareCloseBinary
            escapeEnclosingArrayOrObject = True
        End If
    
        If escapeEnclosingArrayOrObject Then
            ' Initialize previousCloseBinary to the next token End character found after the current position in [json].
            ' Update the next token end and open position.
            ' If open greater than the previous close or next open = 0 then you have exited the object/array that triggered the loop.
            Do
                previousCloseBinary = nextCloseBinary
                nextCloseBinary = SeekNextUnquotedBinaryCharacterIndex(json, tokenIndicator(1), previousCloseBinary + BINARY_OFFSET, maxIndexBinary, True)
                nextOpenBinary = SeekNextUnquotedBinaryCharacterIndex(json, tokenIndicator(0), nextOpenBinary + BINARY_OFFSET, maxIndexBinary, True)
            Loop Until nextOpenBinary > previousCloseBinary Or nextOpenBinary = 0 Or (previousCloseBinary >= maxIndexBinary)

            If previousCloseBinary < maxIndexBinary Then
                ' Find MIN not equal to 0 of the following.
                nextEnclosingTokenIndex = SeekNextUnquotedBinaryCharacterIndex(json, Comma, previousCloseBinary, maxIndexBinary, True)
                
                minValues(0) = SeekNextUnquotedBinaryCharacterIndex(json, RightCurly, previousCloseBinary + BINARY_OFFSET, maxIndexBinary, True)
                minValues(1) = SeekNextUnquotedBinaryCharacterIndex(json, RightSquare, previousCloseBinary + BINARY_OFFSET, maxIndexBinary, True)

                For nextCloseBinary = LBound(minValues) To UBound(minValues)
                    If IsValueInRange(minValues(nextCloseBinary), 1, nextEnclosingTokenIndex, True) Or nextEnclosingTokenIndex <= 0 Then nextEnclosingTokenIndex = minValues(nextCloseBinary)
                Next nextCloseBinary
                
                If nextEnclosingTokenIndex > 0 Then nextEnclosingTokenIndex = ConvertFromBinaryIndex(nextEnclosingTokenIndex)
            Else
                nextEnclosingTokenIndex = maxCursorLocation
            End If
        Else
            nextEnclosingTokenIndex = ConvertFromBinaryIndex(nextValidCommaTokenIndexBinary)
        End If
    End If

    Select Case nextEnclosingTokenIndex
        Case Is <= 0, cursorLocation, Is > maxCursorLocation
            nextEnclosingTokenIndex = maxCursorLocation
    End Select
    
    If InStrB(",]}", Mid$(json, nextEnclosingTokenIndex, 1)) = 0 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid JSON format. Expected one of ',' or ']' or '}'. Calculated token index doesn't reference a value delimiter: " & Mid$(json, nextEnclosingTokenIndex, 1)
    End If
    
    GetNextEnclosingDelimiterLocation = nextEnclosingTokenIndex
    Exit Function
    
Propagate:
    #If DebugActive Then
        Stop: Resume
    #Else
        PropagateError Err, "GetNextEnclosingDelimiterLocation"
    #End If

End Function

Private Function IsValueInRange(numberToCheck&, lowerBound&, upperBound&, inclusiveUpperBound As Boolean) As Boolean
'Summary Checks if numberToCheck is contained within a given range.
    If inclusiveUpperBound Then
        IsValueInRange = numberToCheck >= lowerBound And numberToCheck <= upperBound
    Else
        IsValueInRange = numberToCheck >= lowerBound And numberToCheck < upperBound
    End If
End Function

Private Sub GetValueRangeAndKey(ByRef json$, ByRef minimumIndex&, ByRef maxSearchIndex&, ByRef valueStartLocation&, ByRef valueEndLocation&, ByRef isValueAvailable As Boolean, ByRef keyForValue$, ByRef isKeyAvailalble As Boolean)
'============================================================================================================================================================================================================================================
'Summary: Attempts to retrieve a key within the bounds of [minimumIndex] and [maxSearchIndex]. If successful then the indexes for the start and end of the value will be stored in
'   [Byref: valueStartLocation] and [ByRef:valueEndLocation]
'Variables:
'   [keyForValue] Used to store key if found.
'============================================================================================================================================================================================================================================
    Dim keyValueDelimiterIndex&, keyStartLocation&, keyEndLocation&
    
    On Error GoTo PropagateError
    'If non whitespace characters exist then record their location; otherwise throw an error.
    If TryGetWhitespaceCleanedIndexes(json, minimumIndex, maxSearchIndex, keyStartLocation, valueEndLocation) Then
        ' Does key start with a qukote?
        If Mid$(json, keyStartLocation, 1) = QuoteCharacter Then

            keyValueDelimiterIndex = ConvertFromBinaryIndex(SeekNextUnquotedBinaryCharacterIndex(json, ColonCharacter, keyStartLocation, valueEndLocation, False))
            ' Is there a ColonCharacter outside of quotes?
            If keyValueDelimiterIndex > 0 Then
                ' Find boundaries of key using location of key-value delimiter.
                Call TryGetWhitespaceCleanedIndexes(json, keyStartLocation, keyValueDelimiterIndex - INDEX_OFFSET, keyStartLocation, keyEndLocation)
                ' Find boundaries of value.
                isValueAvailable = TryGetWhitespaceCleanedIndexes(json, keyValueDelimiterIndex + INDEX_OFFSET, valueEndLocation, valueStartLocation, valueEndLocation)
                
                If (keyEndLocation < valueStartLocation) Then
                    ' Will raise error if invalid key.
                    keyForValue = ProcessCharactersInRange(json, keyStartLocation, keyEndLocation, isKeyAvailalble)
                    If Not isKeyAvailalble Then Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key must be enclosed in quotes." & Mid$(json, keyStartLocation, keyEndLocation - keyStartLocation + 1)
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key overlaps with value." & Mid$(json, keyStartLocation, keyEndLocation - valueEndLocation + 1)
                End If
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key-Value delimiter not found." & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
            End If
        Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key doesn't start with a quote or couldn't find valid enclosing token: " & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
        End If
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Expected Key-Value pair not found."
    End If

    isKeyAvailalble = True
    Exit Sub
PropagateError:
    PropagateError Err, "GetValueRangeAndKey"
End Sub

Private Sub AddValueToObject(containerObject As Object, isContainerJsonObject As Boolean, Optional value$, _
                Optional isValueString As Boolean, Optional newChild As Object, Optional keyForValue$, _
                Optional addingCursorLimit As Boolean, Optional cursorLimit As Long)
'======================================================================================================================================
' Summary
'       Adds a value to a dictionary or collection.
' Parameters:
'       [containerObject] - Object that will have a value added to it.
'       [isContainerJsonObject] - True if [containerObject] is a Dictionary.
'       [value] - Finalized value to add to dictionary.
'       [isValueString] - True if vale should be stored as a string.
'       [newChild] - Alternate object to add to [containerObject] instead of [value]
'       [keyForValue] - Key for [value] when [containerObject] is a Dictionary.
'       [addingCursorLimit] - True if adding a temporary number to track a cursor limit.
'       [cursorLimit] - Tracked cursor limit.
'======================================================================================================================================
    On Error GoTo PropagateError
    
    Dim isoDate As Date, dateParseFailed As Boolean: dateParseFailed = True
    
    If newChild Is Nothing Then
        If addingCursorLimit Then
            ' This is just used for tracking byte locations. key is guaranteed
            If isContainerJsonObject Then
                containerObject.Item(keyForValue) = cursorLimit
            Else
                containerObject.Add cursorLimit, keyForValue
            End If
        ElseIf isValueString Then
        
            If pConvertStampsToDate Then
                If TryParseIsoTimestamp(value, pConvertToLocalTime, pConvertDatesToUTC, isoDate) Then
                    If isContainerJsonObject Then
                        containerObject.Item(keyForValue) = isoDate
                    Else
                        containerObject.Add isoDate
                    End If
                    dateParseFailed = False
                End If
            End If
            
            If dateParseFailed Then
                If isContainerJsonObject Then
                    containerObject.Item(keyForValue) = value
                Else
                    containerObject.Add value
                End If
            End If
        ElseIf IsNumeric(value) Then
            If isContainerJsonObject Then
                containerObject.Item(keyForValue) = IIf(InStrB(1, value, DecimalChar) <> 0, CDbl(value), CLng(value))
            Else
                containerObject.Add IIf(InStrB(1, value, DecimalChar) <> 0, CDbl(value), CLng(value))
            End If
        Else
            Select Case LCase$(value)
                Case "true", "false"
                    If isContainerJsonObject Then
                        containerObject.Item(keyForValue) = CBool(value)
                    Else
                        containerObject.Add CBool(value)
                    End If
                Case "null"
                    If isContainerJsonObject Then
                        containerObject.Item(keyForValue) = Null
                    Else
                        containerObject.Add Null
                    End If
                Case Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid format: Value isn't numeric and failed string literal checks. Ensure it's enclosed in unescaped quotes: " & value
            End Select
        End If
    Else
        If isContainerJsonObject Then
            Set containerObject.Item(keyForValue) = newChild
        Else
            containerObject.Add newChild
        End If
    End If

    Exit Sub
PropagateError:
    PropagateError Err, "AddValueToObject"
End Sub

Private Function ProcessCharactersInRange(ByRef json$, ByRef startIndex&, ByRef maximumAllowedIndex&, Optional ByRef isRangeQuoteEnclosed As Boolean) As String
'========================================================================================================================================
'Summary: Converts special characters in the range of [startIndex] to [maximumAllowedIndex].
'   For example unescaped quotes that don't enclose a value or no special character following a \.
'Paramaters:
'   json - json string to search.
'   startIndex - Starting index within <parmaref name="json"/>
'   maximumAllowedIndex - Max index to search for string segments.
'   isRangeQuoteEnclosed - ByRef boolean that will have a value assigned to it. True if returned text had enclosinng quotes pruned.
'Note:
' Front end of [json] will be used as a buffer.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes BackslashCharacter
'==============================================================================================================================================
' To DO hex to utf-16 for mac users since chrw doesn't work.
' https://stackoverflow.com/a/75787820/12287457; https://github.com/guwidoe/VBA-StringTools
' https://stackoverflow.com/questions/56008191/get-unicode-characters-with-charcode-values-greater-hex-ffff
'==============================================================================================================================================
    Dim iCount&, currentBufferIndex&, isArrayOrObject As Boolean, currentCharacter$, minValues&(6), _
    groupedBackslashCount&, countOfBackslashToKeep&, specialCharactersInRange As Boolean, _
    unescapedQuoteCount As Long, iLoop&

    ' Determine if there are any special characters in the range.
    For iCount = LBound(pSpecialCharactersA) To UBound(pSpecialCharactersA)
        ' If within bounds set specialCharactersInRange equal to true and exit the loop.
        If IsValueInRange(InStrB(ConvertToBinaryIndex(startIndex), json, pSpecialCharactersA(iCount)), 1, ConvertToBinaryIndex(maximumAllowedIndex), True) Then
            specialCharactersInRange = True: Exit For
        End If
    Next iCount

    On Error GoTo PropagateErr
    
    If specialCharactersInRange Then

        Dim characterIndex&, regionStart&, tempIndexBinary&, previousSpecialIsValid As Boolean

        regionStart = startIndex: currentBufferIndex = 0
        ' - index_offset as it will be incremented again.
        characterIndex = startIndex - INDEX_OFFSET
        previousSpecialIsValid = True
                
        Do While characterIndex < maximumAllowedIndex
            
            iLoop = iLoop + 1

            characterIndex = ConvertToBinaryIndex(characterIndex + INDEX_OFFSET)
            tempIndexBinary = characterIndex
            ' Determine min special character index > 0.
            For iCount = LBound(pSpecialCharactersA) To UBound(pSpecialCharactersA)
                ' If (special character exists in string or on first loop) and stored location less than current location.
                If (minValues(iCount) > 0 Or iLoop = 1) And minValues(iCount) < characterIndex Then
                    minValues(iCount) = InStrB(tempIndexBinary, json, pSpecialCharactersA(iCount))
                End If
                ' If between (1,characterIndex]
                If IsValueInRange(minValues(iCount), 1, characterIndex, False) Or iCount = LBound(pSpecialCharactersA) Then
                    characterIndex = minValues(iCount)
                End If
                
            Next iCount

            characterIndex = ConvertFromBinaryIndex(characterIndex)
            
            If characterIndex > maximumAllowedIndex Or characterIndex <= 0 Then
                characterIndex = maximumAllowedIndex
            End If
            
            Select Case Mid$(json, characterIndex, 1)

                Case QuoteCharacter
                
                    unescapedQuoteCount = unescapedQuoteCount + 1
                    Call OverwriteBufferSegment(json, Mid$(json, regionStart, 1 + characterIndex - regionStart), currentBufferIndex, False)
                    previousSpecialIsValid = True
                    
                Case BackslashCharacter
                    
                    If (characterIndex - regionStart) > 0 Then
                        Call OverwriteBufferSegment(json, Mid$(json, regionStart, characterIndex - regionStart), currentBufferIndex, False)
                    End If
                    
                    groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        characterIndex = characterIndex + INDEX_OFFSET
                    Loop Until Mid$(json, characterIndex, 1) <> BackslashCharacter Or characterIndex > maximumAllowedIndex
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    ' Write backslashes to buffer.
                    If countOfBackslashToKeep > 0 Then
                        Call OverwriteBufferSegment(json, String$(countOfBackslashToKeep, BackslashCharacter), currentBufferIndex, False)
                    End If
                    
                    If characterIndex <= maximumAllowedIndex Then
                        ' If there is an even number of backslashes then all of them are escaped and the following character doesn't need additional processing.
                        If (groupedBackslashCount Mod 2) <> 0 Then
                            
                            currentCharacter = Mid$(json, characterIndex, 1)
                            currentBufferIndex = currentBufferIndex + INDEX_OFFSET
                            
                            Select Case currentCharacter
                                Case "b" '\b
                                    Mid$(json, currentBufferIndex, 1) = VBA.vbBack  'backspace
                                Case "n" '\n
                                    Mid$(json, currentBufferIndex, 1) = vbNewLine 'New line
                                Case "t" '\t
                                    Mid$(json, currentBufferIndex, 1) = vbTab  'tab
                                Case "0" 'Null character.
                                    currentBufferIndex = currentBufferIndex - INDEX_OFFSET
                                Case QuoteCharacter, "'", "/" '\" or \/ or \0 or \'
                                    Mid$(json, currentBufferIndex, 1) = currentCharacter
                                Case "r" '\r
                                    Mid$(json, currentBufferIndex, 1) = vbCr 'carriage return
                                Case "f" '\f
                                    Mid$(json, currentBufferIndex, 1) = vbFormFeed ' Form feed
                                Case "u" '\u
                                    ' See: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chr-function
                                    On Error Resume Next
                                    #If Mac Then
                                        Mid$(json, currentBufferIndex, 1) = WorksheetFunction.Unichar(val("&h" & Mid$(json, characterIndex + 1, 4)))
                                    #Else
                                        Mid$(json, currentBufferIndex, 1) = ChrW(val("&h" & Mid$(json, characterIndex + 1, 4)))
                                    #End If
                                    
                                    With Err
                                        If .Number <> 0 Then
                                            Mid$(json, currentBufferIndex, 1) = "?": .Clear
                                        End If
                                    End With
                                    On Error GoTo PropagateErr
                                    
                                    characterIndex = characterIndex + 4
                                Case Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                    End If
                    previousSpecialIsValid = True
                    
                Case vbLf, vbCr, vbTab, vbFormFeed
                
                    If previousSpecialIsValid And (characterIndex - regionStart > 0) Then
                        Call OverwriteBufferSegment(json, Mid$(json, regionStart, characterIndex - regionStart), currentBufferIndex, False)
                    End If
                    previousSpecialIsValid = False
                    regionStart = characterIndex + INDEX_OFFSET
                    
                Case Else
                    'For the last index
                    Call OverwriteBufferSegment(json, Mid$(json, regionStart, 1 + characterIndex - regionStart), currentBufferIndex, False)
            End Select
            regionStart = characterIndex + INDEX_OFFSET
        Loop
        ' Code has been using [json] as a buffer beginning at index 1.
        isRangeQuoteEnclosed = Left$(json, 1) = QuoteCharacter And Mid$(json, currentBufferIndex, 1) = QuoteCharacter
        
        If Not isRangeQuoteEnclosed Then
            If (Left$(json, 1) = LeftCurly And Mid$(json, currentBufferIndex, 1) = RightCurly) Then
                isArrayOrObject = True
            ElseIf (Left$(json, 1) = LeftSquare And Mid$(json, currentBufferIndex, 1) = RightSquare) Then
                isArrayOrObject = True
            End If
        End If
        ' If there are 2 unescaped quotes in the value region and they occur at the ends, then the segement is a string.
        If (unescapedQuoteCount = 2 And isRangeQuoteEnclosed) Then
            ProcessCharactersInRange = Mid$(json, 2, currentBufferIndex - 2)
        ElseIf unescapedQuoteCount > 0 And Not isArrayOrObject Then
            ' Ensure quality of string value. ie don't allow "apple to be excepted as a string.
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Invalid unescaped quotation mark(s) detected: " & Left$(json, currentBufferIndex)
        Else
            ProcessCharactersInRange = Left$(json, currentBufferIndex)
        End If
    Else
        isRangeQuoteEnclosed = False
        ProcessCharactersInRange = Mid$(json, startIndex, maximumAllowedIndex - startIndex + 1)
    End If

    Exit Function
PropagateErr:
    PropagateError Err, "ProcessCharactersInRange"
End Function
Private Function ConvertToBinaryIndex(n&) As Long
    ConvertToBinaryIndex = (n * BINARY_OFFSET) - 1
End Function
Private Function ConvertFromBinaryIndex(n&) As Long
    ConvertFromBinaryIndex = (n \ BINARY_OFFSET) + 1
End Function

Private Function TryGetWhitespaceCleanedIndexes(json As String, ByVal regionStartIndex&, ByVal regionEndIndex&, cleanedStartIndex&, cleanedEndIndex&) As Boolean
'========================================================================================================================
'   Summary:
'       Searches for non space,newlines,line feeds and carriage returns from both ends of [json] if they exist.
'       cleanedStartIndex and cleanedEndIndex will be set to the relevant non-white space indexes.
'   Returns
'       True if region between  [regionStartIndex] and [regionEndIndex] isn't just whitespace characters.
'       [cleanedEndIndex]   - First non-whitespace index in region.
'       [cleanedStartIndex] - Last non-whitespace in region.
'   Parameters
'       [json] - String to search
'       [regionStartIndex] - Index to start searching at
'       [regionEndIndex] - Maximum index to return.
'========================================================================================================================
    Dim cleaningLeft As Boolean, cleaningRight As Boolean, errorLeft As Boolean, errorRight As Boolean

    'Max index should be the non null portion of a string byte.
    cleaningLeft = True: cleaningRight = True

    Do
        If cleaningLeft Then
            On Error GoTo CatchLeft
            
            Select Case Mid$(json, regionStartIndex, 1)
                Case SpaceCharacter, vbTab, vbCr, vbLf
                    regionStartIndex = regionStartIndex + INDEX_OFFSET
                    cleaningLeft = Not regionStartIndex > regionEndIndex And regionStartIndex <= regionEndIndex
                Case Else
                    cleaningLeft = False
                    cleanedStartIndex = regionStartIndex
            End Select
            
        End If
Try_Right:
        If cleaningRight Then
            On Error GoTo CatchRight
            Select Case Mid$(json, regionEndIndex, 1)
                Case SpaceCharacter, vbTab, vbCr, vbLf
                    regionEndIndex = regionEndIndex - INDEX_OFFSET
                    cleaningRight = Not regionEndIndex < regionStartIndex And regionEndIndex >= regionStartIndex
                Case Else
                    cleaningRight = False
                    cleanedEndIndex = regionEndIndex
            End Select
        End If
NextLoop:
    Loop While cleaningLeft Or cleaningRight

    If (errorLeft And errorRight) Or (regionStartIndex > regionEndIndex) Then
        TryGetWhitespaceCleanedIndexes = False
    Else
        TryGetWhitespaceCleanedIndexes = True
    End If

    Exit Function
CatchLeft:
    cleaningLeft = False
    errorLeft = True
    Resume Try_Right
CatchRight:
    cleaningRight = False
    errorRight = True
    Resume NextLoop
End Function

Private Function SeekNextUnquotedBinaryCharacterIndex(ByRef json$, ByRef character$, ByRef startIndex&, ByRef maxSearchIndex&, indexLimitsAreBinary As Boolean) As Long
'======================================================================================================================================
' Summary
'       Searches for the next [character] not contained within quotes and returns its binary index
' Parameters:
'       [json] - String to search.
'       [character]     - Character or string to search for.
'       [startIndex]    - Index within [json] to start searching at.
'       [maxSearchIndex]- Index within [json] to stop searching at.
' Returns:
'       The binary Index not within unescaped quotes at which [character] appears. 0 if such a [character] doesn't exist in the region.
' Remarks:
'   [character] within escaped quotes will be ignored.
'======================================================================================================================================
    Dim characterIndexBinary&, leftQuoteIndexBinary&, loopCount&, rightQuoteIndexBinary&, _
    adjustLeftQuote As Long, isNotEscapedQuote As Boolean, quoteIndexBinary&, updateCharacterlocation As Boolean, maxIndexBinary&
    
    On Error GoTo Return_Default
    
    If indexLimitsAreBinary Then
        maxIndexBinary = maxSearchIndex
        rightQuoteIndexBinary = startIndex
    Else
        maxIndexBinary = ConvertToBinaryIndex(maxSearchIndex)
        rightQuoteIndexBinary = ConvertToBinaryIndex(startIndex)
    End If

    ' - BINARY_OFFSET so that start position will be maintained when calculating [characterIndexBinary] and following quotes.
    rightQuoteIndexBinary = rightQuoteIndexBinary - BINARY_OFFSET
    
    updateCharacterlocation = True
    
    Do
        If updateCharacterlocation Then
            characterIndexBinary = InStrB(rightQuoteIndexBinary + BINARY_OFFSET, json, character)
            If characterIndexBinary = 0 Then Exit Do
        End If
                
        ' Start searching from the location of right quote: wiil be incremented before searching.
        leftQuoteIndexBinary = rightQuoteIndexBinary
        ' Loop true and then False. Alter left quote if true else right quote.
        For adjustLeftQuote = -1 To 0
            quoteIndexBinary = leftQuoteIndexBinary
            Do
                quoteIndexBinary = InStrB(quoteIndexBinary + BINARY_OFFSET, json, QuoteCharacter)
                If Not quoteIndexBinary = 0 Then
                    ' Set to TRUE if previous character in the string isn't a backslash.
                    isNotEscapedQuote = Mid$(json, ConvertFromBinaryIndex(quoteIndexBinary) - INDEX_OFFSET, 1) <> BackslashCharacter
                End If
            Loop Until isNotEscapedQuote Or quoteIndexBinary = 0 Or (loopCount = 1 And adjustLeftQuote And characterIndexBinary < quoteIndexBinary And isNotEscapedQuote)
            
            If adjustLeftQuote Then
                leftQuoteIndexBinary = quoteIndexBinary
            Else
                rightQuoteIndexBinary = quoteIndexBinary
            End If
        Next adjustLeftQuote
        ' Update if sandwhiched between quotes.
        updateCharacterlocation = (characterIndexBinary > leftQuoteIndexBinary And characterIndexBinary < rightQuoteIndexBinary)
        ' Exit loop if (character isn't sandwhiched and comma less than left quote.) OR ccharacter not found OR Calculated index is invalid OR no more left quotes.
        loopCount = loopCount + 1
    Loop Until (characterIndexBinary < leftQuoteIndexBinary And Not updateCharacterlocation) Or characterIndexBinary = 0 Or (characterIndexBinary > maxIndexBinary) Or leftQuoteIndexBinary = 0 Or (loopCount = 1 And characterIndexBinary < leftQuoteIndexBinary)
    
    If characterIndexBinary > maxIndexBinary Or updateCharacterlocation Then
        SeekNextUnquotedBinaryCharacterIndex = 0
    ElseIf Not updateCharacterlocation Then
        SeekNextUnquotedBinaryCharacterIndex = characterIndexBinary
    End If
    
    Exit Function
Return_Default:
    SeekNextUnquotedBinaryCharacterIndex = 0
End Function

Private Sub PropagateError(e As ErrObject, ProcedureName$, Optional moreDetails$ = vbNullString)
    Dim firstPropagation As Boolean, sourceParts$()
    
    Const delim$ = ": "
    With e
        AddParentToErrSource e, ProcedureName
        If LenB(moreDetails) <> 0 Then .Description = moreDetails & vbNewLine & .Description
        .Raise .Number, .Source, .Description
    End With
    
End Sub

Private Sub AddParentToErrSource(e As ErrObject, parentName$)
    Dim sourceParts$()
    Const delim$ = ": "
    With e
        If InStrB(1, .Source, delim) = 0 Then
            .Source = LeftSquare & .Source & RightSquare & delim & parentName
        Else
            sourceParts = Split(.Source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .Source = Join(sourceParts, delim)
        End If
    End With
End Sub

Private Sub OverwriteBufferSegment(buffer$, textToAdd$, cursor&, Optional processFurther As Boolean = False)
'======================================================================================================================================
' Summary
'       Adds [textToAdd] to [buffer] starting at [cursor] + 1.
' Parameters:
'       [buffer] - Buffer that is added to.
'       [textToAdd]    - Text that is added to [buffer]
'       [cursor]    - Last index of previous addition.
'       [processFurther] - Set to True if you want to encode strings.
'======================================================================================================================================
    Dim additionLength&
        
    If LenB(textToAdd) <> 0 Then
        If processFurther Then
            If InStrB(1, textToAdd, BackslashCharacter) <> 0 Then textToAdd = Replace$(textToAdd, BackslashCharacter, "\\")
            If InStrB(1, textToAdd, vbTab) <> 0 Then textToAdd = Replace$(textToAdd, vbTab, "\t")
            If InStrB(1, textToAdd, QuoteCharacter) <> 0 Then textToAdd = Replace$(textToAdd, QuoteCharacter, "\""")
            If InStrB(1, textToAdd, vbCr) <> 0 Then textToAdd = Replace$(textToAdd, vbCr, "\r")
            If InStrB(1, textToAdd, vbBack) <> 0 Then textToAdd = Replace$(textToAdd, vbBack, "\b")
            If InStrB(1, textToAdd, vbLf) <> 0 Then textToAdd = Replace$(textToAdd, vbLf, "\n")
            If InStrB(1, textToAdd, vbFormFeed) <> 0 Then textToAdd = Replace$(textToAdd, vbFormFeed, "\f")
            'To Do: Convert Unicode characters.
    '        #If Mac Then
    '
    '        #Else
    '            Dim iCount&, ascwNumber&
    '            iCount = 1
    '            Do While iCount <= Len(textToAdd)
    '                ascwNumber = AscW(Mid$(textToAdd, iCount, 1))
    '                Select Case ascwNumber
    '                    Case Is > 127
    '                        textToAdd = Replace$(textToAdd, Mid$(textToAdd, iCount, 1), "\u" & Format$(Hex(ascwNumber), "0000"))
    '                        iCount = iCount + 6
    '                    Case Else
    '                        iCount = iCount + 1
    '                End Select
    '            Loop
    '        #End If
        End If
        
        additionLength = Len(textToAdd)
        If additionLength <> 0 Then
            ' Increase buffer until it can contain textToAdd.
            Do Until cursor + additionLength + 1 <= Len(buffer)
                buffer = buffer & String$(1000, " ")
            Loop
            
            Mid$(buffer, cursor + 1, additionLength) = textToAdd
            cursor = cursor + additionLength
        End If
        
    End If
    
End Sub
Private Function IsTextValidInteger(queriedText$, wantedLength&, ensureNumberLength As Boolean, ensureTextLength As Boolean, ByRef returnedNumber As Integer) As Boolean
    ' Trim text beforehand.
    If Not (InStrB(queriedText, "$") <> 0 Or InStrB(queriedText, "+") <> 0 Or InStrB(queriedText, "-") <> 0 Or InStrB(queriedText, ".") <> 0) Then
        If ensureTextLength And IsNumeric(queriedText) And Len(queriedText) = wantedLength Then
            returnedNumber = queriedText
            IsTextValidInteger = True
        ElseIf ensureNumberLength And IsNumeric(queriedText) Then
            returnedNumber = queriedText
            IsTextValidInteger = IsValueInRange(Abs(returnedNumber) / (10 ^ wantedLength), 0, 1, False) And (Abs(returnedNumber) / (10 ^ (wantedLength - 1))) >= 1
        End If
    End If
End Function
Private Function GetNewDictionary() As Object

    #If Mac Then
        ' https://github.com/VBA-tools/VBA-Dictionary
        Set GetNewDictionary = New Dictionary
    #Else
        Set GetNewDictionary = CreateObject("Scripting.Dictionary")
    #End If

End Function
Public Function TryParseExactIso8601DateUTC(iso$, convertToUTC As Boolean, convertToLocalTime As Boolean, returnedDate As Date, Optional ByVal dateFormat$) As Boolean
'======================================================================================================================================
' Summary
'       Attempts to parse [iso] as an ISO-8601 date.
' Parameters:
'       [iso] - Text to parse.
'       [returnDate] - Variable that stores the parsed date from [iso] if successful.
' Returns:
'       True if a date was successfuly parsed; otherwise, False.
'======================================================================================================================================
    Dim dateParts$(), isoDate As Date, isoTime As Date
    
    Dim dict As Object, ignoreDict As Boolean, isIsoWeekDate As Boolean, isUtc As Boolean, parsingTime As Boolean, dictKey$, _
    character$, cursor As Byte, tokenCount As Byte, followingTextIsTimeZoneOffset As Boolean, subSection$, formatLength As Byte, offsetMultiplier As Integer, _
    yearAvailalbe As Boolean, isOrdinal As Boolean

    Set dict = GetNewDictionary()
        
    formatLength = Len(dateFormat)

    If formatLength > IIf(InStrB(iso, "K") <> 0, Len(iso) + 6, Len(iso)) Then Exit Function
    
    On Error GoTo ReturnFalse
    
    For cursor = 1 To formatLength

        character = LCase$(Mid$(dateFormat, cursor, 1))
        tokenCount = 0

        Do While LCase$(Mid$(dateFormat, cursor, 1)) = character And cursor <= formatLength
            cursor = cursor + 1
            tokenCount = tokenCount + 1
        Loop

        subSection = Mid$(iso, (cursor - tokenCount), tokenCount)
        ignoreDict = tokenCount = 1 And InStrB(":-", subSection) <> 0

        dictKey = Mid$(dateFormat, (cursor - tokenCount), tokenCount)

        If Not dict.Exists(dictKey) Or ignoreDict Then

            dict.Item(dictKey) = Null

            Select Case dictKey
                Case "YYYY", "yyyy"
                    yearAvailalbe = True
                    returnedDate = DateSerial(subSection, 1, 1)
                Case "MM"
                    If yearAvailalbe Then
                        returnedDate = DateAdd("m", subSection - 1, returnedDate)
                    Else
                        Err.Raise 13, Description:="Invalid ISO-8601 format. Attempted to parse month but year is unavailable."
                    End If
                Case "hh"
                    parsingTime = True
                    returnedDate = DateAdd("h", subSection, returnedDate)
                Case "mm"
                    parsingTime = True
                    returnedDate = DateAdd("n", subSection, returnedDate)
                Case "ss"
                    parsingTime = True
                    returnedDate = DateAdd("s", subSection, returnedDate)
                Case "Www"
                    ' Week of year.
                    If yearAvailalbe Then
                        ' Find thursday >=
                        returnedDate = DateAdd("d", vbThursday - Weekday(returnedDate), returnedDate)
                        ' Find first Monday <
                        returnedDate = DateAdd("d", vbMonday - vbThursday, returnedDate)
                            
                        returnedDate = DateAdd("ww", Right$(subSection, 2) - 1, returnedDate)
                        isIsoWeekDate = True
                    Else
                        Err.Raise 13, Description:="Bad Format"
                    End If
                Case "DDD", "ddd", "DD", "dd", "D", "d"
                
                    isOrdinal = Len(subSection) = 3
                    
                    If yearAvailalbe Then
                        returnedDate = DateAdd("d", subSection - 1, returnedDate)
                    ElseIf isOrdinal Then
                        Err.Raise 13, Description:="Invalid ISO-8601 format. Year not found for ordinal date. Format as YYYY-DDD or YYYYDDD."
                    ElseIf isIsoWeekDate Then
                        Err.Raise 13, Description:="Bad Format"
                    End If
                    
                Case "+", "-"
                    If parsingTime Then
                        followingTextIsTimeZoneOffset = True
                    End If
                Case "K"
                    If convertToUTC Or convertToLocalTime Then
                    
                        subSection = Mid$(iso, cursor, Len(iso) - cursor + 1)
                        
                        Select Case Left$(subSection, 1)
                            Case "+"
                                offsetMultiplier = -1
                            Case "-"
                                offsetMultiplier = 1
                        End Select
                        
                        Select Case Len(subSection)
                            Case 1
                                If subSection = "Z" Then isUtc = True
                            Case 3 ' hh
                                returnedDate = returnedDate + (offsetMultiplier * TimeSerial(Right$(subSection, 2), 0, 0))
                            Case 5, 6 ' hh:mm or hhmm
                                returnedDate = returnedDate + (offsetMultiplier * TimeSerial(Left$(subSection, 2), Right$(subSection, 2), 0))
                            Case Else
                                Err.Raise 13, Description:="Bad Format"
                        End Select
                        
                        If convertToLocalTime Then
                            #If Mac Then
                                returnedDate = utc_ConvertDate(convertedDate, returnedDate)
                            #Else
                                Dim tz As TimeZoneInformation, localSystemTime As SystemTime
                                Call GetTimeZoneInformation(tz)
                                Call SystemTimeToTzSpecificLocalTime(tz, ConvertDateToSystemTime(returnedDate), localSystemTime)
                                returnedDate = ConvertSystemTimeToDate(localSystemTime)
                            #End If
                        End If
                        
                    End If
                Case "T"
                    parsingTime = True
                Case " ", ":"
                Case "."
                    Exit For
                Case Else
                    Err.Raise 13, Description:="Invalid character found in ISO-8601 format. > " & subSection
            End Select
        Else
            Err.Raise 13, Description:="Bad format. " & dictKey & " already parsed with submitted format."
        End If
        cursor = cursor - 1
    Next cursor
    
    TryParseExactIso8601DateUTC = True
    Exit Function
    
ReturnFalse:
    TryParseExactIso8601DateUTC = False
End Function
Public Function TryParseIsoTimestamp(timestamp$, convertToLocalTime As Boolean, convertToUTC As Boolean, output As Date) As Boolean

    Dim stampParts$(), parsedDateTime As SystemTime, segment$(), _
    millisecondsAvailable As Boolean, isUtc As Boolean, timezoneAvailable As Boolean, _
    milliSecondLocation As Byte, timeZoneStart As Byte, tempDate As Date, storedMilliseconds As Integer
    
    On Error GoTo ExitFunction
    
    stampParts = Split(timestamp, "T")

    If UBound(stampParts) = 1 Then

        segment = Split(stampParts(0), "-", 3)
        
        If UBound(segment) = 2 Then
            With parsedDateTime
            
                If Not IsTextValidInteger(segment(0), 4, False, True, .wYear) Then Exit Function
                If Not IsTextValidInteger(segment(1), 2, False, True, .wMonth) Then Exit Function
                If Not IsTextValidInteger(segment(2), 2, False, True, .wDay) Then Exit Function
                
                segment = Split(stampParts(1), ":", 3)
                
                If UBound(segment) = 2 Then
                    
                    If Not IsTextValidInteger(segment(0), 2, False, True, .wHour) Then Exit Function
                    If Not IsTextValidInteger(segment(1), 2, False, True, .wMinute) Then Exit Function
                    If Not IsTextValidInteger(Left$(segment(2), 2), 2, False, True, .wSecond) Then Exit Function
                    
                    isUtc = segment(2) Like "*Z" And InStr(segment(2), "Z") = Len(segment(2))
                                        
                    If isUtc Then
                        timeZoneStart = Len(segment(2))
                    Else
                        timeZoneStart = InStr(segment(2), "-")
                        If timeZoneStart = 0 Then timeZoneStart = InStr(segment(2), "+")
                    End If
                                        
                    timezoneAvailable = timeZoneStart <> 0
                    milliSecondLocation = InStr(segment(2), ".")
                    
                    If milliSecondLocation > 0 Then
                        If timezoneAvailable Then
                            .wMilliseconds = Left$(Mid$(segment(2), (milliSecondLocation + 1), timeZoneStart - (milliSecondLocation + 1)), 3)
                        Else
                            .wMilliseconds = Mid$(segment(2), (milliSecondLocation + 1), 3)
                        End If
                        storedMilliseconds = .wMilliseconds
                    End If
                    
                    If timezoneAvailable And Not isUtc And (convertToUTC Or convertToLocalTime) Then
                    
                        Dim offsetMultiplier As Integer, minutesCountOffset As Integer
                        
                        offsetMultiplier = IIf(Mid$(segment(2), timeZoneStart, 1) = "+", -1, 1)
                        minutesCountOffset = Mid$(segment(2), timeZoneStart + 1, 2) * 60
                        
                        tempDate = ConvertSystemTimeToDate(parsedDateTime) + (offsetMultiplier * TimeSerial(0, minutesCountOffset, 0))
                        
                        If timeZoneStart + 2 < Len(segment(2)) Then
                            minutesCountOffset = Right$(segment(2), 2)
                            tempDate = tempDate + (offsetMultiplier * TimeSerial(0, minutesCountOffset, 0))
                        End If
                                                
                        parsedDateTime = ConvertDateToSystemTime(tempDate)
                        .wMilliseconds = storedMilliseconds
                        
                    End If
                    
                    .wDayOfWeek = Weekday(DateSerial(.wYear, .wMonth, .wDay)) - 1
                    
                    If convertToLocalTime Then
                        #If Mac Then
                            output = utc_ConvertDate(convertedDate, False)
                        #Else
                        
                            Dim tz As TimeZoneInformation, localTime As SystemTime
                            
                            Call GetTimeZoneInformation(tz)
                            SystemTimeToTzSpecificLocalTime tz, parsedDateTime, localTime
                            output = ConvertSystemTimeToDate(localTime)
                            
                        #End If
                    Else
                        output = ConvertSystemTimeToDate(parsedDateTime)
                    End If
                    TryParseIsoTimestamp = True
                End If
                
            End With
        End If
    End If
ExitFunction:

End Function
Private Function ConvertSystemTimeToDate(timeToConvert As SystemTime) As Date
    With timeToConvert
        ConvertSystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
Private Function ConvertDateToSystemTime(timeToConvert As Date) As SystemTime
    With ConvertDateToSystemTime
        .wYear = Year(timeToConvert)
        .wMonth = Month(timeToConvert)
        .wDay = Day(timeToConvert)
        .wHour = Hour(timeToConvert)
        .wMinute = Minute(timeToConvert)
        .wSecond = Second(timeToConvert)
    End With
End Function
Private Function ConvertLocalDatetimeToUTC(convertedDate As Date) As Date
    #If Mac Then
        ConvertLocalDatetimeToUTC = utc_ConvertDate(convertedDate, True)
    #Else
        Dim tz As TimeZoneInformation, utcTime As SystemTime
        Call GetTimeZoneInformation(tz)
        Call TzSpecificLocalTimeToSystemTime(tz, ConvertDateToSystemTime(convertedDate), utcTime)
        ConvertLocalDatetimeToUTC = ConvertSystemTimeToDate(utcTime)
    #End If
End Function

#If Mac Then

    Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
        Dim utc_ShellCommand As String
        Dim utc_Result As utc_ShellResult
        Dim utc_Parts() As String
        Dim utc_DateParts() As String
        Dim utc_TimeParts() As String
    
        If utc_ConvertToUtc Then
            utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
                "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
                " +'%s'` +'%Y-%m-%d %H:%M:%S'"
        Else
            utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
                "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
                "+'%Y-%m-%d %H:%M:%S'"
        End If
    
        utc_Result = utc_ExecuteInShell(utc_ShellCommand)
    
        If utc_Result.utc_Output = "" Then
            Err.Raise 10015, "utc_ConvertDate", "'date' command failed"
        Else
            utc_Parts = Split(utc_Result.utc_Output, " ")
            utc_DateParts = Split(utc_Parts(0), "-")
            utc_TimeParts = Split(utc_Parts(1), ":")
    
            utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
                TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
        End If
    End Function
    
    Private Function utc_ExecuteInShell(utc_ShellCommand As String) As String
        
        #If VBA7 Then
            Dim utc_File As LongPtr
            Dim utc_Read As LongPtr
        #Else
            Dim utc_File As Long
            Dim utc_Read As Long
        #End If
    
        Dim utc_Chunk As String
    
        On Error GoTo utc_ErrorHandling
        utc_File = popen(utc_ShellCommand, "r")
    
        If utc_File = 0 Then: Exit Function
    
        Do While feof(utc_File) = 0
            utc_Chunk = VBA.Space$(50)
            utc_Read = CLng(fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
            If utc_Read > 0 Then
                utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
                utc_ExecuteInShell = utc_ExecuteInShell & utc_Chunk
            End If
        Loop
        
        Exit Function
utc_ErrorHandling:
        PropagateError Err, "utc_ExecuteShell"
    End Function
#End If
