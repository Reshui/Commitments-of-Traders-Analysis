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
Option Explicit
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const INDEX_OFFSET As Byte = 1
Private Const RightCurly$ = "}", LeftCurly$ = "{", LeftSquare$ = "[", RightSquare$ = "]", DecimalChar$ = ".", _
QuoteCharacter$ = """", BackslashCharacter$ = "\", SpaceCharacter$ = " ", ColonCharacter$ = ":", Comma$ = ",", BINARY_OFFSET As Byte = 2

Private SpecialCharactersA$(6)
#Const DebugActive = False

Public Function ParseJsonString(ByRef json$) As Object
    Set ParseJsonString = ParseJSON(json)
End Function
Public Function ParseJsonBytes(ByRef jsonBytes() As Byte) As Object
    'PreProcessData jsonBytes
    Set ParseJsonBytes = ParseJSON(CStr(jsonBytes))
End Function

Private Function ParseJSON(ByRef json$) As Object
''====================================================================================================================================================================================
'' Summary: Non-Recursive JSON parser.
'' Parameters:
''      [json] - JSON string to parse.
'' Reference:
''   https://www.json.org/json-en.html
''====================================================================================================================================================================================
    Dim value$, isValueJsonObject As Boolean, isValueJsonArray As Boolean, isContainerJsonObject As Boolean, isValueAvailable As Boolean

    Dim jsonLength&, keyForValue$, valueStartIndex&, valueEndIndex&, jsonContainer As Object, _
    isValidatedString As Boolean, heirarchyStack As New Collection, cursorLocation&, nextCursorLocation&, isKeyAvailable As Boolean
    
    Const cursorLimit$ = "MoshiM_UC@outlook.com"
    Const cursorLocationAfterCompletion$ = "https://github.com/Reshui?tab=repositories"
    Const arrayPattern$ = "[*]", objectPattern$ = "{*}"
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
            
        SpecialCharactersA(0) = QuoteCharacter: SpecialCharactersA(1) = BackslashCharacter
        SpecialCharactersA(2) = vbCr: SpecialCharactersA(3) = vbTab
        SpecialCharactersA(4) = vbLf: SpecialCharactersA(5) = VBA.vbBack: SpecialCharactersA(6) = VBA.vbFormFeed

        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
        Set jsonContainer = heirarchyStack(1)
        
        AddValueToObject jsonContainer, isValueJsonObject, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=cursorLimit
        AddValueToObject jsonContainer, isValueJsonObject, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=cursorLocationAfterCompletion
        
        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
        
        Do While cursorLocation < jsonContainer.Item(cursorLimit)
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
                        
                        isValueJsonObject = False: isValueJsonArray = False
                        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
                        ' Create heirarchy by nesting child within parent.
                        With heirarchyStack
                            AddValueToObject jsonContainer, isContainerJsonObject, newChild:=.Item(.count), keyForValue:=keyForValue
                            Set jsonContainer = .Item(.count)
                        End With

                        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                        ' Add key to limit the cursor range for object.
                        AddValueToObject jsonContainer, isContainerJsonObject, addingNumber:=True, numberToAdd:=valueEndIndex, keyForValue:=cursorLimit
                        ' Index to update cursorLocation to once array or object has been completed.
                        AddValueToObject jsonContainer, isContainerJsonObject, addingNumber:=True, numberToAdd:=nextCursorLocation, keyForValue:=cursorLocationAfterCompletion
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
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="JSON parser failed. Current cursorLocation >= next cursorLocation."
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
        Set ParseJSON = jsonContainer
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Bytes are wrong."
    End If
    
    Exit Function
CATCH_FAILURE:
    #If DebugActive Then
        Stop: Resume
    #Else
        PropagateError Err, "ParseJSON"
    #End If
End Function

Private Sub CreateObjectOnStack(stack As Collection, createJsonObject As Boolean)

    If createJsonObject Then
        #If Mac Then
            ' https://github.com/VBA-tools/VBA-Dictionary
            stack.Add New Dictionary
        #Else
            stack.Add CreateObject("Scripting.Dictionary")
        #End If
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
    ' Convert to binary location.
    currentIndexBinary = (cursorLocation * BINARY_OFFSET) - 1
    maxIndexBinary = (maxCursorLocation * BINARY_OFFSET) - 1
    ' Start searching from at the index.
    nextCurlyOpenBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, LeftCurly)
    nextCurlyCloseBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, RightCurly)
    nextSquareOpenBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, LeftSquare)
    nextSquareCloseBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, RightSquare)
    nextValidCommaTokenIndexBinary = BinaryCharacterIndexNotWithinQuotes(json, Comma, cursorLocation + INDEX_OFFSET, maxCursorLocation)

    If nextValidCommaTokenIndexBinary > 0 Then nextEnclosingTokenIndex = (nextValidCommaTokenIndexBinary \ BINARY_OFFSET) + 1

    If nextValidCommaTokenIndexBinary > 0 Then
        '======================================================================================================================================
        ' Determine which of [ or { is greater than cursorLocation AND appears first.
        ' If nextValidCommaTokenIndexBinary is greater than that byte,
        ' then determine if the Comma is encapsulated. If it is, then find the next Comma that isn't encapsulated to find the end of the value.
        '======================================================================================================================================
        If IsValueInRange(nextCurlyOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary - BINARY_OFFSET) And (nextCurlyOpenBinary < nextSquareOpenBinary Or nextSquareOpenBinary = 0) Then
            tokenIndicator(0) = LeftCurly: tokenIndicator(1) = RightCurly
            nextOpenBinary = nextCurlyOpenBinary: nextCloseBinary = nextCurlyCloseBinary
            escapeEnclosingArrayOrObject = True
        ElseIf IsValueInRange(nextSquareOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary - BINARY_OFFSET) And (nextSquareOpenBinary < nextCurlyOpenBinary Or nextCurlyOpenBinary = 0) Then
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
                nextCloseBinary = InStrB(previousCloseBinary + BINARY_OFFSET, json, tokenIndicator(1))
                nextOpenBinary = InStrB(nextOpenBinary + BINARY_OFFSET, json, tokenIndicator(0))
            Loop Until nextOpenBinary > previousCloseBinary Or nextOpenBinary = 0 Or (previousCloseBinary >= maxIndexBinary)

            If previousCloseBinary < maxIndexBinary Then
                ' Find MIN not equal to 0 of the following.
                nextEnclosingTokenIndex = BinaryCharacterIndexNotWithinQuotes(json, Comma, (previousCloseBinary \ BINARY_OFFSET) + 1, maxCursorLocation)
                minValues(0) = InStrB(previousCloseBinary + BINARY_OFFSET, json, RightCurly)
                minValues(1) = InStrB(previousCloseBinary + BINARY_OFFSET, json, RightSquare)
                
                For nextCloseBinary = 0 To 1
                    If IsValueInRange(minValues(nextCloseBinary), 1, nextEnclosingTokenIndex) Or nextEnclosingTokenIndex <= 0 Then nextEnclosingTokenIndex = minValues(nextCloseBinary)
                Next nextCloseBinary
                If nextEnclosingTokenIndex > 0 Then nextEnclosingTokenIndex = (nextEnclosingTokenIndex \ BINARY_OFFSET) + 1
            Else
                nextEnclosingTokenIndex = maxCursorLocation
            End If
        End If
    End If

    Select Case nextEnclosingTokenIndex
        Case Is <= 0, cursorLocation, Is > maxCursorLocation
            nextEnclosingTokenIndex = maxCursorLocation
    End Select
    
    If InStrB(1, ",]}", Mid$(json, nextEnclosingTokenIndex, 1)) = 0 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid JSON format. Calculated token index doesn't reference a value delimiter: " & Mid$(json, nextEnclosingTokenIndex, 1)
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

Private Function IsValueInRange(numberToCheck&, lowerBound&, upperBound&) As Boolean
'Summary Checks if numberToCheck is contained within a given range.
    IsValueInRange = numberToCheck >= lowerBound And numberToCheck <= upperBound
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

            keyValueDelimiterIndex = (BinaryCharacterIndexNotWithinQuotes(json, ColonCharacter, keyStartLocation, valueEndLocation) \ BINARY_OFFSET) + 1
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

Private Sub AddValueToObject(containerObject As Object, isContainerJsonObject As Boolean, Optional value$, Optional isValueString As Boolean, Optional newChild As Object, Optional keyForValue$, Optional addingNumber As Boolean, Optional numberToAdd As Long)

    On Error GoTo PropagateError
    
    If newChild Is Nothing Then
        If addingNumber Then
            ' This is just used for tracking byte locations.
            If isContainerJsonObject Then
                containerObject.Item(keyForValue) = numberToAdd
            Else
                containerObject.Add numberToAdd, keyForValue
            End If
        ElseIf isValueString Then
            If isContainerJsonObject Then
                containerObject.Item(keyForValue) = value
            Else
                containerObject.Add value
            End If
        ElseIf IsNumeric(value) Then
            If isContainerJsonObject Then
                containerObject.Item(keyForValue) = IIf(InStrB(1, value, DecimalChar) > 0, CDbl(value), CLng(value))
            Else
                containerObject.Add IIf(InStrB(1, value, DecimalChar) > 0, CDbl(value), CLng(value))
            End If
        Else
            Select Case value
                Case "true", "false"
                    If isContainerJsonObject Then
                        containerObject.Item(keyForValue) = CBool(value)
                    Else
                        containerObject.Add CBool(value)
                    End If
                Case "null", vbNullString
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
            containerObject.Item(keyForValue) = newChild
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
    Dim iLoop&, currentBufferIndex&, isArrayOrObject As Boolean, currentCharacter$, minValues&(6), _
    groupedBackslashCount&, countOfBackslashToKeep&, specialCharactersInRange As Boolean, unescapedQuoteCount As Long

    ' Determine if there are any special characters in the range.
    For iLoop = LBound(SpecialCharactersA) To UBound(SpecialCharactersA)
        ' If within bounds set specialCharactersInRange equal to true and exit the loop.
        If IsValueInRange(InStrB((startIndex * BINARY_OFFSET) - 1, json, SpecialCharactersA(iLoop)), 1, (maximumAllowedIndex * BINARY_OFFSET) - 1) Then
            specialCharactersInRange = True: Exit For
        End If
    Next iLoop

    On Error GoTo PropagateErr
    
    If specialCharactersInRange Then

        Dim characterIndex&, regionStart&, writeClearedRegionToBuffer As Boolean, nonInclusiveRegionEnd&, newBufferIndex&, newRegionStart&

        regionStart = startIndex: characterIndex = startIndex - INDEX_OFFSET: currentBufferIndex = 0
                
        Do While characterIndex < maximumAllowedIndex
            ' Convert to binary.
            characterIndex = (characterIndex * BINARY_OFFSET) - 1 + BINARY_OFFSET
            ' Determine min special character index > 0.
            For iLoop = LBound(SpecialCharactersA) To UBound(SpecialCharactersA)
                'Use InstrB to determine the forward index of any special characters if characterIndex exceeds a previously stored value.
                If minValues(iLoop) < characterIndex Then minValues(iLoop) = InStrB(characterIndex, json, SpecialCharactersA(iLoop))
                If IsValueInRange(minValues(iLoop), 1, characterIndex - BINARY_OFFSET) Or iLoop = LBound(SpecialCharactersA) Then characterIndex = minValues(iLoop)
            Next iLoop
            ' Convert from binary with integer division.
            If characterIndex > 0 Then characterIndex = (characterIndex \ BINARY_OFFSET) + 1
            If characterIndex > maximumAllowedIndex Or characterIndex = 0 Then characterIndex = maximumAllowedIndex

            Select Case Mid$(json, characterIndex, 1)

                Case QuoteCharacter
                    unescapedQuoteCount = unescapedQuoteCount + 1: writeClearedRegionToBuffer = True: nonInclusiveRegionEnd = characterIndex + 1
                    newRegionStart = characterIndex + INDEX_OFFSET
                Case BackslashCharacter
                    writeClearedRegionToBuffer = True: nonInclusiveRegionEnd = characterIndex: newBufferIndex = currentBufferIndex: groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        characterIndex = characterIndex + INDEX_OFFSET
                    Loop Until Mid$(json, characterIndex, 1) <> BackslashCharacter Or characterIndex > maximumAllowedIndex
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    
                    If countOfBackslashToKeep > 0 Then
                        ' Write backslashes to buffer.
                        For newBufferIndex = newBufferIndex + INDEX_OFFSET To newBufferIndex + countOfBackslashToKeep Step INDEX_OFFSET
                            Mid$(json, newBufferIndex, 1) = BackslashCharacter
                        Next newBufferIndex
                        ' For loops will exceed boundary on last loop so reverse by the value of step.
                        newBufferIndex = newBufferIndex - INDEX_OFFSET
                    End If
                    
                    If characterIndex <= maximumAllowedIndex Then
                        ' If there is an even number of backslashes then all of them are escaped and the following character doesn't need additional processing.
                        If (groupedBackslashCount Mod 2) <> 0 Then
                            currentCharacter = Mid$(json, characterIndex, 1): newBufferIndex = newBufferIndex + INDEX_OFFSET
                            Select Case currentCharacter
                                Case "b" '\b
                                    Mid$(json, newBufferIndex, 1) = VBA.vbBack  'backspace
                                Case "n" '\n
                                    Mid$(json, newBufferIndex, 1) = vbNewLine 'newline
                                Case "t" '\t
                                    Mid$(json, newBufferIndex, 1) = vbTab  'tab
                                Case "0" 'Null character.
                                    newBufferIndex = newBufferIndex - INDEX_OFFSET
                                Case QuoteCharacter, "'", "/" '\" or \/ or \0 or \'
                                    Mid$(json, newBufferIndex, 1) = currentCharacter
                                Case "r" '\r
                                    Mid$(json, newBufferIndex, 1) = vbCr 'carriage return
                                Case "f" '\f
                                    Mid$(json, newBufferIndex, 1) = vbFormFeed ' Form feed
                                Case "u" '\u
                                    ' See: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chr-function
                                    On Error Resume Next
                                    #If Mac Then
                                        Mid$(json, newBufferIndex, 1) = WorksheetFunction.Unichar(val("&h" & Mid$(json, characterIndex + 1, 4)))
                                    #Else
                                        Mid$(json, newBufferIndex, 1) = ChrW(val("&h" & Mid$(json, characterIndex + 1, 4)))
                                    #End If
                                    
                                    With Err
                                        If .Number <> 0 Then
                                            Mid$(json, newBufferIndex, 1) = "?": .Clear
                                        End If
                                    End With
                                    On Error GoTo PropagateErr
                                    
                                    characterIndex = characterIndex + 4
                                Case Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                    End If
                    newRegionStart = characterIndex + INDEX_OFFSET
                Case vbLf, vbCr, vbTab, vbFormFeed
                    writeClearedRegionToBuffer = True: nonInclusiveRegionEnd = characterIndex: newRegionStart = characterIndex + INDEX_OFFSET
                Case Else
                    'For the last index
                    writeClearedRegionToBuffer = True: nonInclusiveRegionEnd = characterIndex + 1
            End Select
            
            If writeClearedRegionToBuffer Or characterIndex = maximumAllowedIndex Then
                If Not writeClearedRegionToBuffer Then nonInclusiveRegionEnd = characterIndex + INDEX_OFFSET
                currentBufferIndex = currentBufferIndex + 1
                ' [json] is used as both source and buffer. [currentBufferIndex] initializes at 1.
                Mid$(json, currentBufferIndex, nonInclusiveRegionEnd - regionStart) = Mid$(json, regionStart, nonInclusiveRegionEnd - regionStart)
                ' Set to index of last addition.
                currentBufferIndex = currentBufferIndex + (nonInclusiveRegionEnd - regionStart) - 1
                ' If backslash case was used.
                If newBufferIndex > 0 Then currentBufferIndex = newBufferIndex: newBufferIndex = 0
                regionStart = newRegionStart: writeClearedRegionToBuffer = False
            End If
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

Private Function BinaryCharacterIndexNotWithinQuotes(ByRef json$, ByRef character$, ByRef startIndex&, ByRef maxSearchIndex&) As Long
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
    
    maxIndexBinary = (maxSearchIndex * BINARY_OFFSET) - 1
    ' - BINARY_OFFSET so that start position will be maintained when calculating [characterIndexBinary] and following quotes.
    rightQuoteIndexBinary = ((startIndex * BINARY_OFFSET) - 1) - BINARY_OFFSET
    
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
                    isNotEscapedQuote = Mid$(json, ((quoteIndexBinary \ BINARY_OFFSET) + 1) - INDEX_OFFSET, 1) <> BackslashCharacter
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
        BinaryCharacterIndexNotWithinQuotes = 0
    ElseIf Not updateCharacterlocation Then
        BinaryCharacterIndexNotWithinQuotes = characterIndexBinary
    End If
    
    Exit Function
Return_Default:
    BinaryCharacterIndexNotWithinQuotes = 0
End Function

Private Sub PropagateError(e As ErrObject, procedureName$, Optional moreDetails$ = vbNullString)
    Dim firstPropagation As Boolean, sourceParts$()
    
    Const delim$ = ": "
    With e
        AddParentToErrSource e, procedureName
        If LenB(moreDetails) > 0 Then .Description = moreDetails & vbNewLine & .Description
        .Raise .Number, .source, .Description
    End With
    
End Sub

Private Sub AddParentToErrSource(e As ErrObject, parentName$)
    Dim sourceParts$()
    Const delim$ = ": "
    With e
        If InStrB(1, .source, delim) = 0 Then
            .source = LeftSquare & .source & RightSquare & delim & parentName
        Else
            sourceParts = Split(.source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .source = Join(sourceParts, delim)
        End If
    End With
End Sub

Public Sub Unit_Test()
    Dim success As Boolean, json$
    With Me
        '.ParseJsonString "{""apple"":""tree"",""inner object"":1,""array"":[1,2,3,{""apple"":22,""berry"":45}]}"
        .ParseJsonString "{""ground\n\\\""beaver"":""tr\\ee"",""inner object"":[{""APP"":1,""Sun"":""Moon"",""Ocean"":""Beach""},[""a"",""b"",""c""],[0,1,2,[5,6,7,8],9,10]],""array"":[1,2,3,{""apple"":22,""berry"":45}]}"
        '.ParseJsonString HttpGet("http://worldtimeapi.org/api/timezone/America/New_York", success)
    End With
End Sub
Public Function ConvertToJSON(container As Object) As String

    Dim output$, cursorLocation&, isContainerJsonObject As Boolean, keyOrIndex$(), iCount&, isValueString As Boolean

    Const quote$ = """"
    output = String(1000, " ")
        
    If Not TypeOf container Is Collection Then
        isContainerJsonObject = True
        keyOrIndex = container.keys
    Else
        ReDim keyOrIndex(1 To container.count)
        For iCount = 1 To UBound(keyOrIndex)
            keyOrIndex(iCount) = iCount
        Next iCount
    End If

    If isContainerJsonObject Then
        AddToString output, "{", cursorLocation
    Else
        AddToString output, "[", cursorLocation
    End If
    
    With container

        For iCount = LBound(keyOrIndex) To UBound(keyOrIndex)

            If isContainerJsonObject Then
                AddToString output, quote, cursorLocation
                AddToString output, keyOrIndex(iCount), cursorLocation, True
                AddToString output, quote, cursorLocation
                AddToString output, ":", cursorLocation
            End If
                
            If IsObject(.Item(keyOrIndex(iCount))) Then
                AddToString output, ConvertToJSON(.Item(keyOrIndex(iCount))), cursorLocation
            Else
                If VarType(.Item(keyOrIndex(iCount))) = vbString Then
                    isValueString = True
                Else
                    isValueString = False
                End If
                
                If isValueString Then AddToString output, quote, cursorLocation
                AddToString output, .Item(keyOrIndex(iCount)), cursorLocation, isValueString
                If isValueString Then AddToString output, quote, cursorLocation
                
            End If

            If iCount <> UBound(keyOrIndex) Then
                AddToString output, ",", cursorLocation
            End If

        Next iCount

    End With
    
    If isContainerJsonObject Then
        AddToString output, "}", cursorLocation
    Else
        AddToString output, "]", cursorLocation
    End If
    ConvertToJSON = Left$(output, cursorLocation)
End Function

Private Sub AddToString(json$, textToAdd$, cursor&, Optional processFurther As Boolean = False)

    Dim additionLength&
    Const quote$ = """"
    
    If processFurther Then
        If InStrB(1, textToAdd, vbTab) > 0 Then textToAdd = Replace$(textToAdd, vbTab, "\t")
        If InStrB(1, textToAdd, quote) > 0 Then textToAdd = Replace$(textToAdd, quote, "\""")
        If InStrB(1, textToAdd, vbCr) > 0 Then textToAdd = Replace$(textToAdd, vbCr, "\r")
        If InStrB(1, textToAdd, vbBack) > 0 Then textToAdd = Replace$(textToAdd, vbBack, "\b")
        If InStrB(1, textToAdd, vbLf) > 0 Then textToAdd = Replace$(textToAdd, vbLf, "\n")
        If InStrB(1, textToAdd, vbFormFeed) > 0 Then textToAdd = Replace$(textToAdd, vbFormFeed, "\f")
        If InStrB(1, textToAdd, "\") > 0 Then textToAdd = Replace$(textToAdd, "\", "\\")
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
    ' Increase buffer until it can contain textToAdd.
    Do Until cursor + additionLength + 1 <= Len(json)
        json = json & String(1000, " ")
    Loop
    
    Mid$(json, cursor + 1, additionLength) = textToAdd
    cursor = cursor + additionLength

End Sub
