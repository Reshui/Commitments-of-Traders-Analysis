VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimedTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If Not Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CurrentTimeAPI Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
        Private Declare PtrSafe Function CounterFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
    #Else
        Private Declare Function CurrentTimeAPI Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
        Private Declare Function CounterFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
    #End If
    Private pPerformanceFrequency As Currency
#End If

Private pStartTime As Currency
Private pFinishTime As Currency
Private pElapsedTime As Currency

Private pName$
Private pIsRunning As Boolean
Private pExecutionInterrupted As Boolean
Private pNestedDepth As Long

Private pNestedTimers As Collection
Private pLongestSubTaskLength As Long

Public Property Get ExecutionInterrupted() As Boolean
    ExecutionInterrupted = pExecutionInterrupted
End Property

Public Property Let ExecutionInterrupted(newExecutionValue As Boolean)
    pExecutionInterrupted = newExecutionValue
End Property
Public Property Let name(taskName$)
   pName = taskName
End Property
Public Property Get name$()
    name = pName
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = pIsRunning
End Property

'Private Property Let IsRunning(new_value As Boolean)
'    pIsRunning = new_value
'End Property

Public Property Get ElapsedTime()
    #If Not Mac Then
        ElapsedTime = ((GetCurrentTime - pStartTime) / pPerformanceFrequency) + pElapsedTime
    #Else
        ElapsedTime = (GetCurrentTime - pStartTime) + pElapsedTime
    #End If
End Property

Public Property Let DepthLevel(newDepth As Long)
    pNestedDepth = newDepth
End Property
Public Property Get DepthLevel() As Long
    DepthLevel = pNestedDepth
End Property

Public Property Let CacheFrequency(cc As Currency)
'================================================
'Summary: Sets a Performance Frequency value
'================================================
    pPerformanceFrequency = cc
End Property

Public Function Start(Optional runtimeNameTask$ = vbNullString) As TimedTask
    pIsRunning = True
    If LenB(runtimeNameTask) <> 0 Then pName = runtimeNameTask
    pStartTime = GetCurrentTime()
    Set Start = Me
End Function

Public Sub Pause()

    If pIsRunning Then
        
        #If Not Mac Then
            pFinishTime = (GetCurrentTime() - pStartTime) / pPerformanceFrequency
        #Else
            pFinishTime = GetCurrentTime() - pStartTime
        #End If

        pElapsedTime = pFinishTime + pElapsedTime
        pStartTime = 0
        pFinishTime = 0
        pIsRunning = False
    End If

End Sub
Public Sub Continue()
    Me.Start
End Sub
Public Sub EndTask()

    Dim item As TimedTask

    If pIsRunning Then
        ' Get duration since timer was last started
        #If Not Mac Then
            pFinishTime = (GetCurrentTime() - pStartTime) / pPerformanceFrequency
        #Else
            pFinishTime = GetCurrentTime() - pStartTime
        #End If
        
        If Not pNestedTimers Is Nothing Then
            For Each item In pNestedTimers
                With item
                    If .IsRunning Then
                        .EndTask
                        .ExecutionInterrupted = True
                    End If
                End With
            Next item
        End If

        pElapsedTime = pFinishTime + pElapsedTime
        pIsRunning = False
    End If

End Sub
Public Function ToString$(Optional longestParentTask As Long)

    Dim objectData$(), nestedTimer As TimedTask, i As Long, mainTaskDescription$

    If pIsRunning Then EndTask

    If pNestedDepth = 0 Then
        mainTaskDescription = vbNewLine & "Task: " & pName & " >>> " & Round(pElapsedTime, 3) & " seconds."
    Else
        mainTaskDescription = Space$(pNestedDepth * 4) & Chr(149) & " " & Left$(pName & Space$(longestParentTask), longestParentTask) & Space$(3) & IIf(pExecutionInterrupted = False, Round(pElapsedTime, 3) & "s", "Interrupted")
    End If

    If Not pNestedTimers Is Nothing Then

        ReDim Preserve objectData(pNestedTimers.Count)

        If pNestedDepth = 0 Then
            objectData(i) = mainTaskDescription & vbNewLine
        Else
            objectData(i) = mainTaskDescription
        End If

        For Each nestedTimer In pNestedTimers
            i = i + 1
            objectData(i) = nestedTimer.ToString(pLongestSubTaskLength)
        Next nestedTimer
        ToString = Join(objectData, vbNewLine)
    Else
        ToString = mainTaskDescription
    End If

End Function

Private Function CreateChildTimer(taskToAdd$) As TimedTask
'================================================
'Summary: Creates a nested SubTask Timer.
'================================================
    Dim newChildTimer As New TimedTask

    If pNestedTimers Is Nothing Then Set pNestedTimers = New Collection

    With newChildTimer
        .name = taskToAdd
        .CacheFrequency = pPerformanceFrequency
        .DepthLevel = pNestedDepth + 1
    End With

    pNestedTimers.Add newChildTimer, taskToAdd
    If Len(taskToAdd) > pLongestSubTaskLength Then pLongestSubTaskLength = Len(taskToAdd)

    Set CreateChildTimer = newChildTimer

End Function
Private Function Exists(taskName$) As Boolean
    On Error GoTo Exit_Exists
    Exists = IsObject(pNestedTimers(taskName))
Exit_Exists:
End Function
Public Function SubTask(taskName$) As TimedTask
'=========================================================================================
'Summary: Gets a SubTask Timer. If the specified timer doesn't exist then one is created.
'=========================================================================================
    On Error GoTo CreateTask
    Set SubTask = pNestedTimers(taskName)
    Exit Function
CreateTask:
    Set SubTask = CreateChildTimer(taskName)
End Function
Public Function StartSubTask(taskName$) As TimedTask
'================================================
'Summary: Starts and nests a subtask.
'================================================
    Set StartSubTask = SubTask(taskName): StartSubTask.Start

End Function
Public Sub StopSubTask(taskName$)
'================================================
'Summary: Stops a nested subtask.
'================================================
    If Exists(taskName) Then SubTask(taskName).EndTask
End Sub
Public Function ReturnReference() As TimedTask
    Set ReturnReference = Me
End Function
Private Function GetCurrentTime() As Currency

    #If Mac Then
        GetCurrentTime = Timer
    #Else
        Dim apiTime As Currency

        CurrentTimeAPI apiTime
        GetCurrentTime = apiTime

        If pPerformanceFrequency = 0 Then
            CounterFrequency apiTime
            pPerformanceFrequency = apiTime
        End If

    #End If

End Function
Public Sub DPrint()
    Debug.Print ToString
End Sub
