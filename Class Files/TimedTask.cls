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

Private pStartTime As Single
Private endTime As Single
Private pName$
Private pElapsedTime As Single
Private pIsRunning As Boolean
Private pExecutionInterrupted As Boolean
'Private pTabLevel As Byte

Private childTasks As Collection
Private longestChildName As Long

Public Property Get executionInterrupted() As Boolean
    executionInterrupted = pExecutionInterrupted
End Property

Public Property Let executionInterrupted(newExecutionValue As Boolean)
    pExecutionInterrupted = newExecutionValue
End Property

Public Property Let Name(taskName$)
   pName = taskName
End Property
Public Property Get Name$()
    Name = pName
End Property

Public Property Get isRunning() As Boolean
    isRunning = pIsRunning
End Property
Public Property Let isRunning(new_value As Boolean)
    pIsRunning = new_value
End Property

Public Property Get elapsedTime()
    elapsedTime = pElapsedTime
End Property

Public Sub Start(Optional runtimeNameTask$ = vbNullString)

    pIsRunning = True
    If LenB(runtimeNameTask) > 0 Then pName = runtimeNameTask
    pStartTime = Timer()
    
End Sub
Public Sub Pause()

    If pIsRunning Then
        endTime = Timer() - pStartTime
        pElapsedTime = endTime + pElapsedTime
        pStartTime = 0
        endTime = 0
        pIsRunning = False
    End If

End Sub
Public Sub Continue()
    Me.Start
End Sub
Public Sub EndTask()

    Dim Item As TimedTask
    
    If pIsRunning Then
        If pStartTime > 0 Then endTime = Timer() - pStartTime
        pElapsedTime = endTime + pElapsedTime
        pIsRunning = False
        
        If Not childTasks Is Nothing Then
            For Each Item In childTasks
                With Item
                    If .isRunning Then
                        .executionInterrupted = True
                        .EndTask
                    End If
                End With
            Next Item
        End If
    End If
    
End Sub
Public Function ToString$()
    
    Dim objectData$, ItemData$, Item As TimedTask
    
    If pIsRunning Then EndTask
    
    objectData = vbNewLine & "Task: " & pName & " >>> " & Round(pElapsedTime, 3) & " seconds."
    
    If Not childTasks Is Nothing Then
        
        objectData = objectData + vbNewLine
        
        For Each Item In childTasks

            With Item
                ItemData = Space(6) & "- Subtask: " & Left$(.Name & Space(longestChildName), longestChildName) & Space(3) & IIf(.executionInterrupted = False, Format(Round(.elapsedTime, 2), "@@@@!") & " seconds", "Interrupted")
            End With
            
            objectData = objectData & vbNewLine & ItemData
            
        Next Item
        
    End If
    
    ToString = objectData
    
End Function
Public Sub DPrint()
    Debug.Print ToString
End Sub
Private Function AddChildTask(taskToAdd$) As TimedTask
    
    Dim newChildTimer As New TimedTask
    
    If childTasks Is Nothing Then Set childTasks = New Collection
    
    With newChildTimer
        .Name = taskToAdd
        childTasks.Add newChildTimer, taskToAdd
        If Len(taskToAdd) > longestChildName Then longestChildName = Len(taskToAdd)
    End With
    
    Set AddChildTask = newChildTimer
    
End Function
Private Function Exists(taskName$) As Boolean
    On Error GoTo Exit_Exists
    Exists = IsObject(childTasks(taskName))
Exit_Exists:
End Function
Public Function SubTask(taskName$) As TimedTask
    
    On Error GoTo CreateTask
    Set SubTask = childTasks(taskName)
    Exit Function
CreateTask:
    Set SubTask = AddChildTask(taskName)
End Function
Public Function StartSubTask(taskName$) As TimedTask
    
    Dim newTask As New TimedTask
    Set newTask = SubTask(taskName)
    newTask.Start
    Set StartSubTask = newTask

End Function
Public Sub StopSubTask(taskName$)
    SubTask(taskName).EndTask
End Sub

