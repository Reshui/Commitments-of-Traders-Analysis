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

Private startTime As Double
Private endTime As Double
Private pName As String
Private pElapsedTime As Double
Private pIsRunning As Boolean
Private pExecutionInterrupted As Boolean

Private childTasks As Collection
Private longestChildName As Integer

Public Property Get executionInterrupted() As Boolean
    executionInterrupted = pExecutionInterrupted
End Property

Public Property Let executionInterrupted(newExecutionValue As Boolean)
    pExecutionInterrupted = newExecutionValue
End Property

Public Property Let name(taskName As String)
   pName = taskName
End Property
Public Property Get name() As String
    name = pName
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

Public Sub Start(Optional runtimeNameTask As String = vbNullString)

    pIsRunning = True
    If LenB(runtimeNameTask) > 0 Then pName = runtimeNameTask
    startTime = Timer()

End Sub
Public Sub Pause()

    If pIsRunning Then
        endTime = Timer() - startTime
        pElapsedTime = endTime + pElapsedTime
        startTime = 0
        endTime = 0
        pIsRunning = False
    End If

End Sub
Public Sub Continue()
    Me.Start
End Sub
Public Sub EndTask()

    Dim Item As Variant
    
    If pIsRunning Then
        If startTime > 0 Then endTime = Timer() - startTime
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
Public Function ToString() As String
    
    Dim objectData As String, ItemData As String, Item As Variant
    
    If pIsRunning Then EndTask
    
    objectData = "Task: " & pName & " >>> " & Round(pElapsedTime, 3) & " seconds."
    
    If Not childTasks Is Nothing Then
        
        objectData = objectData + vbNewLine
        
        For Each Item In childTasks

            With Item
                ItemData = Space(6) & "- Subtask: " & Left$(.name & Space(longestChildName), longestChildName) & Space(3) & IIf(.executionInterrupted = False, format(Round(.elapsedTime, 2), "@@@@!") & " seconds", "Interrupted")
            End With
            
            objectData = objectData & vbNewLine & ItemData
            
        Next Item
        
    End If
    
    ToString = objectData
    
End Function
Public Sub DPrint()
    Debug.Print ToString
End Sub
Private Function AddChildTask(taskToAdd As String) As TimedTask
    
    Dim newChildTimer As New TimedTask
    
    If childTasks Is Nothing Then Set childTasks = New Collection
    
    With newChildTimer
    
        .name = taskToAdd
        
        childTasks.Add newChildTimer, taskToAdd
        
        If Len(taskToAdd) > longestChildName Then longestChildName = Len(taskToAdd)
        
    End With
    
    Set AddChildTask = newChildTimer
    
End Function
Private Function Exists(taskName As String) As Boolean
    On Error Resume Next
    Dim TT As TimedTask
    
    Set TT = childTasks(taskName)
    
    If Not TT Is Nothing Then Exists = True
    
End Function
Public Function SubTask(taskName As String) As TimedTask
    
    If Exists(taskName) Then
        Set SubTask = childTasks(taskName)
    Else
        Set SubTask = AddChildTask(taskName)
    End If
    
End Function

