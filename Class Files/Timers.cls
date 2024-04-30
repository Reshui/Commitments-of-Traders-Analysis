VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Tasks As Collection
Private mDescription As String

Private Sub Class_Initialize()
    Set Tasks = New Collection
End Sub

Private Sub Class_Terminate()
    Set Tasks = Nothing
End Sub

Public Property Let description(new_description As String)
    mDescription = new_description
End Property
Public Property Get description() As String
    description = mDescription
End Property
'Public Sub Add(taskName As String, Optional taskExistance As Boolean = False)
'
'    Dim newTask As TimedTask
'
'    If Not Exists(taskName) Then
'        Set newTask = New TimedTask
'
'        newTask.name = taskName
'
'        Tasks.Add newTask, taskName
'    End If
'
'End Sub
Public Function ReturnTimedTask(taskName As String) As TimedTask
    
    Dim newTask As TimedTask
    
    If Not Exists(taskName) Then
    
        Set ReturnTimedTask = New TimedTask
        ReturnTimedTask.name = taskName
                
        Tasks.Add ReturnTimedTask, taskName
        
    Else
        Set ReturnTimedTask = Tasks(taskName)
    End If
    
End Function
Public Sub StartTask(taskName As String)
    
    ReturnTimedTask(taskName).Start

End Sub
Public Sub EndTask(taskName As String)

    Tasks(taskName).EndTask

End Sub

Public Sub Pause(taskName As String)
    Tasks(taskName).Pause
End Sub

Public Sub Continue(taskName As String)
    
    Tasks(taskName).Continue

End Sub
Private Function Exists(taskName As String) As Boolean

    Dim TaskExists As TimedTask
    
    On Error Resume Next
    Set TaskExists = Tasks(taskName)
    If Err.Number <> 0 Then Err.Clear
    Exists = Not TaskExists Is Nothing
    
End Function

Public Function ToString() As String

    Dim AllTasks() As String, Z As Byte
    
    Dim header As String: header = String(80, "-")
    
    With Tasks
        If .count > 0 Then
            ReDim AllTasks(1 To .count)
            For Z = 1 To .count
                AllTasks(Z) = .Item(Z).ToString
            Next Z
            
            ToString = header & vbNewLine & "[" & Me.description & "]" & vbNewLine & vbNewLine & Join(AllTasks, vbNewLine) & vbNewLine & header
        
        Else
            ToString = vbNullString
        End If
    End With
    
End Function

