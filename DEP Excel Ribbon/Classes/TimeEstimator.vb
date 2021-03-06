﻿Imports System.Diagnostics

Public Class TimeEstimator

    Public TotalTasks As Long

    Private GuessedDuration As TimeSpan
    Private Timer As Stopwatch


    Sub New(tasks As Long, Guess As TimeSpan)
        Call Setup(tasks, Guess)
    End Sub

    Sub New(tasks As Long, GuessedSeconds As Double)
        Call Setup(tasks, TimeSpan.FromSeconds(GuessedSeconds))
    End Sub

    Sub Setup(tasks As Long, Guess As TimeSpan)
        TotalTasks = tasks
        GuessedDuration = Guess
        Timer = Stopwatch.StartNew()
    End Sub

    Function TimeRemaining(TasksDone As Long) As TimeSpan
        'if we've done something, measure how long it took
        If TasksDone > 0 Then Call Updateguess(TasksDone)

        TimeRemaining = TimeSpan.FromTicks((TotalTasks - TasksDone) * GuessedDuration.Ticks)
    End Function

    Private Sub Updateguess(tasksDone As Long)
        GuessedDuration = TimeSpan.FromTicks(Timer.ElapsedTicks / tasksDone)
    End Sub

    Sub Done()
        Timer.Stop()
    End Sub
End Class
