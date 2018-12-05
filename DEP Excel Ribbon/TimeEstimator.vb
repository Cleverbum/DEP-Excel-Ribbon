Imports System.Diagnostics

Public Class TimeEstimator

    Public TotalTasks As Long

    Private GuessedDuration As TimeSpan
    Private Timer As Stopwatch


    Sub New(tasks As Long, Guess As TimeSpan)
        TotalTasks = tasks
        GuessedDuration = Guess
        Timer = Stopwatch.StartNew()
    End Sub

    Sub New(tasks As Long, GuessedSeconds As Double)
        TotalTasks = tasks
        GuessedDuration = TimeSpan.FromSeconds(GuessedSeconds)
    End Sub

    Function TimeRemaining(TasksDone As Long) As TimeSpan
        Call Updateguess(TasksDone)
        TimeRemaining = TimeSpan.FromTicks((TotalTasks - TasksDone) * GuessedDuration.Ticks)
    End Function

    Private Sub Updateguess(tasksDone As Long)
        GuessedDuration = TimeSpan.FromTicks(Timer.ElapsedTicks / tasksDone)
    End Sub

    Sub Done()
        Timer.Stop()
    End Sub
End Class
