Public Class Form2
    Public timeestimate As TimeEstimator
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        For i As Integer = 1 To 100
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5))
            SetText(i)
            SetProgress(i)
        Next
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Label1
            .Text = "Wow, a label"
            .Visible = True

        End With
        timeestimate = New TimeEstimator(100, 30.0)
        With ProgressBar1
            .Minimum = 0
            .Maximum = 100
            .Visible = True

        End With
        BackgroundWorker1.RunWorkerAsync()

    End Sub

    Private Sub SetText(ByVal [text] As Integer)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf SetText)
            Me.Invoke(d, New Object() {[text]})
        Else
            Dim timeLeft As TimeSpan = timeestimate.TimeRemaining([text])
            Me.Label1.Text = "About " & PrettyString(timeLeft) & " remaining."

        End If
    End Sub
    Delegate Sub SetTextCallback(ByVal [text] As Integer)

    Private Sub SetProgress(ByVal [progress] As Double)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetProgressCallback(AddressOf SetProgress)
            Me.Invoke(d, New Object() {[progress]})
        Else
            Me.ProgressBar1.Value = [progress]
        End If
    End Sub
    Delegate Sub SetProgressCallback(ByVal [progress] As Double)
End Class