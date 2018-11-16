Imports System.Diagnostics
Imports System.IO
Imports OpenQA.Selenium

Public Class Form4
    Public interrupt As Boolean = False
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Downloading list of tickets"
        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call GetTickets()
    End Sub

    Sub GetTickets()
        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        Globals.ThisAddIn.Application.ActiveWindow.Activate()
        Dim Downloads As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads")
        Dim file As String, oldfile As String

        oldfile = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        wd.Navigate.GoToUrl("http://nextdesk/reports/ticketsClosedBySystem.php")


        wd.FindElementByName("system_selected").SendKeys("SAS")

        ' this doesn't work: wd.FindElementByName("start_date").SendKeys("01/01/2018")

        MsgBox("Please set an appropriate start date before continuing")

        wd.FindElementByName("report").Click()

        For i = 1 To 6
            Call SetText(Label1.Text & ".")
            Threading.Thread.Sleep(300)
            If interrupt Then Exit Sub
        Next
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            For i = 1 To 5
                Call SetText(Label1.Text & ".")
                Threading.Thread.Sleep(1000)
                If interrupt Then Exit Sub
            Next
        End If
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            MsgBox("File did not download correctly.", vbAbort)
            Exit Sub
        End If

        wd.Quit()

        Call SetText("File Downloaded.")

        Dim ignoredTickets As List(Of String) = Readfile(file)
        Dim ticketDetails As New List(Of Dictionary(Of String, String))
        Call SetText("Found " & ignoredTickets.Count & " ignored tickets")
        Dim j As Integer = 0

        wd = ndt.GiveMeChrome(False)

        Dim timeSoFar As Stopwatch
        timeSoFar = Stopwatch.StartNew()
        Dim timeTaken As Long, estimatedTotalTime As Integer
        For Each ticket In ignoredTickets
            If interrupt Then
                wd.Quit()
                Closeme()
                Exit Sub
            End If
            ndt.ticketNumber = ticket

            ticketDetails.Add(ndt.DEPScrape(wd))
            j += 1
            timeTaken = timeSoFar.Elapsed.TotalSeconds
            estimatedTotalTime = CInt(timeTaken / (j / ignoredTickets.Count))
            Call SetText("Read " & j & " of " & ignoredTickets.Count & " ignored tickets.")
            Call SetTextTwo("About " & estimatedTotalTime & "s remaining")
        Next

        wd.Quit()

        Closeme()


    End Sub
    Private Sub SetText(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf SetText)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label1.Text = [text]
        End If
    End Sub
    Delegate Sub SetTextCallback(ByVal [text] As String)


    Private Sub SetTextTwo(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetTextCallback(AddressOf SetText)
            Me.Invoke(d, New Object() {[text]})
        Else
            Me.Label2.Visible = True
            Me.Label2.Text = [text]
        End If
    End Sub



    Function Readfile(file As String) As List(Of String)
        Dim MyReader As New FileIO.TextFieldParser(file)
        Dim currentLine As String()

        MyReader.TextFieldType = FileIO.FieldType.Delimited
        MyReader.HasFieldsEnclosedInQuotes = True
        MyReader.SetDelimiters(",")
        Dim ignoredTickets As New List(Of String)


        While Not MyReader.EndOfData
            If interrupt Then
                Readfile = ignoredTickets
                Exit Function
            End If
            Try
                currentLine = MyReader.ReadFields
                If wasIgnored(currentLine) Then
                    ignoredTickets.Add(currentLine.First)

                End If


            Catch ex As FileIO.MalformedLineException
                Debug.WriteLine("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
        End While
        MyReader.Close()
        Return ignoredTickets

    End Function

    Function WasIgnored(ticketData As String()) As Boolean
        If ticketData.Last.ToLower.Contains("no dep tickets will be raised") Then
            Return True
        ElseIf ticketData.Last.ToLower.Contains(Form3.CloseMessage.ToLower) Then
            Return True
        Else
            Return False
        End If
    End Function


    Private Sub Closeme()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New CloseCallBack(AddressOf Closeme)
            Me.Invoke(d, New Object() {})
        Else
            Me.Close()
        End If
    End Sub
    Delegate Sub CloseCallBack()


End Class