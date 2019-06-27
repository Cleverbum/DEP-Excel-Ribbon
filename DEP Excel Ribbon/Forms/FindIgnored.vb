Imports System.Diagnostics
Imports System.IO
Imports OpenQA.Selenium

Public Class FindIgnored
    Public interrupt As Boolean = False
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        interrupt = False
        Label1.Text = "Downloading list of tickets"
        ' Call GetTickets()
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
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False, True, Globals.ThisAddIn.timingFile)

        wd = ndt.GiveMeChrome(True)
        Globals.ThisAddIn.Application.ActiveWindow.Activate()
        Dim Downloads As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads")
        Dim file As String, oldfile As String

        oldfile = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        wd.Navigate.GoToUrl("http://nextdesk/reports/ticketsClosedBySystem.php")


        wd.FindElementByName("system_selected").SendKeys("SAS")

        ' this doesn't work: wd.FindElementByName("start_date").SendKeys("01/01/2018")

        MsgBox("Please set an appropriate start date before continuing")

        Try
            wd.FindElementByName("report").Click()
        Catch
            If wd.Title <> "nextDesk | Tickets Closed By System" Then
                wd.Quit()
                Closeme()
                Exit Sub
            End If

        End Try

        For i = 1 To 10
            Call SetText(Label1.Text & ".")
            Threading.Thread.Sleep(300)
            If interrupt Then
                wd.Quit()
                Exit Sub
            End If
        Next
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            For i = 1 To 10
                Call SetText(Label1.Text & ".")
                Threading.Thread.Sleep(1000)
                If interrupt Then
                    Closeme()

                    wd.Quit()
                    Exit Sub
                End If
            Next
        End If
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            MsgBox("File did not download correctly.", vbAbort)

            Closeme()
            wd.Quit()
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
                Exit For
            End If
            ndt.ticketNumber = ticket

            ticketDetails.Add(ndt.DEPScrape(wd))
            j += 1
            timeTaken = timeSoFar.Elapsed.TotalSeconds
            estimatedTotalTime = CInt(timeTaken / (j / ignoredTickets.Count))
            Call SetText("Read " & j & " of " & ignoredTickets.Count & " ignored tickets.")
            Call SetTextTwo("About " & PrettyString(TimeSpan.FromSeconds(estimatedTotalTime - timeTaken)) & " remaining")
            'If j = 5 Then Exit For
        Next

        wd.Quit()

        Call SetTextTwo("Writing output to excel")
        Call WriteToExcel(ticketDetails)

        Closeme()


    End Sub

    Private Sub WriteToExcel(ticketDetails As List(Of Dictionary(Of String, String)))
        Dim i As Integer
        Dim tSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add()
        Globals.ThisAddIn.Application.ScreenUpdating = False
        tSheet.Cells(1, 1).value = "TicketNumber"
        tSheet.cells(1, 2).value = "First Ignored"
        tSheet.Cells(1, 3).value = "Account Manager"
        tSheet.Cells(1, 4).value = "Client"
        For i = 0 To ticketDetails.Count - 1
            Try
                tSheet.Cells(i + 2, 1).value = ticketDetails(i)("TicketNumber")
                tSheet.Cells(i + 2, 2).value = ticketDetails(i)("Closed")
                tSheet.Cells(i + 2, 3).value = ticketDetails(i)("AM")
                tSheet.Cells(i + 2, 4).value = ticketDetails(i)("Client")
            Catch ex As Exception
                Debug.WriteLine("Error writing " & i)
                Debug.WriteLine(ex.Message)
            End Try

        Next
        Globals.ThisAddIn.Application.ScreenUpdating = True
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
            Dim d As New SetTextCallback(AddressOf SetTextTwo)
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
        ElseIf ticketData.Last.ToLower.Contains(CloseStale.CloseMessage.ToLower) Then
            Return True
        ElseIf ticketData.Last.ToLower.Contains("dep please advise and i will reopen this ticket") Then
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