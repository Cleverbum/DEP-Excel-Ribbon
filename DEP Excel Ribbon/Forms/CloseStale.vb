Imports System.Diagnostics
Imports System.IO
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class CloseStale
    Public Const CloseMessage As String =
        "This ticket appears to be stale, please let me know if it is still required"
    Public interrupt As Boolean = False
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        wd.Navigate.GoToUrl("http://nextdesk/index.php?outputcsv=1&newbin=1930")





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


        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        Dim newlines As Tuple(Of Integer, Integer)
        newlines = ProcessFile(file)

        Closeme()

        MsgBox("There were " & newlines.Item1 & " tickets in the bin. " &
               newlines.Item2 & " were then closed")

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
    Function Readfile(file As String, Optional startPoint As Integer = 1) As Integer
        Dim MyReader As New FileIO.TextFieldParser(file)
        Dim currentLine As String()

        MyReader.TextFieldType = FileIO.FieldType.Delimited
        MyReader.HasFieldsEnclosedInQuotes = True
        MyReader.SetDelimiters(",")
        Dim i As Integer, bin As Integer
        i = startPoint

        Dim sheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add

        While Not MyReader.EndOfData
            If interrupt Then
                Readfile = i
                Exit Function
            End If
            Try
                currentLine = MyReader.ReadFields
                If i = 1 Then
                    For j = 0 To currentLine.Count - 1
                        If currentLine(j).ToLower.Equals("bin") Then
                            bin = j
                        End If
                        sheet.Cells(i, j + 2).value = currentLine(j)
                    Next
                    i += 1
                ElseIf currentLine(bin).StartsWith("Apple/Dep", vbTextCompare) Then

                    For j = 0 To currentLine.Count - 1
                        sheet.Cells(i, j + 2).value = currentLine(j)
                    Next

                    sheet.Cells(i, 1).value = "Link"

                    If i > 1 Then
                        sheet.Hyperlinks.Add(sheet.Cells(i, 1), "http://nextdesk/ticket.php?setmode=Log&id=" & sheet.Cells(i, 2).value)
                    End If

                    If ToClose(sheet, i) Then
                        sheet.Cells(i, currentLine.Count).value = "Close"
                    End If

                    i += 1
                End If


            Catch ex As FileIO.MalformedLineException
                Debug.WriteLine("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
        End While
        MyReader.Close()
        Return i

    End Function


    Function ProcessFile(file As String) As Tuple(Of Integer, Integer)
        Dim MyReader As New FileIO.TextFieldParser(file)
        Dim currentLine As String()

        MyReader.TextFieldType = FileIO.FieldType.Delimited
        MyReader.HasFieldsEnclosedInQuotes = True
        MyReader.SetDelimiters(",")
        Dim i As Integer, deleted As Integer
        Dim bin As Integer, typeloc As Integer
        i = 1
        deleted = 0
        Call SetText("Processing Tickets...")
        While Not MyReader.EndOfData
            If interrupt Then
                ProcessFile = New Tuple(Of Integer, Integer)(i - 1, deleted)
                Exit Function
            End If
            Try
                currentLine = MyReader.ReadFields
                If i = 1 Then
                    For j = 0 To currentLine.Count - 1
                        If currentLine(j).ToLower.Equals("bin") Then
                            bin = j
                        End If
                        If currentLine(j).ToLower.Equals("type") Then
                            typeloc = j
                        End If
                    Next
                    i += 1
                ElseIf currentLine(bin).StartsWith("Apple/Dep", vbTextCompare) And
                    currentLine(typeloc).StartsWith("Apple", vbTextCompare) Then


                    If ToClose(currentLine.Last) Then
                        Call SetText("Closing ticket " & currentLine(0))
                        CloseTicket(currentLine(0))
                        deleted += 1
                        Call SetText("Processing Tickets...")
                    End If

                    i += 1
                End If


            Catch ex As FileIO.MalformedLineException
                Debug.WriteLine("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
        End While
        MyReader.Close()
        Return New Tuple(Of Integer, Integer)(i - 1, deleted)

    End Function



    Function ToClose(sheet As Excel.Worksheet, line As Integer) As Boolean
        If CStr(sheet.Cells(1, 10).value).StartsWith("Time", vbTextCompare) And
                CStr(sheet.Cells(1, 5).value).StartsWith("Type", vbTextCompare) Then
            If CStr(sheet.Cells(line, 5).value).StartsWith("Apple", vbTextCompare) Then
                Dim time As Double, tmpTime As String
                tmpTime = sheet.Cells(line, 10).value
                time = CDbl(tmpTime.Split(" ")(0))
                ToClose = time > 10
            Else
                ToClose = False
            End If
        Else
            ToClose = False
        End If

    End Function
    Function ToClose(timeString As String) As Boolean
        Dim time As Double
        time = CDbl(timeString.Split(" ")(0))
        ToClose = time > 10
    End Function

    Sub CloseTicket(ticketnumber As Integer)
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
            .ticketNumber = ticketnumber
        }
        ndt.CloseTicket(CloseMessage)
    End Sub

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