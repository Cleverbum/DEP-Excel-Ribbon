Imports System.Diagnostics
Imports System.IO
Imports OpenQA.Selenium

Public Class Form3
    Public interrupt As Boolean = False
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Downloading list of tickets"
        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call get_tickets()
    End Sub

    Sub get_tickets()
        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        Globals.ThisAddIn.Application.ActiveWindow.Activate()
        Dim Downloads As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads")
        Dim file As String, oldfile As String

        oldfile = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        wd.Navigate.GoToUrl("http://nextdesk/index.php?outputcsv=1&newbin=1930")





        For i = 1 To 5
            Call SetText(Label1.Text & ".")
            Threading.Thread.Sleep(1000)
        Next
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            For i = 1 To 10
                Call SetText(Label1.Text & ".")
                Threading.Thread.Sleep(1000)
            Next
        End If
        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

        If oldfile = file Then
            MsgBox("File did not download correctly.", vbAbort)
            Exit Sub
        End If



        Call SetText("File Downloaded.")


        file = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        Dim newlines As Integer = Readfile(file)
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
        Dim i As Integer
        i = startPoint
        Dim sheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add

        While Not MyReader.EndOfData
            Try
                currentLine = MyReader.ReadFields

                For j = 0 To currentLine.Count - 1
                    sheet.Cells(i, j + 2).value = currentLine(j)
                Next

                sheet.Cells(i, 1).value = "Link"

                If i > 1 Then
                    sheet.Hyperlinks.Add(sheet.Cells(i, 1), "http://nextdesk/ticket.php?setmode=Log&id=" & sheet.Cells(i, 2).value)
                End If

                i += 1
            Catch ex As FileIO.MalformedLineException
                Debug.WriteLine("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
        End While
        MyReader.Close()
        Return i

    End Function
End Class