Imports System.Diagnostics

Public Class PivotMail
    Public interrupt As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim totalWork As Integer = CountMails()

        Call SetText(totalWork & " Emails to write.")

        If Not WriteMails() = totalWork Then
            MsgBox("The process did Not complete", vbCritical)

        End If

        Call Closeme()


    End Sub

    Private Function WriteMails() As Integer
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim i As Integer = 2
        Dim j As Integer
        Dim MailCount As Integer = 0

        Dim AM_email As String, htmlTable As String

        While oXlWs.Cells(i, 3).value <> ""
            AM_email = oXlWs.Cells(i, 3).value
            j = i
            While oXlWs.Cells(j, 3).value = AM_email
                j = j + 1
            End While
            htmlTable = MakeTable(i, j - 1)




            i = j
            MailCount += 1
        End While


        Return MailCount
    End Function

    Private Function MakeTable(i As Integer, j As Integer) As String
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet


        MakeTable = "<table><tr><td>Customer Name</td><td>Net Units Bought (in 2018)</td></tr>" & vbCrLf


        For line As Integer = i To j
            MakeTable &= TableLineHTML(oXlWs.Cells(line, 4).value, oXlWs.Cells(line, 5).value)
        Next

        MakeTable &= "</table>"

    End Function

    Private Function CountMails() As Integer
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim i As Integer = 2
        Dim LineCount As Integer = 0
        While oXlWs.Cells(i, 3).value <> ""
            If oXlWs.Cells(i, 3).value <> oXlWs.Cells(i - 1, 3).value Then
                LineCount += 1
            End If
            i += 1
        End While

        Return LineCount

    End Function

    Private Sub PivotMail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Creating list of emails"
        ' Call GetTickets()
        BackgroundWorker1.RunWorkerAsync()
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

    Function TableLineHTML(cellOne As String, CellTwo As String) As String
        Return "<tr>" & vbTab & "<td>" & vbCrLf & vbTab & vbTab & cellOne & vbCrLf & vbTab &
            "</td>" & vbCrLf & vbTab & "<td>" & vbCrLf & vbTab & vbTab & CellTwo &
            vbCrLf & vbTab & "</td>" & vbCrLf & "</tr>"
    End Function
End Class