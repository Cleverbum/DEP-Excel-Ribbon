Imports System.Diagnostics
Imports Microsoft.Office.Interop

Public Class PivotMail
    Public interrupt As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim totalWork As Integer = CountMails()
        If totalWork = 0 Then
            Call Closeme()
            Exit Sub
        End If
        Call SetText(totalWork & " Emails to write.")

        If Not WriteMails() = totalWork Then
            MsgBox("The process did not complete", vbCritical)

        End If

        Call Closeme()


    End Sub

    Private Function WriteMails() As Integer
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim i As Integer = 2
        Dim j As Integer, mailSuccess As Boolean
        Dim MailCount As Integer = 0

        Dim AM_email As String, htmlTable As String

        While oXlWs.Cells(i, 3).value <> ""

            If interrupt Then Return MailCount

            AM_email = oXlWs.Cells(i, 3).value
            j = i
            While oXlWs.Cells(j, 3).value = AM_email
                j = j + 1
            End While
            htmlTable = MakeTable(i, j - 1)


            mailSuccess = EditMail(AM_email, htmlTable)

            i = j
            If mailSuccess Then MailCount += 1
        End While


        Return MailCount
    End Function

    Private Function MakeTable(i As Integer, j As Integer) As String
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        MakeTable = "<style type=""text/css"">
                        .tg  {border-collapse:collapse;border-spacing:0;}
                        .tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 10px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
                        .tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:bold;padding:10px 10px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
                        tr:nth-child(odd) {background: #CCC}
                        tr:nth-child(even) {background: #FFF}
                        </style>
                        <table Class=""tg"">"

        MakeTable &= "<tr><th>Customer Name</th><th>Net Units Bought (in 2018)</th><th># of ""Apple"" Orders (in 2018)</th><</tr>" & vbCrLf


        For line As Integer = i To j
            MakeTable &= TableLineHTML(StrConv(oXlWs.Cells(line, 4).value.ToString, VbStrConv.ProperCase),
                                       oXlWs.Cells(line, 6).value.ToString,
                                       oXlWs.Cells(line, 5).value.ToString)
        Next

        MakeTable &= "</table>"

    End Function

    Private Function CountMails() As Integer
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim i As Integer = 2
        Dim LineCount As Integer = 0
        While oXlWs.Cells(i, 3).value <> ""
            If interrupt Then Return 0
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


    Function TableLineHTML(cellOne As String, CellTwo As String, CellThree As String) As String
        Return "<tr>" & vbTab & "<td>" & vbCrLf & vbTab & vbTab & cellOne & vbCrLf & vbTab &
            "</td>" & vbCrLf & vbTab & "<td>" & vbCrLf & vbTab & vbTab & CellTwo &
            vbCrLf & vbTab & "</td>" & vbCrLf & "<td>" & vbCrLf & vbTab & vbTab & CellThree &
            vbCrLf & vbTab & "</td>" & vbCrLf & "</tr>"
    End Function

    Private Function EditMail(to_address As String, table As String) As Boolean

        Dim templateFile As String = Environ("TEMP") & "\AM DEP Email.oft"
        Try



            My.Computer.FileSystem.WriteAllBytes(templateFile, My.Resources.No_DEP_Mail, False)
        Catch
            Debug.WriteLine("Error writing to FS")
            Return False
        End Try


        Dim AppOutlook As New Outlook.Application
        Dim amEmail As Outlook.MailItem
        amEmail = AppOutlook.CreateItemFromTemplate(templateFile)


        Dim outlookNameSpace As Outlook.NameSpace = AppOutlook.GetNamespace("MAPI")
        Dim myAddressList As Outlook.AddressList = outlookNameSpace.GetGlobalAddressList

        Dim am_name As String
        Try
            Dim objAEntry As Outlook.AddressEntry

            to_address = to_address.Replace("@uk.insight.com", "@insight.com")

            'below are corrections between iCare email addresses and Nextdesk Email addresses
            to_address = to_address.Replace("Scott.Waggstaff@insight.com", "Scott.Wagstaff@insight.com")



            'do final lookup of email to "recipient"
            objAEntry = AppOutlook.Session.CreateRecipient(to_address).AddressEntry

            am_name = objAEntry.GetExchangeUser.FirstName
        Catch
            Debug.WriteLine("Error finding firstname")
            am_name = ""
        End Try

        Try
            With amEmail
                .To = to_address
                .CC = "Chapman, Duncan <Duncan.Chapman@insight.com>; Ings, Jenni <Jenni.Ings@insight.com>"

                .HTMLBody = .HTMLBody.Replace("%TABLE%", table)
                .HTMLBody = .HTMLBody.Replace("%AM%", am_name)


                .Display()
            End With
        Catch
            Debug.WriteLine("Error modifying template")
            Return False
        End Try

        Return True
    End Function
End Class