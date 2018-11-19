Imports System.Diagnostics
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet
        Dim myCount As Integer, i As Integer, doDistiMail As Boolean


        Dim lines As New List(Of clsDepLine), snglLine As ClsDepLine

        MsgBox("Acting on " & oXlWb.Name)

        doDistiMail = (MsgBox("Would you like to generate the emails to distribution at the same time", vbYesNo) = vbYes)
        Dim mailPath As String

        mailPath = Environ("TEMP") & "\DistiEmail.msg"

        i = 2

        While oXlWs.Cells(i, 1).value <> ""
            snglLine = readExcelLine(oXlWs, i)
            lines.Add(snglLine)
            i += 1
        End While

        MsgBox("Found " & lines.Count & " total lines.")

        myCount = discardNoDEP(lines) ' number of lines removed

        MsgBox("Discarded " & myCount)
        i = 1
        For Each line As clsDepLine In lines
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
            Globals.ThisAddIn.Application.StatusBar = "Creating ticket " & i & " of " & lines.Count

            Try
                line.NDT_Number = ndt.CreateTicket(2, line.toTicket())
            Catch ex As Exception
                line.NDT_Number = 0
            End Try

            If line.NDT_Number = 0 Then
                MsgBox("Error creating ticket for Order: " & line.Sales_ID & " Exiting")
                Exit Sub
            End If

            ndt.ticketNumber = line.NDT_Number

            If line.Units > 10 Then
                ndt.UpdateNextDesk("Please note that there were " & line.Units & " units on this order - the below serials list is not exhaustive")
            End If

            Dim tmpAlias As String = Globals.ThisAddIn.FindAlias(line.Account_Manager_Email)


            If tmpAlias <> "NN" Then
                ndt.AddToNotify(tmpAlias)
            Else
                ndt.UpdateNextDesk("Could not find the nextdesk username for " & line.Account_Manager_Email)
                MsgBox("Could not find the nextdesk username for " & line.Account_Manager_Email & ". Please click OK to continue without adding them to the ticket")
            End If

            If line.Action.Equals("Reg", Globals.ThisAddIn.ignoreCase) And
                        doDistiMail And line.Units < 11 Then
                Dim distiMail As New clsDistiEmail, thisMail As Outlook.MailItem
                Globals.ThisAddIn.Application.StatusBar = "Generating an email if Required"
                thisMail = distiMail.generateMail(line)

                If thisMail.To IsNot Nothing Then ' Techdata don't do emails so techdata lines have no "to" address
                    thisMail.Display()
                    thisMail.SaveAs(mailPath)

                    ndt.UpdateNextDeskAttach(mailPath, Globals.ThisAddIn.distiEmailMessage)
                    My.Computer.FileSystem.DeleteFile(mailPath)
                Else
                    ndt.UpdateNextDesk("No mail was sent for this as the distributor is " & line.Suppliername & ". DEP Team: Please complete their process manually.")
                End If
            ElseIf line.action.Equals("Only", Globals.ThisAddIn.ignoreCase) Then
                ndt.UpdateNextDesk("There is an 'Only' condition in this customer's registration preferences, and so this registration will need to be completed manually. Thanks.")
            ElseIf line.Action.Equals("Fake Serial", Globals.ThisAddIn.ignoreCase) Then
                ndt.UpdateNextDesk("It seems that some of the serial numbers recorded in iCare do not match normal Apple patterns - please can you investigate this prior to submitting these for DEP.")
            ElseIf line.Action.Equals("Ticket", Globals.ThisAddIn.ignoreCase) Then
                If Not line.Order_Type_Desc.ToLower.Contains("return") Then
                    ndt.UpdateNextDesk("Hi, this shipped yesterday, would the client like this to be added to DEP? If so, please provide DEP ID.  Would the customer also like all Apple devices adding to DEP when shipped Thanks")
                    Call Send_AM_Email(line)
                End If

            End If
            i += 1
        Next
        Globals.ThisAddIn.Application.StatusBar = "All Done!"
    End Sub
    Function ReadExcelLine(ByRef worksheet As Object, ByVal i As Integer) As ClsDepLine
        Dim tmpLine As New ClsDepLine With {
            .Entity = worksheet.Cells(i, 1).value,
            .Account_Number = worksheet.Cells(i, 2).value,
            .Company = worksheet.Cells(i, 3).value,
            .DEP = worksheet.Cells(i, 4).value,
            .Post_Code = worksheet.Cells(i, 5).value,
            .Customer_PO = worksheet.Cells(i, 6).value,
            .Sales_ID = worksheet.Cells(i, 7).value,
            .Order_Date = worksheet.Cells(i, 8).value,
            .Invoice_Date = worksheet.Cells(i, 9).value,
            .Order_Type_Desc = worksheet.Cells(i, 10).value,
            .Invoice_ID = worksheet.Cells(i, 11).value,
            .Item_ID = worksheet.Cells(i, 12).value,
            .Manufacturer_Part_Number = worksheet.Cells(i, 13).value,
            .Item_Name = worksheet.Cells(i, 14).value,
            .Sub_Cat = worksheet.Cells(i, 15).value,
            .Sub_Cat_Description = worksheet.Cells(i, 16).value,
            .Manufacturer_Name = worksheet.Cells(i, 17).value,
            .Units = worksheet.Cells(i, 18).value,
            .POto_Supplier = worksheet.Cells(i, 19).value,
            .Suppliername = worksheet.Cells(i, 20).value,
            .Account_Manager = worksheet.Cells(i, 21).value,
            .Account_Manager_Email = worksheet.Cells(i, 22).value,
            .POType = worksheet.Cells(i, 23).value,
            .POCreated_Date = worksheet.Cells(i, 24).value
        }

        Dim j As Integer, pSerials() As String
        ReDim pSerials(0 To 10)
        For j = 25 To 35
            pSerials(j - 25) = worksheet.Cells(i, j).value
        Next

        tmpLine.Serials = pSerials

        'if no dep = discard
        'if includes reg then reg
        'if blank but with account manager email then ticket
        'if blank and no email then discard

        If tmpLine.DEP Is Nothing OrElse (tmpLine.DEP = "" And
                    (tmpLine.Account_Manager_Email IsNot Nothing AndAlso
                    tmpLine.Account_Manager_Email <> "")
                ) Then
            If tmpLine.Account_Manager_Email Is Nothing Then
                tmpLine.Action = "Discard"
            Else
                tmpLine.Action = "Ticket"
            End If
        ElseIf tmpLine.DEP.ToLower.Contains("reg") Then
            If tmpLine.Account_Manager_Email Is Nothing Then
                tmpLine.Action = "Discard"
            Else
                tmpLine.Action = "Reg"
            End If
        Else
            tmpLine.Action = "Discard"

        End If

        If tmpLine.Action.Equals("Reg", Globals.ThisAddIn.ignoreCase) Then
            If tmpLine.DEP.ToLower.Contains("only") Then
                tmpLine.Action = "Only"
            End If
        End If

        If tmpLine.Action <> "Discard" Then
            If fakeSerials(tmpLine.Serials) Then
                tmpLine.Action = "Fake Serials"
            End If
        End If

        If tmpLine.Suppliername.ToLower.Contains("gbm") Then
            tmpLine.Action = "Discard"
        End If

        Return tmpLine
    End Function
    Function DiscardNoDEP(ByRef rawLines As List(Of clsDepLine)) As Integer
        Dim count As Integer, i As Integer
        count = 0
        For i = rawLines.Count - 1 To 0 Step -1
            If rawLines(i).Action = "Discard" Then
                rawLines.RemoveAt(i)
                count = count + 1
            End If


        Next
        Return count
    End Function

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim frm As New Form3
        frm.Show()
    End Sub



    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim frm As New Form4
        frm.Show()
    End Sub



    Function FakeSerials(serials As String()) As Boolean
        For Each serial In serials
            Try
                If serial.ToLower.StartsWith("po") Then
                    Return True
                End If
            Catch ex As Exception
                'error handler?
                Debug.Print(ex.Message)
            End Try
        Next
        Return False
    End Function

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wd As Chrome.ChromeDriver
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        wd = ndt.GiveMeChrome(True)
        Dim url As String = "http://nextdesk/ticket.php?id=5953026"

        wd.Navigate.GoToUrl(url)
        Dim depscrape As New Dictionary(Of String, String)

        Dim elements = wd.FindElementsByClassName("ticketCell")
        For Each element In elements
            If element.Text.ToLower.Contains("client name") Then
                depscrape.Add("Client", TrimClient(element.Text))
                depscrape.Add("AM", TrimAM(element.Text))
                depscrape.Add("Closed", TrimClosed(element.Text))
            End If

        Next

    End Sub
    Function TrimClosed(txt As String) As String

        Dim list As String() = txt.ToLower.Split(vbCrLf)

        For i = 0 To list.Count - 1
            If list(i).ToLower.Contains("closed") Then
                Return Trim(list(i + 1).Split(" ")(0).Replace(vbLf, " "))
            End If
        Next
        Return ""
    End Function
    Function TrimClient(txt As String) As String

        Dim list As String() = txt.ToLower.Split(vbCrLf)

        For i = 0 To list.Count - 1
            If list(i).ToLower.Contains("client name") Then
                Return Trim(list(i + 1).Replace(vbLf, " "))
            End If
        Next
        Return ""
    End Function
    Function TrimAM(txt As String) As String

        Dim list As String() = txt.ToLower.Split(vbCrLf)

        For i = 0 To list.Count - 1
            If list(i).ToLower.Contains("description") Then
                Dim words = list(i + 1).Split(" ")
                For Each word In words
                    If word.Contains("@") Then
                        Return word
                    End If
                Next
            End If
        Next
        Return ""
    End Function
End Class
