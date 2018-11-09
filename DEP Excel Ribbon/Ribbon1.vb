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


        Dim lines As New List(Of clsDepLine), snglLine As clsDepLine

        doDistiMail = (MsgBox("Would you like to generate the emails to distribution at the same time", vbYesNo) = vbYes)
        Dim mailPath As String

        mailPath = Environ("TEMP") & "\DistiEmail.msg"

        i = 2

        While oXlWs.Cells(i, 1).value <> ""
            snglLine = readExcelLine(oXlWs, i)
            lines.Add(snglLine)
            i += 1
        End While

        myCount = discardNoDEP(lines) ' number of lines removed

        For Each line As clsDepLine In lines
            Dim ndt As New clsNextDeskTicket.clsNextDeskTicket

            line.NDT_Number = ndt.createTicket(2, line.toTicket())
            ndt.ticketNumber = line.NDT_Number
            ndt.AddToNotify(Globals.ThisAddIn.findAlias(line.Account_Manager_Email))
            If line.Action.Equals("Reg", Globals.ThisAddIn.ignoreCase) And doDistiMail Then
                Dim distiMail As New clsDistiEmail, thisMail As Outlook.MailItem
                thisMail = distiMail.generateMail(line)

                If thisMail.To IsNot Nothing Then ' Techdata don't do emails so techdata lines have no "to" address
                    thisMail.Display()
                End If
                thisMail.SaveAs(mailPath)

                ndt.UpdateNextDeskAttach(mailPath, Globals.ThisAddIn.distiEmailMessage)
                My.Computer.FileSystem.DeleteFile(mailPath)
            End If
        Next

    End Sub
    Function readExcelLine(ByRef worksheet As Object, ByVal i As Integer) As clsDepLine
        Dim tmpLine As New clsDepLine

        tmpLine.Entity = worksheet.Cells(i, 1).value
        tmpLine.Account_Number = worksheet.Cells(i, 2).value
        tmpLine.Company = worksheet.Cells(i, 3).value
        tmpLine.DEP = worksheet.Cells(i, 4).value
        tmpLine.Post_Code = worksheet.Cells(i, 5).value
        tmpLine.Customer_PO = worksheet.Cells(i, 6).value
        tmpLine.Sales_ID = worksheet.Cells(i, 7).value
        tmpLine.Order_Date = worksheet.Cells(i, 8).value
        tmpLine.Invoice_Date = worksheet.Cells(i, 9).value
        tmpLine.Order_Type_Desc = worksheet.Cells(i, 10).value
        tmpLine.Invoice_ID = worksheet.Cells(i, 11).value
        tmpLine.Item_ID = worksheet.Cells(i, 12).value
        tmpLine.Manufacturer_Part_Number = worksheet.Cells(i, 13).value
        tmpLine.Item_Name = worksheet.Cells(i, 14).value
        tmpLine.Sub_Cat = worksheet.Cells(i, 15).value
        tmpLine.Sub_Cat_Description = worksheet.Cells(i, 16).value
        tmpLine.Manufacturer_Name = worksheet.Cells(i, 17).value
        tmpLine.Units = worksheet.Cells(i, 18).value
        tmpLine.POto_Supplier = worksheet.Cells(i, 19).value
        tmpLine.Suppliername = worksheet.Cells(i, 20).value
        tmpLine.Account_Manager = worksheet.Cells(i, 21).value
        tmpLine.Account_Manager_Email = worksheet.Cells(i, 22).value
        tmpLine.POType = worksheet.Cells(i, 23).value
        tmpLine.POCreated_Date = worksheet.Cells(i, 24).value

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
            tmpLine.Action = "Ticket"
        ElseIf tmpLine.DEP.ToLower.Contains("reg") Then
            tmpLine.Action = "Reg"
        Else
            tmpLine.Action = "Discard"

        End If


        Return tmpLine
    End Function
    Function discardNoDEP(ByRef rawLines As List(Of clsDepLine)) As Integer
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
        Debug.WriteLine(Globals.ThisAddIn.findAlias("Sam Brennan"))
    End Sub
End Class
