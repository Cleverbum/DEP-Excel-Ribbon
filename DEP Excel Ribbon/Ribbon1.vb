﻿Imports System.Diagnostics
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
                thisMail = distiMail.generateMail(line)

                If thisMail.To IsNot Nothing Then ' Techdata don't do emails so techdata lines have no "to" address
                    thisMail.Display()
                    thisMail.SaveAs(mailPath)

                    ndt.UpdateNextDeskAttach(mailPath, Globals.ThisAddIn.distiEmailMessage)
                    My.Computer.FileSystem.DeleteFile(mailPath)
                Else
                    ndt.UpdateNextDesk("No mail was sent for this as the distributor is " & line.Suppliername & ". Please complete their process manually.")
                End If
            ElseIf line.action.Equals("Only", Globals.ThisAddIn.ignoreCase) Then
                ndt.UpdateNextDesk("There is an 'Only' condition in this customer's registration preferences, and so this registration will need to be completed manually. Thanks.")
            ElseIf line.Action.Equals("Ticket", Globals.ThisAddIn.ignoreCase) Then
                ndt.UpdateNextDesk("Hi, this shipped yesterday, would the client like this to be added to DEP? If so, please provide DEP ID.  Would the customer also like all Apple devices adding to DEP when shipped Thanks")
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
        Dim frm As New Form3
        frm.Show()
    End Sub
End Class
