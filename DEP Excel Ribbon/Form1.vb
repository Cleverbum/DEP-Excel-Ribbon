Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class Form1

    Public Interrupt As Boolean = False
    Public timeEstimate As TimeEstimator

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Reading in the Excel file"
        Label2.Text = "Calculating Duration Estimate"
        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Call MakeTickets()
        Call Closeme()
    End Sub

    Sub MakeTickets()
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet
        Dim myCount As Integer, i As Integer, doDistiMail As Boolean
        Dim errorCount As Integer = 0

        Dim lines As New List(Of ClsDepLine), snglLine As ClsDepLine

        Dim TDLines As New List(Of ClsDepLine)

        UpdateStatus("Acting on " & oXlWb.Name)

        doDistiMail = (MsgBox("Would you like to generate the emails to distribution at the same time", vbYesNo) = vbYes)
        Dim mailPath As String

        mailPath = Environ("TEMP") & "\DistiEmail.msg"

        i = 2

        While oXlWs.Cells(i, 1).value <> ""
            Try
                snglLine = ReadExcelLine(oXlWs, i)
                lines.Add(snglLine)
            Catch
                HighlightError(oXlWs.Cells(i, 1))
                errorCount += 1

            End Try

            i += 1
        End While
        Dim total As Long = lines.LongCount

        myCount = DiscardNoDEP(lines) ' number of lines removed

        UpdateStatus("Found " & total & " total lines. Discarded " & myCount)

        Call SetProgressMax(myCount + 1)



        i = 1
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False)
        Dim browser As Chrome.ChromeDriver
        browser = ndt.GiveMeChrome(False)

        For Each line As ClsDepLine In lines
            Call SetProgress(i)
            If interrupt Then Exit For

            Call UpdateStatus("Creating ticket " & i & " of " & lines.Count)

            Try
                line.NDT_Number = ndt.CreateTicket(2, line.ToTicket(), browser)
            Catch ex As Exception
                line.NDT_Number = 0
            End Try

            If line.NDT_Number = 0 Then
                HighlightError(line.Sales_ID)
                errorCount += 1
                i += 1
                ' go to next "line"
                Continue For
            End If

            ndt.ticketNumber = line.NDT_Number

            If line.Units > 10 Then
                ndt.UpdateNextDesk("Please note that there were " & line.Units & " units on this order - the below serials list is not exhaustive", browser)
            End If

            Dim tmpAlias As String = Globals.ThisAddIn.FindAlias(line.Account_Manager_Email)

            If interrupt Then Exit For

            Call SetProgress(i + 1.0 / 3.0)

            If tmpAlias <> "NN" Then
                Try
                    ndt.AddToNotify(tmpAlias, browser)
                Catch ex As Exception
                    Debug.WriteLine("Failed during notify")
                    Debug.WriteLine(ex.Message)
                End Try

            Else
                Try
                    ndt.UpdateNextDesk("Could not find the nextdesk username for " & line.Account_Manager_Email, browser)

                Catch ex As Exception
                    Debug.WriteLine("Failed during update")
                    Debug.WriteLine(ex.Message)
                End Try
            End If

            If interrupt Then Exit For

            Call SetProgress(i + 2.0 / 3.0)

            If line.Action.Equals("Reg", comparisonType:=ThisAddIn.ignoreCase) And
                        doDistiMail And line.Units < 11 Then



                Dim distiMail As New ClsDistiEmail, thisMail As Outlook.MailItem
                UpdateStatus("For ticket " & i & " of " & lines.Count & ": Generating an email if Required")
                thisMail = distiMail.GenerateMail(line)

                If thisMail.To IsNot Nothing Then ' Techdata don't do emails so techdata lines have no "to" address
                    thisMail.Display()
                    thisMail.SaveAs(mailPath)
                    thisMail.CC = ThisAddIn.ccList
                    thisMail.Send()

                    Try
                        ndt.UpdateNextDeskAttach(mailPath, distiEmailMessage)
                    Catch ex As Exception
                        HighlightError(line.Sales_ID)
                        errorCount += 1
                        Debug.WriteLine("Failed during attach")
                        Debug.WriteLine(ex.Message)
                    End Try
                    Try
                        My.Computer.FileSystem.DeleteFile(mailPath)
                    Catch ex As Exception
                        HighlightError(line.Sales_ID)
                        errorCount += 1
                        Debug.WriteLine("Failed during file delete")
                        Debug.WriteLine(ex.Message)
                    End Try

                Else
                    Try
                        TDLines.Add(line)
                        ndt.UpdateNextDesk(Replace(NoEmailSent, "%SupplierName%", line.Suppliername), browser)
                    Catch ex As Exception
                        HighlightError(line.Sales_ID)
                        errorCount += 1
                        Debug.WriteLine("Failed during  update")
                        Debug.WriteLine(ex.Message)
                    End Try
                End If
            ElseIf line.Action.Equals("Only", ThisAddIn.ignoreCase) Then
                Try
                    ndt.UpdateNextDesk(OnlyMessage, browser)
                Catch ex As Exception
                    HighlightError(line.Sales_ID)
                    errorCount += 1
                    Debug.WriteLine("Failed during update")
                    Debug.WriteLine(ex.Message)
                End Try
            ElseIf line.Action.Equals("Fake Serial", ThisAddIn.ignoreCase) Then
                Try
                    ndt.UpdateNextDesk(FakeMessage, browser)
                Catch ex As Exception
                    HighlightError(line.Sales_ID)
                    errorCount += 1
                    Debug.WriteLine("Failed during update")
                    Debug.WriteLine(ex.Message)
                End Try
            ElseIf line.Action.Equals("Ticket", ThisAddIn.ignoreCase) Then
                If Not line.Order_Type_Desc.ToLower.Contains("return") Then
                    Try
                        ndt.UpdateNextDesk(DepQuestion, browser)
                    Catch ex As Exception
                        HighlightError(line.Sales_ID)
                        errorCount += 1
                        Debug.WriteLine("Failed during update")
                        Debug.WriteLine(ex.Message)
                    End Try
                    Try
                        Call Send_AM_Email(line)
                    Catch ex As Exception
                        HighlightError(line.Sales_ID)
                        errorCount += 1
                        Debug.WriteLine("Failed during mail generation")
                        Debug.WriteLine(ex.Message)
                    End Try
                End If

            End If
            i += 1
        Next

        browser.Quit()

        'If TDLines.Count > 0 AndAlso
        '   MsgBox("Do you want to do the Techdata Regsitrations now?", vbYesNo) = vbYes Then
        '    For Each line In TDLines
        '        If Not RegisterTechdata(line) Then
        '            HighlightError(line.Sales_ID)
        '            errorCount += 1
        '            Debug.WriteLine("Failed during TD Registration")

        '        End If
        '    Next

        'End If

        UpdateStatus("All Done!")

        If errorCount > 0 Then
            MsgBox("There were " & errorCount & " errors during creation of tickets, these have been highlighted red.")
        Else
            MsgBox("Completed tasks with no errors")
        End If

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
    Private Sub HighlightError(sales_ID As Long)
        Dim tsheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet
        Dim tCell As Excel.Range = tsheet.Cells.Find(sales_ID)

        With tCell.EntireRow.Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
    End Sub

    Private Sub HighlightError(tCell As Excel.Range)
        With tCell.EntireRow.Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
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
                Try
                    tmpLine.Customer_DEP_ID = tmpLine.DEP.Split(" ")(0)
                Catch
                    tmpLine.Customer_DEP_ID = 0
                End Try
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

        'if there is a clause to "only" register x then don't try automated submission

        If tmpLine.Action.Equals("Reg", ThisAddIn.ignoreCase) Then
            If tmpLine.DEP.ToLower.Contains("only") Then
                tmpLine.Action = "Only"
            End If
        End If

        'if it includes a fake serial then modify later behaviour

        If tmpLine.Action <> "Discard" Then
            If FakeSerials(tmpLine.Serials) Then
                tmpLine.Action = "Fake Serials"
            End If
        End If

        'if the supplier is GBM discard - we can't do anything

        If tmpLine.Suppliername.ToLower.Contains("gbm") Then
            tmpLine.Action = "Discard"
        End If

        Return tmpLine
    End Function
    Function DiscardNoDEP(ByRef rawLines As List(Of ClsDepLine)) As Integer
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

    Private Sub UpdateStatus(message As String)
        Globals.ThisAddIn.Application.StatusBar = message
        Call SetText(message)
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        interrupt = True
    End Sub

    Private Sub SetProgress(ByVal [progress] As Double)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetProgressCallback(AddressOf SetProgress)
            Me.Invoke(d, New Object() {[progress]})
        Else
            If [progress] = CInt([progress]) Then
                Dim timeLeft As TimeSpan = timeEstimate.TimeRemaining([progress])
                Me.Label2.Text = "About " & PrettyString(timeLeft) & " remaining."
            End If

            Me.ProgressBar1.Value = [progress]

        End If
    End Sub
    Delegate Sub SetProgressCallback(ByVal [progress] As Double)

    Private Sub SetProgressMax(ByVal [progressMax] As Long)

        timeEstimate = New TimeEstimator(progressMax, 30.0)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetProgressMaxCallback(AddressOf SetProgressMax)
            Me.Invoke(d, New Object() {[progressMax]})
        Else

            Me.ProgressBar1.Maximum = [progressMax]
        End If
    End Sub
    Delegate Sub SetProgressMaxCallback(ByVal [progressMax] As Long)



End Class