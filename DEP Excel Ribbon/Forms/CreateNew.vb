Imports System.Diagnostics
Imports System.IO
Imports System.Net.Mail
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class CreateNew

    Public Interrupt As Boolean = False
    Public timeEstimate As TimeEstimator
    Private DoAll As Boolean
    Private DoTD As Boolean
    Private DoWC As Boolean
    Private debugFrm As DebugForm
    Private debugMode As Boolean
    Private timingFile As String = Environ("Temp") & "\timinglog.csv"

    Public Sub New(Optional tDoAll As Boolean = True, Optional showDebugInfo As Boolean = False, Optional tDoWC As Boolean = False, Optional tDoTD As Boolean = False)
        InitializeComponent()
        DoAll = tDoAll
        DoWC = tDoWC
        DoTD = tDoTD
        debugMode = showDebugInfo
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.ThisAddIn.RegistrationRunning = True
        Me.Label1.Text = "Reading in the Excel file"
        Me.Label2.Text = "Calculating Duration Estimate"
        'parallel
        BackgroundWorker1.RunWorkerAsync()

        If debugMode Then
            debugFrm = New DebugForm
            UpdateDebugMessage("Starting processes.")
        End If

        'testing version (no threading for debugging)
        'Call MakeTickets()
    End Sub

    Private Sub UpdateDebugMessage(MessageString As String)
        'write it to the Visual Studio Debugger for luck
        Debug.WriteLine(MessageString)

        'write it to the form if it's active
        If debugMode Then debugFrm.debugText &= MessageString & vbCrLf
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Interrupt = True
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Try
            My.Computer.FileSystem.DeleteFile(timingFile)
        Catch
        End Try

        Call MakeTickets()

        Globals.ThisAddIn.RegistrationRunning = False


        Call Closeme()

    End Sub

    Sub MakeTickets()
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet
        Dim myCount As Integer, i As Integer, doDistiMail As Boolean
        Dim errorCount As Integer = 0

        Dim lines As New List(Of ClsDepLine), snglLine As ClsDepLine

        Dim TDLines As New List(Of ClsDepLine)

        Dim WClines As New List(Of ClsDepLine)

        UpdateStatus("Acting on " & oXlWb.Name)

        'no longer need to ask to do this - just assume yes
        doDistiMail = True

        Dim mailPath As String

        mailPath = Environ("TEMP") & "\DistiEmail.msg"

        i = 2

        While oXlWs.Cells(i, 1).value <> ""
            Try
                snglLine = ReadExcelLine(oXlWs, i)
                lines.Add(snglLine)
            Catch
                Globals.ThisAddIn.HighlightError(oXlWs.Cells(i, 1))
                errorCount += 1

            End Try

            i += 1
            If debugMode Then UpdateDebugMessage("Read in line " & i)
        End While
        Dim total As Long = lines.LongCount

        myCount = DiscardNoDEP(lines) ' number of lines removed


        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False, True, timingFile)

        If DoAll Then
            UpdateStatus("Found " & total & " total lines. Discarded " & myCount)

            Call SetProgressMax(lines.Count)
            i = 1

            Dim browser As Chrome.ChromeDriver
            browser = ndt.GiveMeChrome(False)

            For Each line As ClsDepLine In lines
                Call SetProgress(i - 1)

                If Not Globals.ThisAddIn.OnIntranet Then
                    Globals.ThisAddIn.HighlightError(line.Serials(0))
                    UpdateStatus("Can't access nextdesk, skipping this line and retrying")
                    Threading.Thread.Sleep(TimeSpan.FromSeconds(1))
                    Continue For ' This tells it to skip the rest of the for loop and go to the next depline
                End If

                If Interrupt Or Not Globals.ThisAddIn.RegistrationRunning Then
                    Globals.ThisAddIn.RegistrationRunning = False
                    Exit Sub
                End If

                Call UpdateStatus("Creating ticket " & i & " of " & lines.Count)

                Try
                    line.NDT_Number = ndt.CreateTicket(2, line.ToTicket(), browser)

                Catch ex As Exception
                    line.NDT_Number = 0
                End Try

                If debugMode Then UpdateDebugMessage("Ticket created: " & line.NDT_Number)

                If line.NDT_Number = 0 Then
                    Globals.ThisAddIn.HighlightError(line.Serials(0))
                    errorCount += 1
                    i += 1
                    ' go to next "line"
                    Continue For
                End If

                ndt.TicketNumber = line.NDT_Number

                If line.Units > 10 Then
                    ndt.UpdateNextDesk("Please note that there were " & line.Units & " units on this order - the below serials list is not exhaustive", browser)
                    If debugMode Then UpdateDebugMessage("Too many units: " & line.Units)
                End If

                Dim tmpAlias As String = Globals.ThisAddIn.FindAlias(line.Account_Manager_Email)

                If Interrupt Then Exit For

                Call SetProgress(i - 1 + 1.0 / 3.0)

                If tmpAlias <> "NN" Then
                    Try
                        If debugMode Then UpdateDebugMessage("Adding to notify: " & tmpAlias)
                        ndt.AddToNotify(tmpAlias, browser)
                    Catch ex As Exception
                        UpdateDebugMessage("Failed during notify")
                        UpdateDebugMessage(ex.Message)
                    End Try

                Else
                    Try
                        If debugMode Then UpdateDebugMessage("No alias for: " & line.Account_Manager_Email)
                        If debugMode Then UpdateDebugMessage("Updating NDT")
                        ndt.UpdateNextDesk("Could not find the nextdesk username for " & line.Account_Manager_Email, browser)

                    Catch ex As Exception
                        UpdateDebugMessage("Failed during update")
                        UpdateDebugMessage(ex.Message)
                    End Try
                End If

                If Interrupt Then Exit For

                Call SetProgress(i - 1 + 2.0 / 3.0)

                If line.Action.Equals("Reg", comparisonType:=ThisAddIn.ignoreCase) And
                            doDistiMail And line.Units < 11 Then



                    Dim distiMail As New ClsDistiEmail, thisMail As Outlook.MailItem
                    UpdateStatus("For ticket " & i & " of " & lines.Count & ": Generating an email if Required")
                    thisMail = distiMail.GenerateMail(line)

                    If thisMail.To IsNot Nothing Then ' Techdata & Westcoast don't do emails so techdata lines have no "to" address
                        If debugMode Then UpdateDebugMessage("Sending Mail")
                        thisMail.Display()
                        thisMail.SaveAs(mailPath)
                        thisMail.CC = ThisAddIn.ccList
                        thisMail.Send()

                        Try
                            If debugMode Then UpdateDebugMessage("Attaching Mail to ticket")
                            ndt.UpdateNextDeskAttach(mailPath, distiEmailMessage)
                        Catch ex As Exception
                            Globals.ThisAddIn.HighlightError(line.Serials(0))
                            errorCount += 1
                            UpdateDebugMessage("Failed during attach")
                            UpdateDebugMessage(ex.Message)
                        End Try
                        Try
                            If debugMode Then UpdateDebugMessage("Deleting Mail temporary file")
                            My.Computer.FileSystem.DeleteFile(mailPath)
                        Catch ex As Exception
                            Globals.ThisAddIn.HighlightError(line.Serials(0))
                            errorCount += 1
                            UpdateDebugMessage("Failed during file delete")
                            UpdateDebugMessage(ex.Message)
                        End Try

                    Else
                        Try
                            If line.Suppliername.ToLower.Contains("tech data") Then
                                TDLines.Add(line)
                            ElseIf line.Suppliername.ToLower.Contains("westcoast") Then
                                WClines.Add(line)
                            End If

                            ndt.UpdateNextDesk(Replace(NoEmailSent, "%SupplierName%", line.Suppliername), browser)
                        Catch ex As Exception
                            Globals.ThisAddIn.HighlightError(line.Serials(0))
                            errorCount += 1
                            UpdateDebugMessage("Failed during  update")
                            UpdateDebugMessage(ex.Message)
                        End Try
                    End If
                ElseIf line.Action.Equals("Only", ThisAddIn.ignoreCase) Then
                    Try
                        If debugMode Then UpdateDebugMessage("Updating NDT with 'only' message")
                        ndt.UpdateNextDesk(OnlyMessage, browser)
                    Catch ex As Exception
                        Globals.ThisAddIn.HighlightError(line.Serials(0))
                        errorCount += 1
                        UpdateDebugMessage("Failed during update")
                        UpdateDebugMessage(ex.Message)
                    End Try
                ElseIf line.Action.Equals("Fake Serials", ThisAddIn.ignoreCase) Then
                    Try
                        If debugMode Then UpdateDebugMessage("Updating NDT re: Fake Serials")
                        ndt.UpdateNextDesk(FakeMessage, browser)
                    Catch ex As Exception
                        Globals.ThisAddIn.HighlightError(line.Serials(0))
                        errorCount += 1
                        UpdateDebugMessage("Failed during update")
                        UpdateDebugMessage(ex.Message)
                    End Try
                ElseIf line.Action.Equals("Ticket", ThisAddIn.ignoreCase) Then
                    If Not line.Order_Type_Desc.ToLower.Contains("return") Then
                        Try
                            If debugMode Then UpdateDebugMessage("Updating NDT with 'do you want to dep' message")
                            ndt.UpdateNextDesk(DepQuestion, browser)
                        Catch ex As Exception
                            Globals.ThisAddIn.HighlightError(line.Serials(0))
                            errorCount += 1
                            UpdateDebugMessage("Failed during update")
                            UpdateDebugMessage(ex.Message)
                        End Try
                        Try
                            If debugMode Then UpdateDebugMessage("Sending Account Manager 'do you want to DEP' eMail")
                            Call Send_AM_Email(line)
                        Catch ex As Exception
                            Globals.ThisAddIn.HighlightError(line.Serials(0))
                            errorCount += 1
                            UpdateDebugMessage("Failed during mail generation")
                            UpdateDebugMessage(ex.Message)
                        End Try
                    End If

                End If
                i += 1
            Next

            browser.Quit()
        Else ' We're only doing TechData or Westcoast
            For Each line As ClsDepLine In lines
                If line.Action.Equals("Reg", comparisonType:=ThisAddIn.ignoreCase) And
                        doDistiMail And line.Units < 11 Then

                    If line.Suppliername.ToLower.Contains("tech data") And DoTD Then
                        TDLines.Add(line)
                    ElseIf line.Suppliername.ToLower.Contains("westcoast") And DoWC Then
                        WClines.Add(line)
                    End If

                End If
            Next
        End If

        If TDLines.Count > 0 AndAlso
           MsgBox("Do you want to do the " & TDLines.Count & " Techdata Regsitrations now?", vbYesNo) = vbYes Then
            Dim wd As Chrome.ChromeDriver = DoTDLogin()
            For Each line In TDLines
                If Not RegisterTechdata(line, wd) Then
                    Globals.ThisAddIn.HighlightError(line.Serials(0))
                    errorCount += 1
                    ndt.TicketNumber = line.NDT_Number
                    ndt.UpdateNextDesk(tdFail)
                    UpdateDebugMessage("Failed during TD Registration")
                Else
                    ndt.TicketNumber = line.NDT_Number
                    ndt.UpdateNextDesk(tdSuccess)
                End If
            Next

            wd.Quit()

        End If
        If WClines.Count > 0 AndAlso
           MsgBox("Do you want to do the " & WClines.Count & " Westcoast Regsitrations now?", vbYesNo) = vbYes Then
            Dim wd As Chrome.ChromeDriver = DoWCLogin()
            For Each line In WClines
                If Not DoOneWC_DEP(line, wd) Then
                    Globals.ThisAddIn.HighlightError(line.Serials(0))
                    errorCount += 1
                    If DoAll Then
                        ndt.TicketNumber = line.NDT_Number
                        ndt.UpdateNextDesk(wcFail)

                    End If
                    UpdateDebugMessage("Failed during WC Registration")

                Else
                    If DoAll Then
                        ndt.TicketNumber = line.NDT_Number

                        ndt.UpdateNextDesk(wcSuccess)
                    End If

                End If
            Next

            wd.Quit()

        End If
        SetProgress(lines.Count)

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

        ElseIf tmpLine.DEP.ToLower.Contains("ask") Then
            tmpLine.Action = "Ticket"
        Else
            tmpLine.Action = "Discard"

        End If

        Try
            If tmpLine.DEP.Split(" ")(0).ToLower.Equals("reg") Or tmpLine.DEP.Split(" ")(0).ToLower.Equals("ask") Then
                tmpLine.Customer_DEP_ID = tmpLine.DEP.Split(" ")(1)
            Else
                tmpLine.Customer_DEP_ID = tmpLine.DEP.Split(" ")(0)
            End If

        Catch
            tmpLine.Customer_DEP_ID = 0
        End Try

        'if there is a clause to "only" register x then don't try automated submission

        If tmpLine.Action.Equals("Reg", ThisAddIn.ignoreCase) Then
            If tmpLine.DEP.ToLower.Contains("only") Then
                tmpLine.Action = "Only"
            End If
        End If

        'if it includes a fake serial then modify later behaviour

        If tmpLine.Action.Equals("Reg") Then
            If FakeSerials(tmpLine.Serials) Then
                tmpLine.Action = "Fake Serials"
            End If
        End If

        'if the supplier is GBM or Insight GmbH Germany        discard -we can't do anything

        If tmpLine.Suppliername.ToLower.Contains("gbm") Or tmpLine.Suppliername.ToLower.Contains("insight gmbh germany") Then
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
                If debugMode Then UpdateDebugMessage("Discarded line " & i)
            End If


        Next
        Return count
    End Function
    Function FakeSerials(serials As String()) As Boolean
        For Each serial In serials
            Try
                If serial IsNot Nothing AndAlso serial.ToLower.StartsWith("po") Then
                    If debugMode Then UpdateDebugMessage("Fake Serial Found: " & serial)
                    Return True
                End If
            Catch ex As Exception
                UpdateDebugMessage("serial exception")
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
            Globals.Ribbons.Ribbon1.EnableButtons()
            Try
                Call SendTimingFile()
            Catch
            End Try

            Me.Close()
        End If
    End Sub
    Delegate Sub CloseCallBack()

    Private Sub UpdateStatus(message As String)
        Globals.ThisAddIn.Application.StatusBar = message
        If debugMode Then UpdateDebugMessage(message)
        Call SetText(message)
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Interrupt = True
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



    Private Sub SetClipText(ByVal [text] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetClipTextCallback(AddressOf SetClipText)
            Me.Invoke(d, New Object() {[text]})
        Else

            My.Computer.Clipboard.SetText([text])
        End If
    End Sub
    Delegate Sub SetClipTextCallback(ByVal [text] As String)

    Private Sub SendTimingFile()
        Dim AppOutlook As New Outlook.Application
        Dim timingEmail As Outlook.MailItem
        timingEmail = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)

        timingEmail.Attachments.Add(timingFile, Outlook.OlAttachmentType.olByValue, 1, timingFile)
        timingEmail.To = "martin.klefas@insight.com"
        timingEmail.Subject = "Timing File"
        timingEmail.Send()

    End Sub
End Class