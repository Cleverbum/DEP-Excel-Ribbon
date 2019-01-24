Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class CheckProgress

    Public interrupt As Boolean = False
    Private debugMode As Boolean
    Private showDebugInfo As Boolean
    Private debugFrm As DebugForm
    Public timeEstimate As TimeEstimator

    Public Sub New(showDebugInfo As Boolean)
        InitializeComponent()
        Me.showDebugInfo = showDebugInfo
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        interrupt = True
        Me.Close()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.ThisAddIn.RegistrationRunning = True
        Me.Label1.Text = "Reading in the Excel file"
        Me.Label2.Text = "Calculating Duration Estimate"
        'parallel
        BackgroundWorker1.RunWorkerAsync()

        If debugMode Then
            debugFrm = New DebugForm
            updateDebugMessage("Starting processes.")
        End If

        'testing version (no threading for debugging)
        'Call 
    End Sub

    Private Sub UpdateDebugMessage(MessageString As String)
        'write it to the Visual Studio Debugger for luck
        Debug.WriteLine(MessageString)

        'write it to the form if it's active
        If debugMode Then debugFrm.debugText &= MessageString & vbCrLf
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim i, errorcount As Integer
        Dim createFrm As New CreateNew
        Dim lines As New List(Of ClsDepLine)

        i = 2

        While oXlWs.Cells(i, 1).value <> ""
            Try
                lines.Add(createFrm.ReadExcelLine(oXlWs, i))
            Catch
                Globals.ThisAddIn.HighlightError(oXlWs.Cells(i, 1))
                errorcount += 1

            End Try

            i += 1
            If debugMode Then UpdateDebugMessage("Read in line " & i)
        End While

        DiscardIgnore(lines)


        Call SetProgressMax(lines.LongCount)


        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

        Dim success As Boolean
        Dim wcbrowser As Chrome.ChromeDriver = createFrm.DoWCLogin()
        Dim tdbrowser As Chrome.ChromeDriver = createFrm.DoTDLogin()

        For Each OneLine As ClsDepLine In lines
            Dim tSerial As String = OneLine.Serials(0)
            If OneLine.Suppliername.ToLower.Contains("westcoast") Then
                success = CheckOneWC(wcbrowser, tSerial)
            ElseIf OneLine.Suppliername.ToLower.Contains("tech data") Then
                success = CheckOneTD(tdbrowser, tSerial)
            End If

            ndt.TicketNumber = ndt.FindTicket(0, tSerial)
            If success Then
                ndt.UpdateNextDesk(SuccessfulRegMsg)
                ndt.CloseTicket(CloseMsg)
            Else
                ndt.UpdateNextDesk(FailedRegMsg)
                Globals.ThisAddIn.HighlightError(tSerial)
                errorcount += 1
            End If
        Next


        Closeme()

        If errorcount = 0 Then
            MsgBox("All tasks completed without error")
        Else
            MsgBox("There were " & errorcount & " errors during this process. They have been highlighted in red.", vbCritical)
        End If





    End Sub

    Private Function CheckOneWC(wd As Chrome.ChromeDriver, firstSerial As String) As Boolean
        wd.Navigate.GoToUrl("https://www.westcoast.co.uk/StartApp.aspx?name=dep&update=0&delete=0&mixbasket=0&pricecheck=1")

        wd.Navigate.GoToUrl("http://dep.westcoast.co.uk/searchbyserial.php")

        wd.FindElementByName("ordsearch").SendKeys(firstSerial)

        wd.FindElementByClassName("searchbyord").Click()

        Return wd.FindElementByClassName("tableHeightSpacing").Text.Contains("Complete")

    End Function


    Function CheckOneTD(wd As Chrome.ChromeDriver, firstSerial As String) As Boolean
        wd.Navigate.GoToUrl("https://intouch.techdata.com/intouch/Home.aspx")
        wd.FindElementByLinkText("Apple DEP").Click()

        wd.SwitchTo.Frame("ctl00_CPH_iframeCat")
        wd.FindElementById("ancTran").Click()

        Threading.Thread.Sleep(TimeSpan.FromSeconds(8))

        wd.FindElementById("txtDepid").SendKeys(firstSerial)
        wd.FindElementByClassName("btn_serch").Click()
        Dim elements = wd.FindElementsByClassName("webgrid-row-style")

        If elements.Count > 1 Then
            Return (MsgBox("Was the most recent registration shown successful?", vbYesNo) = vbYes)
        Else
            Return elements(0).Text.ToLower.Contains("complete")
        End If

        Return True
    End Function

    Sub DiscardIgnore(rawlines As List(Of ClsDepLine))
        For i = rawlines.Count - 1 To 0 Step -1
            If rawlines(i).Action.Equals("Reg", ThisAddIn.ignoreCase) And (rawlines(i).Suppliername.ToLower.Contains("westcoast") Or rawlines(i).Suppliername.ToLower.Contains("tech data")) Then
                rawlines.RemoveAt(i)

                If debugMode Then UpdateDebugMessage("Discarded line " & i)
            End If


        Next

    End Sub




    Private Sub Closeme()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New CloseCallBack(AddressOf Closeme)
            Me.Invoke(d, New Object() {})
        Else
            Globals.Ribbons.Ribbon1.EnableButtons()
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



End Class