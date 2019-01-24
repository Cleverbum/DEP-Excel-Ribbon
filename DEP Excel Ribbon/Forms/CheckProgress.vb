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
                createFrm.HighlightError(oXlWs.Cells(i, 1))
                errorCount += 1

            End Try

            i += 1
            If debugMode Then UpdateDebugMessage("Read in line " & i)
        End While
        Dim total As Long = lines.LongCount

        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket
        Dim browser As Chrome.ChromeDriver = createFrm.DoWCLogin()



    End Sub
End Class