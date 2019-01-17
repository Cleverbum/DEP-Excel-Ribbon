Imports System.Diagnostics
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles CreateNew.Click
        If Not Globals.ThisAddIn.RegistrationRunning Then
            Dim frm As New CreateNew(True)
            frm.Show()
        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles CloseStale.Click
        Dim frm As New CloseStale
        frm.Show()
    End Sub



    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles FindIgnored.Click
        Dim frm As New FindIgnored
        frm.Show()
    End Sub





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

    Private Sub Button4_Click_1(sender As Object, e As RibbonControlEventArgs)
        Dim options As New Chrome.ChromeOptions
        Dim service As ChromeDriverService = ChromeDriverService.CreateDefaultService


        MsgBox("trying to open a chrome window")


        Try
            Dim wd As New Chrome.ChromeDriver(service, options)
            wd.Navigate.GoToUrl("http://www.google.com")
        Catch
            MsgBox("oh dear, that didn't work at all!")
        End Try

    End Sub

    Private Sub WriteMails_Click(sender As Object, e As RibbonControlEventArgs) Handles WriteMails.Click
        Dim frm As New PivotMail
        frm.Show()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As RibbonControlEventArgs) Handles TDOnly.Click
        Dim frm As New CreateNew(False)
        frm.Show()
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim oXlWb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        Dim oXlWs As Excel.Worksheet = oXlWb.ActiveSheet

        Dim snglLine As ClsDepLine
        Dim i As Integer
        Dim frm As New CreateNew
        i = 2
        snglLine = frm.ReadExcelLine(oXlWs, i)
        Dim wd As Chrome.ChromeDriver = frm.DoWCLogin()

        Dim success As Boolean = frm.DoOneWC_DEP(snglLine, wd)
    End Sub


End Class
