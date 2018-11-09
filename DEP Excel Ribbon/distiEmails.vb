Imports Excel = Microsoft.Office.Interop.Excel

Imports Microsoft.Office.Interop

Public Class clsDistiEmail

    Public Function generateMail(DepInfo As clsDepLine) As Outlook.MailItem
        Dim AppOutlook As New Outlook.Application
        Dim distiEmail As Outlook.MailItem
        distiEmail = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)

        If DepInfo.Suppliername.StartsWith("Ingram", Globals.ThisAddIn.ignoreCase) Then
            Dim templateFile As String
            templateFile = createIngramSpreadsheet(DepInfo)
            distiEmail.Attachments.Add(templateFile)
            distiEmail.Body = "Please could you register the attached devices?"
            distiEmail.To = "MobilityAppleDEPEMEA@ingrammicro.com"
            distiEmail.Subject = "Please can you register the attached devices"
            My.Computer.FileSystem.DeleteFile(templateFile)
        ElseIf DepInfo.Suppliername.StartsWith("Westcoast", Globals.ThisAddIn.ignoreCase) Then
            distiEmail.HTMLBody = westCoastBody(DepInfo)
            distiEmail.To = "dep@westcoast.co.uk"
            distiEmail.Subject = "Please can you register the below devices"
        Else
            'probably Techdata - at this stage do nothing
        End If



        Return distiEmail
    End Function

    Function createIngramSpreadsheet(DEPInfo As clsDepLine) As String
        Dim templateFile As String = Environ("TEMP") & "\Ingram DEP Enrol.xlsx"

        My.Computer.FileSystem.WriteAllBytes(templateFile, My.Resources.Ingram_DEP_Enrol, False)

        Dim oXlApp As Excel.Application = Nothing
        Dim oWb As Excel.Workbook, oXlSheet As Excel.Worksheet

        oXlApp = New Excel.Application
        oXlApp.Visible = True
        oWb = oXlApp.Workbooks.Open(templateFile)
        oXlSheet = oWb.Worksheets("Please fill out")
        With oXlSheet
            .Range("B11").Value = DEPInfo.Sales_ID
            .Range("B12").Value = DEPInfo.Invoice_Date
            .Range("B13").Value = DEPInfo.Invoice_Date
            .Range("B16").Value = DEPInfo.Company
            .Range("B17").Value = DEPInfo.DEP
            .Range("B18").Value = DEPInfo.Customer_PO

        End With
        For i = 0 To 9
            oWb.Worksheets("Serial #").Range("A" & i + 2).value = DEPInfo.Serials(i)
        Next

        oWb.Save()
        oXlApp.Quit()

        Return templateFile
    End Function
    Function westCoastBody(depLine As clsDepLine) As String
        Dim tmp As String

        tmp = "Hi all, <br> Could you please log the below devices for us:<br>" & vbCrLf
        tmp &= "<table><tr><td colspan=""2"">Reseller Information</td></tr>" & vbCrLf

        tmp &= tableLineHTML("Our DEP ID", "3960F70")
        tmp &= tableLineHTML("Our Sales Order number", depLine.Sales_ID.ToString)

        tmp &= "</table>" & vbCrLf & "<br>" & vbCrLf & "<table><tr><td colspan=""2"">End-User Information</td></tr>" & vbCrLf
        tmp &= tableLineHTML("End-User Organization Name", depLine.Company)
        tmp &= tableLineHTML("End-User DEP ID", depLine.DEP)
        tmp &= tableLineHTML("End-User PO#", depLine.Customer_PO)

        tmp &= "</table>" & vbCrLf & "<br>" & vbCrLf & "<table><tr><td colspan=""2"">Device Serial Numbers</td></tr>" & vbCrLf
        For i = 0 To 9
            tmp &= tableLineHTML("Serial #" & i, depLine.Serials(i))
        Next

        tmp &= "</table>"

        Return tmp

    End Function

    Function tableLineHTML(cellOne As String, CellTwo As String) As String
        Return "<tr>" & vbTab & "<td>" & vbCrLf & vbTab & vbTab & cellOne & vbCrLf & vbTab &
            "</td>" & vbCrLf & vbTab & "<td>" & vbCrLf & vbTab & vbTab & CellTwo &
            vbCrLf & vbTab & "</td>" & vbCrLf & "</tr>"
    End Function
End Class
