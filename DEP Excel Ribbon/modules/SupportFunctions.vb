Imports System.Diagnostics
Imports System.Net
Imports outlook = Microsoft.Office.Interop.Outlook

Partial Class ThisAddIn

    Public Sub HighlightError(FirstSerial As String)
        Dim tsheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet
        Dim tCell As Excel.Range = tsheet.Cells.Find(FirstSerial)

        With tCell.EntireRow.Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
    End Sub
    Public Function OnIntranet() As Boolean
        Try
            Dim hostentry As IPHostEntry = Dns.GetHostEntry("nextdesk")
            Return hostentry.HostName.ToLower.Contains("insight")
        Catch ex As Exception
            Return False
        End Try
    End Function


    Shared Sub PlaceHolder()
        MsgBox("This part of the code isn't quite ready yet")
    End Sub

    Function FindAlias(ByVal emailAddress As String) As String
        Dim AppOutlook As New outlook.Application
        Dim outlookNameSpace As outlook.NameSpace = AppOutlook.GetNamespace("MAPI")
        Dim myAddressList As outlook.AddressList = outlookNameSpace.GetGlobalAddressList


        Dim objAEntry As outlook.AddressEntry

        emailAddress = emailAddress.Replace("@uk.insight.com", "@insight.com")

        'below are corrections between iCare email addresses and Nextdesk Email addresses
        emailAddress = emailAddress.Replace("Scott.Waggstaff@insight.com", "Scott.Wagstaff@insight.com")



        'do final lookup of email to "recipient"
        objAEntry = AppOutlook.Session.CreateRecipient(emailAddress).AddressEntry

        'get alias of "recipient"

        Try
            FindAlias = objAEntry.GetExchangeUser.Alias
        Catch
            FindAlias = "NN"
        End Try

    End Function
End Class

