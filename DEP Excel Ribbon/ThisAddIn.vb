Imports System.Diagnostics
Imports outlook = Microsoft.Office.Interop.Outlook

Public Class ThisAddIn

    Public distiEmailMessage = "As attached, an email was sent to distribution to ask them
                to register these devices to the appropriate DEP ID. They will confirm by email when this is done,
                at which point this ticket will be updated again."

    Public ignoreCase As StringComparison = StringComparison.CurrentCultureIgnoreCase

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

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
