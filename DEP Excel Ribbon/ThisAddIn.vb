Imports System.Diagnostics
Imports System.Net
Imports outlook = Microsoft.Office.Interop.Outlook

Public Class ThisAddIn

    Public Const ccList As String = "Chapman, Duncan <Duncan.Chapman@insight.com>; Ings, Jenni <Jenni.Ings@insight.com>"


    Public Const ignoreCase As StringComparison = StringComparison.CurrentCultureIgnoreCase

    Public RegistrationRunning As Boolean = False

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Public Function OnIntranet() As Boolean
        Try
            Dim hostentry As IPHostEntry = Dns.GetHostEntry("nextdesk")
            Return hostentry.HostName.ToLower.Contains("insight")
        Catch ex As Exception
            Return False
        End Try
    End Function



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
