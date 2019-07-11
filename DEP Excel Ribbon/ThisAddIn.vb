

Public Class ThisAddIn

    Public Const ccList As String = "Chapman, Duncan <Duncan.Chapman@insight.com>; Ings, Jenni <Jenni.Ings@insight.com>"


    Public Const ignoreCase As StringComparison = StringComparison.CurrentCultureIgnoreCase

    Public RegistrationRunning As Boolean = False

    Public Const timingFile As String = "\\insight.com\root\Shared\Sales\public sector\Martin Klefas\Data\NextDesk Metrics\internaltimingfile.csv"

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub


    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub



End Class
